package com.template.reportgenerator.processor;

import com.template.reportgenerator.dto.BlockRegion;
import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.MissingValuePolicy;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.ResolvedText;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.exception.TemplateDataBindingException;
import com.template.reportgenerator.exception.TemplateReadWriteException;
import com.template.reportgenerator.exception.TemplateSyntaxException;
import com.template.reportgenerator.util.TemplateScanner;
import com.template.reportgenerator.util.TokenResolver;
import com.template.reportgenerator.util.ValueWriter;
import com.template.reportgenerator.util.WarningCollector;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.doc.table.OdfTableColumn;
import org.odftoolkit.odfdom.doc.table.OdfTableRow;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class OdsWorkbookProcessor implements WorkbookProcessor {

    private final OdfSpreadsheetDocument document;

    public OdsWorkbookProcessor(byte[] bytes) {
        try {
            this.document = OdfSpreadsheetDocument.loadDocument(new ByteArrayInputStream(bytes));
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to read ODS template", e);
        }
    }

    @Override
    public TemplateScanResult scan() {
        return TemplateScanner.scanOds(document);
    }

    @Override
    public void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector) {
        List<OdfTable> tables = document.getTableList(false);
        for (int t = 0; t < tables.size(); t++) {
            OdfTable table = tables.get(t);
            String tableName = table.getTableName() == null ? ("Sheet" + t) : table.getTableName();

            int maxRows = Math.min(table.getRowCount(), 5000);
            int maxCols = Math.min(table.getColumnCount(), 512);

            for (int row = 0; row < maxRows; row++) {
                for (int col = 0; col < maxCols; col++) {
                    OdfTableCell cell = table.getCellByPosition(col, row);
                    applyTokenToCell(
                        cell,
                        scalars,
                        options,
                        warningCollector,
                        false,
                        tableName + "!R" + (row + 1) + "C" + (col + 1)
                    );
                }
            }
        }
    }

    @Override
    public void expandTableBlocks(
        List<BlockRegion> tableBlocks,
        ReportData data,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        List<BlockRegion> sorted = tableBlocks.stream()
            .sorted(Comparator.comparingInt(BlockRegion::sheetIndex).reversed()
                .thenComparing(Comparator.comparingInt(BlockRegion::startRow).reversed()))
            .toList();

        for (BlockRegion block : sorted) {
            expandSingleTable(block, data, options, warningCollector);
        }
    }

    @Override
    public void expandColumnBlocks(
        List<BlockRegion> columnBlocks,
        ReportData data,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        List<BlockRegion> sorted = columnBlocks.stream()
            .sorted(Comparator.comparingInt(BlockRegion::sheetIndex).reversed()
                .thenComparing(Comparator.comparingInt(BlockRegion::startCol).reversed()))
            .toList();

        for (BlockRegion block : sorted) {
            expandSingleColumns(block, data, options, warningCollector);
        }
    }

    @Override
    public void clearMarkers(List<BlockRegion> blockRegions) {
        for (BlockRegion block : blockRegions) {
            OdfTable table = getTable(block.sheetIndex());
            table.getCellByPosition(block.startCol(), block.startRow()).setStringValue("");
            table.getCellByPosition(block.endCol(), block.endRow()).setStringValue("");
        }
    }

    @Override
    public void recalculateFormulas(GenerateOptions options) {
        // ODF formula recalc is delegated to office application on open.
    }

    @Override
    public byte[] serialize() {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            document.save(outputStream);
            return outputStream.toByteArray();
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to serialize ODS document", e);
        }
    }

    @Override
    public void close() {
        try {
            document.close();
        } catch (Exception ignored) {
            // no-op
        }
    }

    private void expandSingleTable(
        BlockRegion block,
        ReportData data,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        OdfTable table = getTable(block.sheetIndex());

        int innerStartRow = block.innerStartRow();
        int innerEndRow = block.innerEndRow();
        int innerStartCol = block.innerStartCol();
        int innerEndCol = block.innerEndCol();

        int templateRowCount = innerEndRow - innerStartRow + 1;
        if (templateRowCount <= 0) {
            throw new TemplateSyntaxException("TABLE block has empty internal rows: " + block.asLocation());
        }

        List<Map<String, Object>> items = data.tables().getOrDefault(block.key(), List.of());
        if (items.isEmpty()) {
            table.removeRowsByIndex(innerStartRow, templateRowCount);
            return;
        }

        for (int group = 1; group < items.size(); group++) {
            int insertAt = innerEndRow + 1 + (group - 1) * templateRowCount;
            table.insertRowsBefore(insertAt, templateRowCount);
            cloneRowRange(table, innerStartRow, innerEndRow, insertAt, innerStartCol, innerEndCol);
        }

        for (int group = 0; group < items.size(); group++) {
            Map<String, Object> context = buildContext(data.scalars(), items.get(group), group);
            for (int rowOffset = 0; rowOffset < templateRowCount; rowOffset++) {
                int rowIndex = innerStartRow + group * templateRowCount + rowOffset;
                for (int col = innerStartCol; col <= innerEndCol; col++) {
                    OdfTableCell cell = table.getCellByPosition(col, rowIndex);
                    applyTokenToCell(
                        cell,
                        context,
                        options,
                        warningCollector,
                        true,
                        block.sheetName() + "!R" + (rowIndex + 1) + "C" + (col + 1)
                    );
                }
            }
        }
    }

    private void expandSingleColumns(
        BlockRegion block,
        ReportData data,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        OdfTable table = getTable(block.sheetIndex());

        int innerStartRow = block.innerStartRow();
        int innerEndRow = block.innerEndRow();
        int innerStartCol = block.innerStartCol();
        int innerEndCol = block.innerEndCol();

        int templateColCount = innerEndCol - innerStartCol + 1;
        if (templateColCount <= 0) {
            throw new TemplateSyntaxException("COL block has empty internal columns: " + block.asLocation());
        }

        List<Map<String, Object>> items = data.columns().getOrDefault(block.key(), List.of());
        if (items.isEmpty()) {
            table.removeColumnsByIndex(innerStartCol, templateColCount);
            return;
        }

        for (int group = 1; group < items.size(); group++) {
            int insertAt = innerEndCol + 1 + (group - 1) * templateColCount;
            table.insertColumnsBefore(insertAt, templateColCount);
            cloneColumnRange(table, innerStartCol, innerEndCol, insertAt, innerStartRow, innerEndRow);
        }

        for (int group = 0; group < items.size(); group++) {
            Map<String, Object> context = buildContext(data.scalars(), items.get(group), group);
            for (int rowIndex = innerStartRow; rowIndex <= innerEndRow; rowIndex++) {
                for (int colOffset = 0; colOffset < templateColCount; colOffset++) {
                    int colIndex = innerStartCol + group * templateColCount + colOffset;
                    OdfTableCell cell = table.getCellByPosition(colIndex, rowIndex);
                    applyTokenToCell(
                        cell,
                        context,
                        options,
                        warningCollector,
                        true,
                        block.sheetName() + "!R" + (rowIndex + 1) + "C" + (colIndex + 1)
                    );
                }
            }
        }
    }

    private void cloneRowRange(
        OdfTable table,
        int sourceStartRow,
        int sourceEndRow,
        int targetStartRow,
        int colStart,
        int colEnd
    ) {
        int rowCount = sourceEndRow - sourceStartRow + 1;
        for (int i = 0; i < rowCount; i++) {
            int sourceRowIndex = sourceStartRow + i;
            int targetRowIndex = targetStartRow + i;

            OdfTableRow sourceRow = table.getRowByIndex(sourceRowIndex);
            OdfTableRow targetRow = table.getRowByIndex(targetRowIndex);

            targetRow.setHeight(sourceRow.getHeight(), sourceRow.isOptimalHeight());

            for (int col = colStart; col <= colEnd; col++) {
                copyCell(table.getCellByPosition(col, sourceRowIndex), table.getCellByPosition(col, targetRowIndex));
            }
        }
    }

    private void cloneColumnRange(
        OdfTable table,
        int sourceStartCol,
        int sourceEndCol,
        int targetStartCol,
        int rowStart,
        int rowEnd
    ) {
        int colCount = sourceEndCol - sourceStartCol + 1;

        for (int i = 0; i < colCount; i++) {
            OdfTableColumn sourceCol = table.getColumnByIndex(sourceStartCol + i);
            OdfTableColumn targetCol = table.getColumnByIndex(targetStartCol + i);
            targetCol.setWidth(sourceCol.getWidth());
            targetCol.setUseOptimalWidth(sourceCol.isOptimalWidth());
        }

        for (int row = rowStart; row <= rowEnd; row++) {
            for (int i = 0; i < colCount; i++) {
                copyCell(table.getCellByPosition(sourceStartCol + i, row), table.getCellByPosition(targetStartCol + i, row));
            }
        }
    }

    private void applyTokenToCell(
        OdfTableCell cell,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector,
        boolean allowItemTokens,
        String location
    ) {
        String formula = cell.getFormula();
        if (TokenResolver.hasTokens(formula)) {
            warningCollector.add(
                "FORMULA_TOKEN_SKIPPED",
                "Formula contains token and was not modified",
                location
            );
            return;
        }

        String original = cell.getStringValue();
        if (!TokenResolver.hasTokens(original)) {
            return;
        }

        String exactToken = TokenResolver.getExactToken(original);
        if (exactToken != null && (allowItemTokens || !TokenResolver.isItemOrIndexToken(exactToken))) {
            Object resolved = TokenResolver.resolvePath(context, exactToken);
            if (resolved == null) {
                handleMissingExactToken(cell, exactToken, options.missingValuePolicy(), warningCollector, location);
                return;
            }
            ValueWriter.writeOdsValue(cell, resolved, options.zoneId());
            return;
        }

        ResolvedText resolvedText = TokenResolver.resolve(
            original,
            context,
            options.missingValuePolicy(),
            warningCollector,
            location,
            allowItemTokens
        );

        if (resolvedText.changed()) {
            cell.setStringValue(resolvedText.value());
        }
    }

    private void handleMissingExactToken(
        OdfTableCell cell,
        String token,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location
    ) {
        switch (policy) {
            case EMPTY_AND_LOG -> {
                warningCollector.add("MISSING_TOKEN", "Token not found: " + token, location);
                cell.setStringValue("");
            }
            case LEAVE_TOKEN -> {
                // no-op
            }
            case FAIL_FAST -> throw new TemplateDataBindingException("Token not found: " + token + " at " + location);
        }
    }

    private void copyCell(OdfTableCell source, OdfTableCell target) {
        String styleName = source.getStyleName();
        if (styleName != null) {
            target.getOdfElement().setTableStyleNameAttribute(styleName);
        }

        String formula = source.getFormula();
        if (formula != null) {
            target.setFormula(formula);
        }

        target.setTextWrapped(source.isTextWrapped());
        target.setHorizontalAlignment(source.getHorizontalAlignment());
        target.setVerticalAlignment(source.getVerticalAlignment());

        String valueType = source.getValueType();
        if (valueType == null) {
            target.setStringValue(source.getStringValue());
            return;
        }

        switch (valueType) {
            case "float", "currency", "percentage" -> target.setDoubleValue(source.getDoubleValue());
            case "boolean" -> target.setBooleanValue(source.getBooleanValue());
            case "date" -> target.setDateValue(source.getDateValue());
            case "time" -> target.setTimeValue(source.getTimeValue());
            default -> target.setStringValue(source.getStringValue());
        }
    }

    private OdfTable getTable(int index) {
        List<OdfTable> tables = document.getTableList(false);
        if (index < 0 || index >= tables.size()) {
            throw new TemplateSyntaxException("Invalid table index: " + index);
        }
        return tables.get(index);
    }

    private Map<String, Object> buildContext(Map<String, Object> scalars, Map<String, Object> item, int index) {
        Map<String, Object> context = new HashMap<>(scalars);
        context.put("item", item == null ? Map.of() : item);
        context.put("index", index);
        return context;
    }
}
