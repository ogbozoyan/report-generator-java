package com.template.reportgenerator.processor;

import com.template.reportgenerator.dto.BlockRegion;
import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.MissingValuePolicy;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.ResolvedText;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.exception.TemplateDataBindingException;
import com.template.reportgenerator.exception.TemplateReadWriteException;
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
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

/**
 * Spreadsheet processor for ODS format based on ODFDOM.
 * <p>
 * Legacy TABLE/COL DSL expansion is intentionally disabled. Table insertion is
 * based on exact-placeholder tokens where token value is {@code List<Map<...>>}.
 */
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
        Map<String, Object> context = scalars == null ? Map.of() : scalars;
        List<OdfTable> sheets = document.getTableList(false);

        for (int sheetIndex = 0; sheetIndex < sheets.size(); sheetIndex++) {
            OdfTable sheet = sheets.get(sheetIndex);
            String sheetName = sheet.getTableName() == null ? ("Sheet" + sheetIndex) : sheet.getTableName();
            int maxRows = Math.min(sheet.getRowCount(), 5000);
            int maxCols = Math.min(sheet.getColumnCount(), 512);

            List<TableAnchor> anchors = new ArrayList<>();

            for (int rowIndex = 0; rowIndex < maxRows; rowIndex++) {
                for (int colIndex = 0; colIndex < maxCols; colIndex++) {
                    OdfTableCell cell = sheet.getCellByPosition(colIndex, rowIndex);
                    String location = cellLocation(sheetName, rowIndex, colIndex);
                    String original = cell.getStringValue();

                    String formula = cell.getFormula();
                    if (TokenResolver.hasTokens(formula)) {
                        warningCollector.add(
                            "FORMULA_TOKEN_SKIPPED",
                            "Formula contains token and was not modified",
                            location
                        );
                        continue;
                    }

                    if (!TokenResolver.hasTokens(original)) {
                        continue;
                    }

                    String exactToken = TokenResolver.getExactToken(original);
                    if (exactToken != null && !TokenResolver.isItemOrIndexToken(exactToken)) {
                        Object resolved = TokenResolver.resolvePath(context, exactToken);
                        if (TokenResolver.isTableValue(resolved)) {
                            List<Map<String, Object>> rows = TokenResolver.toTableRows(resolved);
                            if (rows == null) {
                                warningCollector.add(
                                    "TABLE_TOKEN_INVALID",
                                    "Table token has invalid structure: " + exactToken,
                                    location
                                );
                            } else {
                                OdfTableRow row = sheet.getRowByIndex(rowIndex);
                                anchors.add(new TableAnchor(
                                    rowIndex,
                                    colIndex,
                                    exactToken,
                                    rows,
                                    cell.getStyleName(),
                                    cell.getHorizontalAlignment(),
                                    cell.getVerticalAlignment(),
                                    cell.isTextWrapped(),
                                    row == null ? 0 : row.getHeight(),
                                    row != null && row.isOptimalHeight()
                                ));
                            }
                            continue;
                        }
                    }

                    applyTokenToCell(
                        cell,
                        context,
                        options,
                        warningCollector,
                        location
                    );
                }
            }

            anchors.sort(Comparator.comparingInt(TableAnchor::rowIndex).reversed()
                .thenComparing(Comparator.comparingInt(TableAnchor::colIndex).reversed()));

            for (TableAnchor anchor : anchors) {
                insertTableAtAnchor(sheet, sheetName, anchor, options, warningCollector);
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
        // Legacy TABLE DSL is disabled.
    }

    @Override
    public void expandColumnBlocks(
        List<BlockRegion> columnBlocks,
        ReportData data,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        // Legacy COL DSL is disabled.
    }

    @Override
    public void clearMarkers(List<BlockRegion> blockRegions) {
        // Marker cleanup is not required in token-only mode.
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

    private void applyTokenToCell(
        OdfTableCell cell,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector,
        String location
    ) {
        String original = cell.getStringValue();
        String exactToken = TokenResolver.getExactToken(original);

        if (exactToken != null && !TokenResolver.isItemOrIndexToken(exactToken)) {
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
            false
        );

        if (resolvedText.changed()) {
            cell.setStringValue(resolvedText.value());
        }
    }

    private void insertTableAtAnchor(
        OdfTable table,
        String tableName,
        TableAnchor anchor,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        String location = cellLocation(tableName, anchor.rowIndex(), anchor.colIndex());
        List<Map<String, Object>> rows = anchor.rows();
        OdfTableCell anchorCell = table.getCellByPosition(anchor.colIndex(), anchor.rowIndex());

        if (rows.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor);
            ValueWriter.writeOdsValue(anchorCell, null, options.zoneId());
            return;
        }

        List<String> columns = buildColumnOrder(rows);
        if (columns.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor);
            ValueWriter.writeOdsValue(anchorCell, null, options.zoneId());
            return;
        }

        int dataRowCount = rows.size();
        if (dataRowCount > 0) {
            table.insertRowsBefore(anchor.rowIndex() + 1, dataRowCount);
        }

        OdfTableRow headerRow = table.getRowByIndex(anchor.rowIndex());
        applyBaselineHeight(headerRow, anchor);
        for (int c = 0; c < columns.size(); c++) {
            OdfTableCell cell = table.getCellByPosition(anchor.colIndex() + c, anchor.rowIndex());
            applyBaselineStyle(cell, anchor);
            cell.setStringValue(columns.get(c));
        }

        for (int r = 0; r < rows.size(); r++) {
            int rowIndex = anchor.rowIndex() + 1 + r;
            OdfTableRow row = table.getRowByIndex(rowIndex);
            applyBaselineHeight(row, anchor);

            Map<String, Object> values = rows.get(r);
            for (int c = 0; c < columns.size(); c++) {
                String column = columns.get(c);
                OdfTableCell cell = table.getCellByPosition(anchor.colIndex() + c, rowIndex);
                applyBaselineStyle(cell, anchor);
                ValueWriter.writeOdsValue(cell, values.get(column), options.zoneId());
            }
        }

        autoResizeTableColumns(table, anchor.colIndex(), columns, rows);
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

    private void autoResizeTableColumns(
        OdfTable table,
        int startColumnIndex,
        List<String> columns,
        List<Map<String, Object>> rows
    ) {
        for (int c = 0; c < columns.size(); c++) {
            String column = columns.get(c);
            int maxLength = column.length();
            for (Map<String, Object> row : rows) {
                maxLength = Math.max(maxLength, stringifyLength(row.get(column)));
            }

            long desiredWidth = calculateDesiredWidth(maxLength);
            OdfTableColumn target = table.getColumnByIndex(startColumnIndex + c);
            if (target.getWidth() < desiredWidth) {
                target.setWidth(desiredWidth);
            }
        }
    }

    private long calculateDesiredWidth(int maxLength) {
        long width = (long) (maxLength + 2) * 260L;
        long min = 1400L;
        long max = 12000L;
        return Math.max(min, Math.min(width, max));
    }

    private int stringifyLength(Object value) {
        return value == null ? 0 : String.valueOf(value).length();
    }

    private List<String> buildColumnOrder(List<Map<String, Object>> rows) {
        LinkedHashSet<String> ordered = new LinkedHashSet<>();
        ordered.addAll(rows.get(0).keySet());
        for (Map<String, Object> row : rows) {
            ordered.addAll(row.keySet());
        }
        return List.copyOf(ordered);
    }

    private void applyBaselineStyle(OdfTableCell cell, TableAnchor anchor) {
        if (anchor.styleName() != null) {
            cell.getOdfElement().setTableStyleNameAttribute(anchor.styleName());
        }
        if (anchor.horizontalAlignment() != null) {
            cell.setHorizontalAlignment(anchor.horizontalAlignment());
        }
        if (anchor.verticalAlignment() != null) {
            cell.setVerticalAlignment(anchor.verticalAlignment());
        }
        cell.setTextWrapped(anchor.wrapped());
    }

    private void applyBaselineHeight(OdfTableRow row, TableAnchor anchor) {
        if (row != null && anchor.rowHeight() > 0) {
            row.setHeight(anchor.rowHeight(), anchor.rowOptimalHeight());
        }
    }

    private String cellLocation(String tableName, int rowIndex, int colIndex) {
        return tableName + "!R" + (rowIndex + 1) + "C" + (colIndex + 1);
    }

    private record TableAnchor(
        int rowIndex,
        int colIndex,
        String token,
        List<Map<String, Object>> rows,
        String styleName,
        String horizontalAlignment,
        String verticalAlignment,
        boolean wrapped,
        long rowHeight,
        boolean rowOptimalHeight
    ) {
    }
}
