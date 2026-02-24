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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PoiWorkbookProcessor implements WorkbookProcessor {

    private final Workbook workbook;

    public PoiWorkbookProcessor(byte[] bytes) {
        try {
            this.workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes));
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to read XLS/XLSX template", e);
        }
    }

    @Override
    public TemplateScanResult scan() {
        return TemplateScanner.scanPoi(workbook);
    }

    @Override
    public void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector) {
        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    applyTokenToCell(
                        cell,
                        scalars,
                        options.missingValuePolicy(),
                        options,
                        warningCollector,
                        false,
                        sheet.getSheetName() + "!R" + (row.getRowNum() + 1) + "C" + (cell.getColumnIndex() + 1)
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
            Sheet sheet = workbook.getSheetAt(block.sheetIndex());
            clearCell(sheet, block.startRow(), block.startCol());
            clearCell(sheet, block.endRow(), block.endCol());
        }
    }

    @Override
    public void recalculateFormulas(GenerateOptions options) {
        if (!options.recalculateFormulas()) {
            return;
        }
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
    }

    @Override
    public byte[] serialize() {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to serialize XLS/XLSX document", e);
        }
    }

    @Override
    public void close() {
        try {
            workbook.close();
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
        Sheet sheet = workbook.getSheetAt(block.sheetIndex());

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
            removeRows(sheet, innerStartRow, innerEndRow);
            return;
        }

        for (int group = 1; group < items.size(); group++) {
            int insertAt = innerEndRow + 1 + (group - 1) * templateRowCount;
            if (insertAt <= sheet.getLastRowNum()) {
                sheet.shiftRows(insertAt, sheet.getLastRowNum(), templateRowCount, true, false);
            }
            cloneRowRange(sheet, innerStartRow, innerEndRow, insertAt, innerStartCol, innerEndCol);
        }

        for (int group = 0; group < items.size(); group++) {
            Map<String, Object> context = buildContext(data.scalars(), items.get(group), group);
            for (int rowOffset = 0; rowOffset < templateRowCount; rowOffset++) {
                int rowIndex = innerStartRow + group * templateRowCount + rowOffset;
                Row row = getOrCreateRow(sheet, rowIndex);
                for (int col = innerStartCol; col <= innerEndCol; col++) {
                    Cell cell = getOrCreateCell(row, col);
                    applyTokenToCell(
                        cell,
                        context,
                        options.missingValuePolicy(),
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
        Sheet sheet = workbook.getSheetAt(block.sheetIndex());

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
            deleteColumns(sheet, innerStartCol, templateColCount);
            return;
        }

        for (int group = 1; group < items.size(); group++) {
            int insertAt = innerEndCol + 1 + (group - 1) * templateColCount;
            int maxCol = getMaxColumn(sheet);
            if (insertAt <= maxCol) {
                shiftColumns(sheet, insertAt, maxCol, templateColCount);
            }
            cloneColumnRange(sheet, innerStartCol, innerEndCol, insertAt, innerStartRow, innerEndRow);
        }

        for (int group = 0; group < items.size(); group++) {
            Map<String, Object> context = buildContext(data.scalars(), items.get(group), group);
            for (int rowIndex = innerStartRow; rowIndex <= innerEndRow; rowIndex++) {
                Row row = getOrCreateRow(sheet, rowIndex);
                for (int colOffset = 0; colOffset < templateColCount; colOffset++) {
                    int colIndex = innerStartCol + group * templateColCount + colOffset;
                    Cell cell = getOrCreateCell(row, colIndex);
                    applyTokenToCell(
                        cell,
                        context,
                        options.missingValuePolicy(),
                        options,
                        warningCollector,
                        true,
                        block.sheetName() + "!R" + (rowIndex + 1) + "C" + (colIndex + 1)
                    );
                }
            }
        }
    }

    private void applyTokenToCell(
        Cell cell,
        Map<String, Object> context,
        MissingValuePolicy policy,
        GenerateOptions options,
        WarningCollector warningCollector,
        boolean allowItemTokens,
        String location
    ) {
        if (cell.getCellType() == CellType.FORMULA) {
            String formula = cell.getCellFormula();
            if (TokenResolver.hasTokens(formula)) {
                warningCollector.add(
                    "FORMULA_TOKEN_SKIPPED",
                    "Formula contains token and was not modified",
                    location
                );
            }
            return;
        }

        if (cell.getCellType() != CellType.STRING && cell.getCellType() != CellType.BLANK) {
            return;
        }

        String original = cell.getCellType() == CellType.BLANK ? "" : cell.getStringCellValue();
        if (original == null || original.isEmpty()) {
            return;
        }

        String exactToken = TokenResolver.getExactToken(original);
        if (exactToken != null && (allowItemTokens || !TokenResolver.isItemOrIndexToken(exactToken))) {
            Object resolved = TokenResolver.resolvePath(context, exactToken);
            if (resolved == null) {
                handleMissingExactToken(cell, exactToken, policy, warningCollector, location, options);
                return;
            }
            ValueWriter.writePoiValue(cell, resolved, options.zoneId());
            return;
        }

        ResolvedText resolvedText = TokenResolver.resolve(
            original,
            context,
            policy,
            warningCollector,
            location,
            allowItemTokens
        );

        if (resolvedText.changed()) {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(resolvedText.value());
        }
    }

    private void handleMissingExactToken(
        Cell cell,
        String token,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location,
        GenerateOptions options
    ) {
        switch (policy) {
            case EMPTY_AND_LOG -> {
                warningCollector.add("MISSING_TOKEN", "Token not found: " + token, location);
                ValueWriter.writePoiValue(cell, null, options.zoneId());
            }
            case LEAVE_TOKEN -> {
                // no-op
            }
            case FAIL_FAST -> throw new TemplateDataBindingException("Token not found: " + token + " at " + location);
        }
    }

    private void cloneRowRange(
        Sheet sheet,
        int sourceStartRow,
        int sourceEndRow,
        int targetStartRow,
        int colStart,
        int colEnd
    ) {
        int rowCount = sourceEndRow - sourceStartRow + 1;

        for (int i = 0; i < rowCount; i++) {
            Row source = sheet.getRow(sourceStartRow + i);
            Row target = getOrCreateRow(sheet, targetStartRow + i);

            if (source == null) {
                continue;
            }
            target.setHeight(source.getHeight());

            for (int col = colStart; col <= colEnd; col++) {
                Cell sourceCell = source.getCell(col);
                if (sourceCell == null) {
                    Cell existing = target.getCell(col);
                    if (existing != null) {
                        target.removeCell(existing);
                    }
                    continue;
                }
                Cell targetCell = target.getCell(col);
                if (targetCell == null) {
                    targetCell = target.createCell(col, sourceCell.getCellType());
                }
                copyCell(sourceCell, targetCell);
            }
        }

        copyMergedRegionsByRowShift(
            sheet,
            sourceStartRow,
            sourceEndRow,
            colStart,
            colEnd,
            targetStartRow - sourceStartRow
        );
    }

    private void cloneColumnRange(
        Sheet sheet,
        int sourceStartCol,
        int sourceEndCol,
        int targetStartCol,
        int rowStart,
        int rowEnd
    ) {
        int columnCount = sourceEndCol - sourceStartCol + 1;

        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            for (int i = 0; i < columnCount; i++) {
                int sourceCol = sourceStartCol + i;
                int targetCol = targetStartCol + i;

                Cell sourceCell = row.getCell(sourceCol);
                if (sourceCell == null) {
                    continue;
                }

                Cell targetCell = row.getCell(targetCol);
                if (targetCell == null) {
                    targetCell = row.createCell(targetCol, sourceCell.getCellType());
                }
                copyCell(sourceCell, targetCell);
            }
        }

        for (int i = 0; i < columnCount; i++) {
            sheet.setColumnWidth(targetStartCol + i, sheet.getColumnWidth(sourceStartCol + i));
        }

        copyMergedRegionsByColumnShift(
            sheet,
            rowStart,
            rowEnd,
            sourceStartCol,
            sourceEndCol,
            targetStartCol - sourceStartCol
        );
    }

    private void copyMergedRegionsByRowShift(
        Sheet sheet,
        int sourceStartRow,
        int sourceEndRow,
        int sourceStartCol,
        int sourceEndCol,
        int rowShift
    ) {
        List<CellRangeAddress> rangesToAdd = new ArrayList<>();
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.getFirstRow() >= sourceStartRow
                && region.getLastRow() <= sourceEndRow
                && region.getFirstColumn() >= sourceStartCol
                && region.getLastColumn() <= sourceEndCol) {
                rangesToAdd.add(new CellRangeAddress(
                    region.getFirstRow() + rowShift,
                    region.getLastRow() + rowShift,
                    region.getFirstColumn(),
                    region.getLastColumn()
                ));
            }
        }

        for (CellRangeAddress range : rangesToAdd) {
            try {
                sheet.addMergedRegion(range);
            } catch (Exception ignored) {
                // ignore duplicates / invalid combinations
            }
        }
    }

    private void copyMergedRegionsByColumnShift(
        Sheet sheet,
        int sourceStartRow,
        int sourceEndRow,
        int sourceStartCol,
        int sourceEndCol,
        int colShift
    ) {
        List<CellRangeAddress> rangesToAdd = new ArrayList<>();
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.getFirstRow() >= sourceStartRow
                && region.getLastRow() <= sourceEndRow
                && region.getFirstColumn() >= sourceStartCol
                && region.getLastColumn() <= sourceEndCol) {
                rangesToAdd.add(new CellRangeAddress(
                    region.getFirstRow(),
                    region.getLastRow(),
                    region.getFirstColumn() + colShift,
                    region.getLastColumn() + colShift
                ));
            }
        }

        for (CellRangeAddress range : rangesToAdd) {
            try {
                sheet.addMergedRegion(range);
            } catch (Exception ignored) {
                // ignore duplicates / invalid combinations
            }
        }
    }

    private void shiftColumns(Sheet sheet, int startCol, int endCol, int shiftBy) {
        if (startCol > endCol || shiftBy == 0) {
            return;
        }

        try {
            sheet.shiftColumns(startCol, endCol, shiftBy);
            return;
        } catch (UnsupportedOperationException ignored) {
            // fallback to manual shift below
        }

        if (shiftBy > 0) {
            for (Row row : sheet) {
                int last = row.getLastCellNum() - 1;
                for (int col = last; col >= startCol; col--) {
                    Cell source = row.getCell(col);
                    if (source == null) {
                        continue;
                    }
                    Cell target = row.getCell(col + shiftBy);
                    if (target != null) {
                        row.removeCell(target);
                    }
                    target = row.createCell(col + shiftBy, source.getCellType());
                    copyCell(source, target);
                    row.removeCell(source);
                }
            }
        } else {
            for (Row row : sheet) {
                int last = row.getLastCellNum() - 1;
                for (int col = startCol; col <= last; col++) {
                    Cell source = row.getCell(col);
                    if (source == null) {
                        continue;
                    }
                    int targetCol = col + shiftBy;
                    if (targetCol < 0) {
                        continue;
                    }
                    Cell target = row.getCell(targetCol);
                    if (target == null) {
                        target = row.createCell(targetCol, source.getCellType());
                    }
                    copyCell(source, target);
                    row.removeCell(source);
                }
            }
        }
    }

    private void deleteColumns(Sheet sheet, int startCol, int count) {
        if (count <= 0) {
            return;
        }
        int maxCol = getMaxColumn(sheet);
        int fromCol = startCol + count;
        if (fromCol <= maxCol) {
            shiftColumns(sheet, fromCol, maxCol, -count);
        }
    }

    private int getMaxColumn(Sheet sheet) {
        int maxCol = 0;
        for (Row row : sheet) {
            maxCol = Math.max(maxCol, row.getLastCellNum() - 1);
        }
        return maxCol;
    }

    private void removeRows(Sheet sheet, int startRow, int endRow) {
        if (startRow > endRow) {
            return;
        }

        removeMergedRegionsInRowRange(sheet, startRow, endRow);

        int rowsToRemove = endRow - startRow + 1;
        int lastRow = sheet.getLastRowNum();

        if (endRow < lastRow) {
            sheet.shiftRows(endRow + 1, lastRow, -rowsToRemove, true, false);
        }

        for (int i = lastRow; i > lastRow - rowsToRemove; i--) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    private void removeMergedRegionsInRowRange(Sheet sheet, int startRow, int endRow) {
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() >= startRow && region.getLastRow() <= endRow) {
                sheet.removeMergedRegion(i);
            }
        }
    }

    private void clearCell(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return;
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            return;
        }
        if (cell.getCellType() == CellType.FORMULA) {
            cell.setBlank();
            return;
        }
        cell.setCellType(CellType.STRING);
        cell.setCellValue("");
    }

    private void copyCell(Cell source, Cell target) {
        target.setCellStyle(source.getCellStyle());
        if (source.getHyperlink() != null) {
            target.setHyperlink(source.getHyperlink());
        }
        if (source.getCellComment() != null) {
            target.setCellComment(source.getCellComment());
        }

        switch (source.getCellType()) {
            case STRING -> target.setCellValue(source.getRichStringCellValue());
            case NUMERIC -> target.setCellValue(source.getNumericCellValue());
            case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
            case FORMULA -> target.setCellFormula(source.getCellFormula());
            case ERROR -> target.setCellErrorValue(source.getErrorCellValue());
            case BLANK, _NONE -> target.setBlank();
        }
    }

    private Map<String, Object> buildContext(Map<String, Object> scalars, Map<String, Object> item, int index) {
        Map<String, Object> context = new HashMap<>();
        context.putAll(scalars);
        context.put("item", item == null ? Map.of() : item);
        context.put("index", index);
        return context;
    }

    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    private Cell getOrCreateCell(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex, CellType.STRING);
        }
        return cell;
    }
}
