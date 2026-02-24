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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

/**
 * Spreadsheet processor for XLS/XLSX formats based on Apache POI.
 * <p>
 * Legacy TABLE/COL DSL expansion is intentionally disabled. Table insertion is
 * based on exact-placeholder tokens where token value is {@code List<Map<...>>}.
 */
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
        Map<String, Object> context = scalars == null ? Map.of() : scalars;

        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            List<TableAnchor> anchors = new ArrayList<>();

            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }

                short lastCellNum = row.getLastCellNum();
                if (lastCellNum <= 0) {
                    continue;
                }

                for (int colIndex = 0; colIndex < lastCellNum; colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        continue;
                    }

                    String location = cellLocation(sheet, rowIndex, colIndex);

                    if (cell.getCellType() == CellType.FORMULA) {
                        String formula = cell.getCellFormula();
                        if (TokenResolver.hasTokens(formula)) {
                            warningCollector.add(
                                "FORMULA_TOKEN_SKIPPED",
                                "Formula contains token and was not modified",
                                location
                            );
                        }
                        continue;
                    }

                    if (cell.getCellType() != CellType.STRING && cell.getCellType() != CellType.BLANK) {
                        continue;
                    }

                    String original = cell.getCellType() == CellType.BLANK ? "" : cell.getStringCellValue();
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
                                anchors.add(new TableAnchor(
                                    rowIndex,
                                    colIndex,
                                    exactToken,
                                    rows,
                                    cell.getCellStyle(),
                                    row.getHeight()
                                ));
                            }
                            continue;
                        }
                    }

                    applyTokenToCell(
                        cell,
                        context,
                        options.missingValuePolicy(),
                        options,
                        warningCollector,
                        location
                    );
                }
            }

            anchors.sort(Comparator.comparingInt(TableAnchor::rowIndex).reversed()
                .thenComparing(Comparator.comparingInt(TableAnchor::colIndex).reversed()));

            for (TableAnchor anchor : anchors) {
                insertTableAtAnchor(sheet, anchor, options, warningCollector);
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

    private void applyTokenToCell(
        Cell cell,
        Map<String, Object> context,
        MissingValuePolicy policy,
        GenerateOptions options,
        WarningCollector warningCollector,
        String location
    ) {
        String original = cell.getCellType() == CellType.BLANK ? "" : cell.getStringCellValue();
        String exactToken = TokenResolver.getExactToken(original);

        if (exactToken != null && !TokenResolver.isItemOrIndexToken(exactToken)) {
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
            false
        );

        if (resolvedText.changed()) {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(resolvedText.value());
        }
    }

    private void insertTableAtAnchor(
        Sheet sheet,
        TableAnchor anchor,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        String location = cellLocation(sheet, anchor.rowIndex(), anchor.colIndex());
        List<Map<String, Object>> rows = anchor.rows();
        Row anchorRow = getOrCreateRow(sheet, anchor.rowIndex());
        Cell anchorCell = getOrCreateCell(anchorRow, anchor.colIndex());

        if (rows.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor.baselineStyle());
            ValueWriter.writePoiValue(anchorCell, null, options.zoneId());
            return;
        }

        List<String> columns = buildColumnOrder(rows);
        if (columns.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor.baselineStyle());
            ValueWriter.writePoiValue(anchorCell, null, options.zoneId());
            return;
        }

        int dataRowCount = rows.size();
        if (dataRowCount > 0 && anchor.rowIndex() + 1 <= sheet.getLastRowNum()) {
            sheet.shiftRows(anchor.rowIndex() + 1, sheet.getLastRowNum(), dataRowCount, true, false);
        }

        Row headerRow = getOrCreateRow(sheet, anchor.rowIndex());
        headerRow.setHeight(anchor.baselineRowHeight());
        for (int c = 0; c < columns.size(); c++) {
            Cell cell = getOrCreateCell(headerRow, anchor.colIndex() + c);
            applyBaselineStyle(cell, anchor.baselineStyle());
            cell.setCellType(CellType.STRING);
            cell.setCellValue(columns.get(c));
        }

        for (int r = 0; r < rows.size(); r++) {
            Map<String, Object> values = rows.get(r);
            Row dataRow = getOrCreateRow(sheet, anchor.rowIndex() + 1 + r);
            dataRow.setHeight(anchor.baselineRowHeight());

            for (int c = 0; c < columns.size(); c++) {
                String column = columns.get(c);
                Cell cell = getOrCreateCell(dataRow, anchor.colIndex() + c);
                applyBaselineStyle(cell, anchor.baselineStyle());
                ValueWriter.writePoiValue(cell, values.get(column), options.zoneId());
            }
        }

        autoResizeTableColumns(sheet, anchor.colIndex(), columns, rows);
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

    private void autoResizeTableColumns(
        Sheet sheet,
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

            int desiredWidth = calculateDesiredWidth(maxLength);
            int targetColumn = startColumnIndex + c;
            if (desiredWidth > sheet.getColumnWidth(targetColumn)) {
                sheet.setColumnWidth(targetColumn, desiredWidth);
            }
        }
    }

    private int calculateDesiredWidth(int maxLength) {
        int width = (maxLength + 2) * 256;
        int min = 8 * 256;
        int max = 100 * 256;
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

    private void applyBaselineStyle(Cell cell, CellStyle baselineStyle) {
        if (baselineStyle != null) {
            cell.setCellStyle(baselineStyle);
        }
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
            cell = row.createCell(colIndex);
        }
        return cell;
    }

    private String cellLocation(Sheet sheet, int rowIndex, int colIndex) {
        return sheet.getSheetName() + "!R" + (rowIndex + 1) + "C" + (colIndex + 1);
    }

    private record TableAnchor(
        int rowIndex,
        int colIndex,
        String token,
        List<Map<String, Object>> rows,
        CellStyle baselineStyle,
        short baselineRowHeight
    ) {
    }
}
