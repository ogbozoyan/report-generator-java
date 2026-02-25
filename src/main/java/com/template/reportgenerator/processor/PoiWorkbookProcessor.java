package com.template.reportgenerator.processor;

import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.MissingValuePolicy;
import com.template.reportgenerator.dto.ResolvedText;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.exception.TemplateDataBindingException;
import com.template.reportgenerator.exception.TemplateReadWriteException;
import com.template.reportgenerator.util.TemplateScanner;
import com.template.reportgenerator.util.TokenResolver;
import com.template.reportgenerator.util.ValueWriter;
import com.template.reportgenerator.util.WarningCollector;
import lombok.extern.slf4j.Slf4j;
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
@Slf4j
public class PoiWorkbookProcessor implements WorkbookProcessor {

    private final Workbook workbook;

    public PoiWorkbookProcessor(byte[] bytes) {
        log.info("Initializing POI workbook processor with {} bytes", bytes.length);
        try {
            this.workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes));
            log.info("Successfully loaded POI workbook with {} sheets", workbook.getNumberOfSheets());
        } catch (Exception e) {
            log.error("Failed to read XLS/XLSX template: {} bytes", bytes.length, e);
            throw new TemplateReadWriteException("Failed to read XLS/XLSX template", e);
        }
    }

    @Override
    public TemplateScanResult scan() {
        log.info("Starting POI template scanning");
        TemplateScanResult result = TemplateScanner.scanPoi(workbook);
        log.info("POI template scan completed - found tokens across sheets {}", result);
        return result;
    }

    @Override
    public void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector) {
        Map<String, Object> context = scalars == null ? Map.of() : scalars;

        log.info("Applying scalar tokens to {} sheets with context size: {}",
            workbook.getNumberOfSheets(), context.size());
        log.info("Processing options: missingValuePolicy={}, zoneId={}, recalculateFormulas={}",
            options.missingValuePolicy(), options.zoneId(), options.recalculateFormulas());

        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            String sheetName = sheet.getSheetName();
            List<TableAnchor> anchors = new ArrayList<>();
            int processedCells = 0;
            int tableTokensFound = 0;
            int scalarTokensApplied = 0;

            log.info("Processing sheet '{}' ({}/{}): last row = {}, last cell = {}",
                sheetName, s + 1, workbook.getNumberOfSheets(),
                sheet.getLastRowNum(),
                sheet.getLastRowNum() >= 0 ? sheet.getRow(sheet.getLastRowNum()).getLastCellNum() : 0);

            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }

                short lastCellNum = row.getLastCellNum();
                if (lastCellNum <= 0) {
                    continue;
                }

                log.info("Processing row {} with {} cells", rowIndex, lastCellNum);

                for (int colIndex = 0; colIndex < lastCellNum; colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        continue;
                    }

                    String location = cellLocation(sheet, rowIndex, colIndex);
                    processedCells++;

                    if (cell.getCellType() == CellType.FORMULA) {
                        String formula = cell.getCellFormula();
                        if (TokenResolver.hasTokens(formula)) {
                            log.info("Skipping formula token at {}: {}", location, formula);
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

                    log.info("Found token in cell {}: {}", location, original);

                    String exactToken = TokenResolver.getExactToken(original);
                    if (exactToken != null && !TokenResolver.isItemOrIndexToken(exactToken)) {
                        Object resolved = TokenResolver.resolvePath(context, exactToken);
                        if (TokenResolver.isTableValue(resolved)) {
                            List<Map<String, Object>> rows = TokenResolver.toTableRows(resolved);
                            if (rows == null) {
                                log.warn("Invalid table structure for token {} at {}", exactToken, location);
                                warningCollector.add(
                                    "TABLE_TOKEN_INVALID",
                                    "Table token has invalid structure: " + exactToken,
                                    location
                                );
                            } else {
                                log.info("Found table token {} with {} rows at {}",
                                    exactToken, rows.size(), location);
                                tableTokensFound++;
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
                    scalarTokensApplied++;
                }
            }

            anchors.sort(Comparator.comparingInt(TableAnchor::rowIndex).reversed()
                .thenComparing(Comparator.comparingInt(TableAnchor::colIndex).reversed()));

            log.info("Sheet '{}': processed {} cells, found {} table tokens, applied {} scalar tokens, inserting {} tables ", sheetName, processedCells, tableTokensFound,
                scalarTokensApplied, anchors.size());

            for (TableAnchor anchor : anchors) {
                insertTableAtAnchor(sheet, anchor, options, warningCollector);
            }
        }
        log.info("Completed scalar token application across all sheets");
    }

    @Override
    public void recalculateFormulas(GenerateOptions options) {
        if (!options.recalculateFormulas()) {
            log.info("Formula recalculation skipped as per options");
            return;
        }
        log.info("Recalculating all formulas in workbook");
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
        log.info("Formula recalculation completed");
    }

    @Override
    public byte[] serialize() {
        log.info("Serializing POI workbook");
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            byte[] result = outputStream.toByteArray();
            log.info("Successfully serialized POI workbook: {} bytes", result.length);
            return result;
        } catch (Exception e) {
            log.error("Failed to serialize XLS/XLSX document", e);
            throw new TemplateReadWriteException("Failed to serialize XLS/XLSX document", e);
        }
    }

    @Override
    public void close() {
        log.info("Closing POI workbook processor");
        try {
            workbook.close();
            log.info("POI workbook closed successfully");
        } catch (Exception e) {
            log.warn("Error closing POI workbook", e);
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
                log.info("Handling missing exact token {} at {}", exactToken, location);
                handleMissingExactToken(cell, exactToken, policy, warningCollector, location, options);
                return;
            }
            log.info("Writing resolved value for token {} at {}: {}", exactToken, location, resolved);
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
            log.info("Updating cell {} with resolved text: {} -> {}",
                location, original, resolvedText.value());
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

        log.info("Inserting table '{}' at {} with {} rows", anchor.token(), location, rows.size());

        if (rows.isEmpty()) {
            log.warn("Empty table token {} at {}", anchor.token(), location);
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor.baselineStyle());
            ValueWriter.writePoiValue(anchorCell, null, options.zoneId());
            return;
        }

        List<String> columns = buildColumnOrder(rows);
        if (columns.isEmpty()) {
            log.warn("Table token with no columns {} at {}", anchor.token(), location);
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor.baselineStyle());
            ValueWriter.writePoiValue(anchorCell, null, options.zoneId());
            return;
        }

        log.info("Table '{}' structure: {} columns [{}], {} rows",
            anchor.token(), columns.size(), String.join(",", columns), rows.size());

        int dataRowCount = rows.size();
        if (dataRowCount > 0 && anchor.rowIndex() + 1 <= sheet.getLastRowNum()) {
            log.info("Shifting {} rows starting from row {}", dataRowCount, anchor.rowIndex() + 1);
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
        log.info("Completed table insertion for '{}' at {}", anchor.token(), location);
    }

    private void handleMissingExactToken(
        Cell cell,
        String token,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location,
        GenerateOptions options
    ) {
        log.info("Handling missing token '{}' with policy {} at {}", token, policy, location);
        switch (policy) {
            case EMPTY_AND_LOG -> {
                log.info("Setting empty value for missing token {} at {}", token, location);
                warningCollector.add("MISSING_TOKEN", "Token not found: " + token, location);
                ValueWriter.writePoiValue(cell, null, options.zoneId());
            }
            case LEAVE_TOKEN -> {
                log.info("Leaving original token {} unchanged at {}", token, location);
                // no-op
            }
            case FAIL_FAST -> {
                log.error("Failing fast for missing token {} at {}", token, location);
                throw new TemplateDataBindingException("Token not found: " + token + " at " + location);
            }
        }
    }

    private void autoResizeTableColumns(
        Sheet sheet,
        int startColumnIndex,
        List<String> columns,
        List<Map<String, Object>> rows
    ) {
        log.info("Auto-resizing {} columns starting from index {}", columns.size(), startColumnIndex);
        for (int c = 0; c < columns.size(); c++) {
            String column = columns.get(c);
            int maxLength = column.length();
            for (Map<String, Object> row : rows) {
                maxLength = Math.max(maxLength, stringifyLength(row.get(column)));
            }

            int desiredWidth = calculateDesiredWidth(maxLength);
            int targetColumn = startColumnIndex + c;
            int currentWidth = sheet.getColumnWidth(targetColumn);

            log.info("Column '{}': max content length={}, current width={}, desired width={}",
                column, maxLength, currentWidth, desiredWidth);

            if (currentWidth < desiredWidth) {
                log.info("Resizing column '{}' from {} to {}", column, currentWidth, desiredWidth);
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
            log.info("Creating new row at index {}", rowIndex);
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    private Cell getOrCreateCell(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            log.info("Creating new cell at column {}", colIndex);
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
