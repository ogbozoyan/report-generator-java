package com.template.reportgenerator.processor;

import com.template.reportgenerator.contract.GenerateOptions;
import com.template.reportgenerator.contract.MissingValuePolicy;
import com.template.reportgenerator.contract.PoiTableAnchor;
import com.template.reportgenerator.contract.ResolvedText;
import com.template.reportgenerator.contract.TemplateScanResult;
import com.template.reportgenerator.exception.TemplateDataBindingException;
import com.template.reportgenerator.exception.TemplateReadWriteException;
import com.template.reportgenerator.util.TemplateScanner;
import com.template.reportgenerator.util.TokenResolver;
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
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Spreadsheet processor for XLS/XLSX formats based on Apache POI.
 * <p>
 */
@Slf4j
public class PoiWorkbookProcessor implements WorkbookProcessor {

    private final Workbook workbook;

    public PoiWorkbookProcessor(byte[] bytes) {
        log.info("PoiWorkbookProcessor() - start: bytesLength={}", bytes.length);
        try {
            this.workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes));
            log.info("PoiWorkbookProcessor() - end: sheetCount={}", workbook.getNumberOfSheets());
        } catch (Exception e) {
            log.error("PoiWorkbookProcessor() - end with error: bytesLength={}", bytes.length, e);
            throw new TemplateReadWriteException("Failed to read XLS/XLSX template", e);
        }
    }

    @Override
    public TemplateScanResult scan() {
        log.info("scan() - start: sheetCount={}", workbook.getNumberOfSheets());
        TemplateScanResult result = TemplateScanner.scanPoi(workbook);
        log.info("scan() - end: markers={}, tokens={}", result.markers().size(), result.scalarTokens().size());
        return result;
    }

    @Override
    public void applyTemplateTokens(Map<String, Object> templateTokens, GenerateOptions options, WarningCollector warningCollector) {
        Map<String, Object> context = templateTokens == null ? Map.of() : templateTokens;
        int sheetCount = workbook.getNumberOfSheets();

        log.info("applyTemplateTokens() - start: sheetCount={}, tokenCount={}",
            sheetCount, context.size());
        log.info("applyTemplateTokens() - options: missingValuePolicy={}, zoneId={}, recalculateFormulas={}",
            options.missingValuePolicy(), options.zoneId(), options.recalculateFormulas());

        for (int s = 0; s < sheetCount; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            processSheetTokens(sheet, s, sheetCount, context, options, warningCollector);
        }
        log.info("applyTemplateTokens() - end: sheetCount={}", sheetCount);
    }

    private void processSheetTokens(
        Sheet sheet,
        int sheetIndex,
        int sheetCount,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        SheetProcessingState state = new SheetProcessingState();
        String sheetName = sheet.getSheetName();

        log.info("processSheetTokens() - start: sheetName={}, sheetIndex={}, sheetCount={}, lastRow={}, lastCell={}",
            sheetName,
            sheetIndex + 1,
            sheetCount,
            sheet.getLastRowNum(),
            resolveSheetLastCellNum(sheet));

        for (Row row : sheet) {
            if (row.getLastCellNum() <= 0) {
                continue;
            }
            processRowTokens(sheet, row, context, options, warningCollector, state);
        }

        state.anchors.sort(Comparator.comparingInt(PoiTableAnchor::rowIndex).reversed()
            .thenComparing(Comparator.comparingInt(PoiTableAnchor::colIndex).reversed()));

        log.info("processSheetTokens() - end: sheetName={}, processedCells={}, tableTokensFound={}, scalarTokensApplied={}, tableInsertions={}",
            sheetName, state.processedCells, state.tableTokensFound, state.scalarTokensApplied, state.anchors.size());

        for (PoiTableAnchor anchor : state.anchors) {
            insertTableAtAnchor(sheet, anchor, options, warningCollector);
        }
    }

    private void processRowTokens(
        Sheet sheet,
        Row row,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector,
        SheetProcessingState state
    ) {
        int rowIndex = row.getRowNum();
        log.info("processRowTokens() - start: rowIndex={}, cellCount={}", rowIndex, row.getLastCellNum());
        for (Cell cell : row) {
            processCellToken(sheet, row, cell, context, options, warningCollector, state);
        }
        log.info("processRowTokens() - end: rowIndex={}", rowIndex);
    }

    private void processCellToken(
        Sheet sheet,
        Row row,
        Cell cell,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector,
        SheetProcessingState state
    ) {
        int rowIndex = row.getRowNum();
        int colIndex = cell.getColumnIndex();
        String location = cellLocation(sheet, rowIndex, colIndex);
        state.processedCells++;

        if (isFormulaTokenCell(cell, location, warningCollector)) {
            return;
        }
        if (!isEligibleCellType(cell)) {
            return;
        }

        String original = getCellText(cell);
        if (!TokenResolver.hasTokens(original)) {
            return;
        }

        log.info("processCellToken() - tokenFound: location={}, tokenText={}", location, original);
        if (tryAddTableAnchor(row, rowIndex, colIndex, cell, original, context, warningCollector, location, state)) {
            return;
        }

        applyTokenToCell(
            cell,
            context,
            options.missingValuePolicy(),
            options,
            warningCollector,
            location
        );
        state.scalarTokensApplied++;
    }

    private boolean tryAddTableAnchor(
        Row row,
        int rowIndex,
        int colIndex,
        Cell cell,
        String original,
        Map<String, Object> context,
        WarningCollector warningCollector,
        String location,
        SheetProcessingState state
    ) {
        String exactToken = TokenResolver.getExactToken(original);
        String singleToken = TokenResolver.getSingleToken(original);
        String tableAnchorToken = exactToken != null ? exactToken : singleToken;
        if (tableAnchorToken == null || TokenResolver.isItemOrIndexToken(tableAnchorToken)) {
            return false;
        }

        Object resolved = TokenResolver.resolvePath(context, tableAnchorToken);
        if (!TokenResolver.isTableValue(resolved)) {
            return false;
        }

        List<Map<String, Object>> rows = TokenResolver.toTableRows(resolved);
        if (rows == null) {
            log.warn("tryAddTableAnchor() - invalidTableStructure: token={}, location={}", tableAnchorToken, location);
            warningCollector.add(
                "TABLE_TOKEN_INVALID",
                "Table token has invalid structure: " + tableAnchorToken,
                location
            );
            return true;
        }

        if (exactToken == null) {
            warningCollector.add(
                "TABLE_TOKEN_INLINE_TEXT_DROPPED",
                "Inline text around table token was removed during table insertion",
                location
            );
        }

        log.info("tryAddTableAnchor() - tableTokenFound: token={}, rowCount={}, location={}", tableAnchorToken, rows.size(), location);
        state.tableTokensFound++;
        state.anchors.add(new PoiTableAnchor(
            rowIndex,
            colIndex,
            tableAnchorToken,
            rows,
            cell.getCellStyle(),
            row.getHeight(),
            resolveConfiguredColumnOrder(context, tableAnchorToken)
        ));
        return true;
    }

    private boolean isFormulaTokenCell(Cell cell, String location, WarningCollector warningCollector) {
        if (cell.getCellType() != CellType.FORMULA) {
            return false;
        }
        String formula = cell.getCellFormula();
        if (TokenResolver.hasTokens(formula)) {
            log.info("isFormulaTokenCell() - formulaTokenSkipped: location={}, formula={}", location, formula);
            warningCollector.add(
                "FORMULA_TOKEN_SKIPPED",
                "Formula contains token and was not modified",
                location
            );
        }
        return true;
    }

    private boolean isEligibleCellType(Cell cell) {
        return cell.getCellType() == CellType.STRING || cell.getCellType() == CellType.BLANK;
    }

    private String getCellText(Cell cell) {
        return cell.getCellType() == CellType.BLANK ? "" : cell.getStringCellValue();
    }

    @Override
    public void recalculateFormulas(GenerateOptions options) {
        if (!options.recalculateFormulas()) {
            log.info("recalculateFormulas() - end: recalculated=false");
            return;
        }
        log.info("recalculateFormulas() - start: recalculated=true");
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
        log.info("recalculateFormulas() - end: recalculated=true");
    }

    @Override
    public byte[] serialize() {
        log.info("serialize() - start");
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            byte[] result = outputStream.toByteArray();
            log.info("serialize() - end: bytesLength={}", result.length);
            return result;
        } catch (Exception e) {
            log.error("serialize() - end with error", e);
            throw new TemplateReadWriteException("Failed to serialize XLS/XLSX document", e);
        }
    }

    @Override
    public void close() {
        log.info("close() - start");
        try {
            workbook.close();
            log.info("close() - end: closed=true");
        } catch (Exception e) {
            log.warn("close() - end with warning: failedToClose=true", e);
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
                log.info("applyTokenToCell() - missingExactToken: token={}, location={}", exactToken, location);
                handleMissingExactToken(cell, exactToken, policy, warningCollector, location, options);
                return;
            }
            log.info("applyTokenToCell() - exactTokenResolved: token={}, location={}, value={}", exactToken, location, resolved);
            writeValueToCell(cell, resolved, options.zoneId());
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
            log.info("applyTokenToCell() - inlineTokenResolved: location={}, from={}, to={}",
                location, original, resolvedText.value());
            cell.setCellType(CellType.STRING);
            cell.setCellValue(resolvedText.value());
        }
    }

    private void insertTableAtAnchor(
        Sheet sheet,
        PoiTableAnchor anchor,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        String location = cellLocation(sheet, anchor.rowIndex(), anchor.colIndex());
        List<Map<String, Object>> rows = anchor.rows();
        Row anchorRow = getOrCreateRow(sheet, anchor.rowIndex());
        Cell anchorCell = getOrCreateCell(anchorRow, anchor.colIndex());

        log.info("insertTableAtAnchor() - start: token={}, location={}, rowCount={}", anchor.token(), location, rows.size());

        if (rows.isEmpty()) {
            log.warn("insertTableAtAnchor() - emptyTable: token={}, location={}", anchor.token(), location);
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor.baselineStyle());
            writeValueToCell(anchorCell, null, options.zoneId());
            return;
        }

        List<String> columns = buildColumnOrder(rows, anchor.configuredColumnOrder());
        if (columns.isEmpty()) {
            log.warn("insertTableAtAnchor() - emptyColumns: token={}, location={}", anchor.token(), location);
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor.baselineStyle());
            writeValueToCell(anchorCell, null, options.zoneId());
            return;
        }

        log.info("insertTableAtAnchor() - structure: token={}, columnCount={}, columns={}, rowCount={}",
            anchor.token(), columns.size(), String.join(",", columns), rows.size());

        int dataRowCount = rows.size();
        if (dataRowCount > 0 && anchor.rowIndex() + 1 <= sheet.getLastRowNum()) {
            log.info("insertTableAtAnchor() - shiftRows: shiftCount={}, startRow={}", dataRowCount, anchor.rowIndex() + 1);
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
                Object value = values.get(column);
                writeValueToCell(cell, value, options.zoneId());
            }
        }

        autoResizeTableColumns(sheet, anchor.colIndex(), columns, rows);
        log.info("insertTableAtAnchor() - end: token={}, location={}, columnCount={}, rowCount={}",
            anchor.token(), location, columns.size(), rows.size());
    }

    private void handleMissingExactToken(
        Cell cell,
        String token,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location,
        GenerateOptions options
    ) {
        log.info("handleMissingExactToken() - start: token={}, policy={}, location={}", token, policy, location);
        switch (policy) {
            case EMPTY_AND_LOG -> {
                log.info("handleMissingExactToken() - action: setBlank, token={}, location={}", token, location);
                warningCollector.add("MISSING_TOKEN", "Token not found: " + token, location);
                writeValueToCell(cell, null, options.zoneId());
            }
            case LEAVE_TOKEN -> {
                log.info("handleMissingExactToken() - action: leaveToken, token={}, location={}", token, location);
                // no-op
            }
            case FAIL_FAST -> {
                log.error("handleMissingExactToken() - action: failFast, token={}, location={}", token, location);
                throw new TemplateDataBindingException("Token not found: " + token + " at " + location);
            }
        }
        log.info("handleMissingExactToken() - end: token={}, policy={}, location={}", token, policy, location);
    }

    private void autoResizeTableColumns(
        Sheet sheet,
        int startColumnIndex,
        List<String> columns,
        List<Map<String, Object>> rows
    ) {
        log.info("autoResizeTableColumns() - start: columnCount={}, startColumnIndex={}", columns.size(), startColumnIndex);
        for (int c = 0; c < columns.size(); c++) {
            String column = columns.get(c);
            int maxLength = column.length();
            for (Map<String, Object> row : rows) {
                maxLength = Math.max(maxLength, stringifyLength(row.get(column)));
            }

            int desiredWidth = calculateDesiredWidth(maxLength);
            int targetColumn = startColumnIndex + c;
            int currentWidth = sheet.getColumnWidth(targetColumn);

            log.info("autoResizeTableColumns() - evaluateColumn: column={}, maxContentLength={}, currentWidth={}, desiredWidth={}",
                column, maxLength, currentWidth, desiredWidth);

            if (currentWidth < desiredWidth) {
                log.info("autoResizeTableColumns() - resizeColumn: column={}, from={}, to={}", column, currentWidth, desiredWidth);
                sheet.setColumnWidth(targetColumn, desiredWidth);
            }
        }
        log.info("autoResizeTableColumns() - end: columnCount={}", columns.size());
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

    private List<String> buildColumnOrder(List<Map<String, Object>> rows, List<String> configuredColumnOrder) {
        LinkedHashSet<String> ordered = new LinkedHashSet<>();
        if (configuredColumnOrder != null && !configuredColumnOrder.isEmpty()) {
            ordered.addAll(configuredColumnOrder);
        } else {
            ordered.addAll(rows.get(0).keySet());
        }
        for (Map<String, Object> row : rows) {
            ordered.addAll(row.keySet());
        }
        return List.copyOf(ordered);
    }

    private List<String> resolveConfiguredColumnOrder(Map<String, Object> context, String tableToken) {
        Object raw = context.get(tableToken + "__columns");
        if (raw == null) {
            raw = context.get(tableToken + "_columns");
        }
        if (raw == null) {
            raw = context.get(tableToken + ".columns");
        }
        if (raw instanceof List<?> list) {
            List<String> columns = list.stream()
                .filter(Objects::nonNull)
                .map(String::valueOf)
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .toList();
            if (!columns.isEmpty()) {
                return columns;
            }
        }
        return List.of();
    }

    private void writeValueToCell(Cell cell, Object value, ZoneId zoneId) {
        if (value == null) {
            cell.setBlank();
            return;
        }
        switch (value) {
            case Number number -> {
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(number.doubleValue());
            }
            case Boolean bool -> {
                cell.setCellType(CellType.BOOLEAN);
                cell.setCellValue(bool);
            }
            case Date date -> cell.setCellValue(date);
            case LocalDate localDate -> {
                Date date = Date.from(localDate.atStartOfDay(zoneId).toInstant());
                cell.setCellValue(date);
            }
            case LocalDateTime localDateTime -> {
                Date date = Date.from(localDateTime.atZone(zoneId).toInstant());
                cell.setCellValue(date);
            }
            case Instant instant -> cell.setCellValue(Date.from(instant));
            default -> {
                cell.setCellType(CellType.STRING);
                cell.setCellValue(String.valueOf(value));
            }
        }
    }

    private void applyBaselineStyle(Cell cell, CellStyle baselineStyle) {
        if (baselineStyle != null) {
            cell.setCellStyle(baselineStyle);
        }
    }

    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            log.info("getOrCreateRow() - create: rowIndex={}", rowIndex);
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    private Cell getOrCreateCell(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            log.info("getOrCreateCell() - create: columnIndex={}", colIndex);
            cell = row.createCell(colIndex);
        }
        return cell;
    }

    private String cellLocation(Sheet sheet, int rowIndex, int colIndex) {
        return sheet.getSheetName() + "!R" + (rowIndex + 1) + "C" + (colIndex + 1);
    }

    private short resolveSheetLastCellNum(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return 0;
        }
        Row row = sheet.getRow(lastRowNum);
        return row == null ? 0 : row.getLastCellNum();
    }

    private static final class SheetProcessingState {
        private int processedCells;
        private int tableTokensFound;
        private int scalarTokensApplied;
        private final List<PoiTableAnchor> anchors = new ArrayList<>();
    }

}
