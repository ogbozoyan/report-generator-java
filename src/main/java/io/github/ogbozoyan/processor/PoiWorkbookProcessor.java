package io.github.ogbozoyan.processor;

import io.github.ogbozoyan.contract.GenerateOptions;
import io.github.ogbozoyan.contract.MissingValuePolicy;
import io.github.ogbozoyan.contract.PoiTableAnchor;
import io.github.ogbozoyan.contract.ResolvedText;
import io.github.ogbozoyan.contract.TemplateScanResult;
import io.github.ogbozoyan.exception.TemplateDataBindingException;
import io.github.ogbozoyan.exception.TemplateReadWriteException;
import io.github.ogbozoyan.util.TemplateScanner;
import io.github.ogbozoyan.util.TokenResolver;
import io.github.ogbozoyan.util.WarningCollector;
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
 * Spreadsheet io.github.ogbozoyan.processor for {@code .xls}/{@code .xlsx} templates based on Apache POI.
 *
 * <p>This io.github.ogbozoyan.processor is responsible for:
 * <ul>
 *     <li>scalar token replacement for cells containing {@code {{token}}} expressions,</li>
 *     <li>table token insertion for values of type {@code List<Map<String, Object>>},</li>
 *     <li>baseline style reuse from marker cell,</li>
 *     <li>optional formula recalculation,</li>
 *     <li>serialization back to workbook bytes.</li>
 * </ul>
 *
 * <p>Key design choices:
 * <ul>
 *     <li>iterate physical rows/cells only, to avoid full-grid scans on sparse templates,</li>
 *     <li>collect table anchors first and insert them in reverse order, so row shifting does not break coordinates,</li>
 *     <li>resize only inserted table columns and clamp width to a safe range,</li>
 *     <li>preserve table marker style for both header and data rows.</li>
 * </ul>
 *
 * <p>Typical usage from io.github.ogbozoyan.service layer:
 * <pre>{@code
 * try (PoiWorkbookProcessor io.github.ogbozoyan.processor = new PoiWorkbookProcessor(templateBytes)) {
 *     io.github.ogbozoyan.processor.applyTemplateTokens(reportData.templateTokens(), options, warningCollector);
 *     io.github.ogbozoyan.processor.recalculateFormulas(options);
 *     byte[] generated = io.github.ogbozoyan.processor.serialize();
 * }
 * }</pre>
 */
@Slf4j
public class PoiWorkbookProcessor implements WorkbookProcessor {

    private final Workbook workbook;

    /**
     * Creates io.github.ogbozoyan.processor and parses XLS/XLSX bytes via Apache POI.
     *
     * @param bytes source workbook bytes
     * @throws TemplateReadWriteException when workbook cannot be parsed
     */
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

    /**
     * Scans workbook for scalar token occurrences and legacy markers.
     *
     * @return scan result with token occurrences
     */
    @Override
    public TemplateScanResult scan() {
        log.info("scan() - start: sheetCount={}", workbook.getNumberOfSheets());
        TemplateScanResult result = TemplateScanner.scanPoi(workbook);
        log.info("scan() - end: markers={}, tokens={}", result.markers().size(), result.scalarTokens().size());
        return result;
    }

    /**
     * Applies scalar and table tokens to all sheets in the workbook.
     *
     * <p>Table insertion is triggered when token value resolves to
     * {@code List<Map<String, Object>>}. During scan phase table anchors are collected,
     * then applied in reverse order after scalar pass.
     *
     * <p>Example:
     * <pre>{@code
     * Map<String, Object> tokens = Map.of(
     *     "period", "2026-Q1",
     *     "rows", List.of(
     *         Map.of("name", "North", "amount", 1200.25),
     *         Map.of("name", "South", "amount", 900.00)
     *     ),
     *     "rows__columns", List.of("name", "amount")
     * );
     * io.github.ogbozoyan.processor.applyTemplateTokens(tokens, GenerateOptions.defaults(), warningCollector);
     * }</pre>
     *
     * @param templateTokens   token map; table token must be {@code List<Map<String, Object>>}
     * @param options          generation options
     * @param warningCollector collector for non-fatal issues
     */
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

    /**
     * Processes a single sheet: scans tokens, collects table anchors, then inserts tables.
     *
     * <p>The method is intentionally split into phases:
     * <ol>
     *     <li>scan physical rows and cells,</li>
     *     <li>accumulate table anchors,</li>
     *     <li>insert tables from bottom-to-top.</li>
     * </ol>
     */
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

    /**
     * Processes all physical cells in one row.
     *
     * @param sheet current sheet
     * @param row current row
     * @param context token context
     * @param options generation options
     * @param warningCollector warning collector
     * @param state mutable sheet processing state
     */
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

    /**
     * Processes single cell and routes to table-anchor or scalar replacement flow.
     *
     * @param sheet current sheet
     * @param row current row
     * @param cell current cell
     * @param context token context
     * @param options generation options
     * @param warningCollector warning collector
     * @param state mutable sheet processing state
     */
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

    /**
     * Attempts to treat current token cell as table anchor.
     *
     * <p>Returns {@code true} when cell was consumed as table-related case
     * (valid anchor or invalid table structure warning) and should not be processed as scalar.
     */
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

    /**
     * Detects formula cells and emits warning when formula contains token syntax.
     *
     * @param cell source cell
     * @param location diagnostic location
     * @param warningCollector warning collector
     * @return {@code true} when caller must skip further token replacement
     */
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

    /**
     * Checks whether cell type participates in token replacement phase.
     *
     * @param cell source cell
     * @return {@code true} for string/blank cells
     */
    private boolean isEligibleCellType(Cell cell) {
        return cell.getCellType() == CellType.STRING || cell.getCellType() == CellType.BLANK;
    }

    /**
     * Reads text from eligible cell.
     *
     * @param cell source cell
     * @return text for token inspection
     */
    private String getCellText(Cell cell) {
        return cell.getCellType() == CellType.BLANK ? "" : cell.getStringCellValue();
    }

    /**
     * Recalculates all workbook formulas if enabled in options.
     *
     * @param options generation options
     */
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

    /**
     * Serializes modified workbook into byte array.
     *
     * @return generated report bytes
     */
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

    /**
     * Closes underlying POI workbook.
     */
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

    /**
     * Applies scalar token replacement for one cell.
     *
     * <p>Exact token replacement preserves semantic type. Inline replacement writes string value.
     *
     * @param cell destination cell
     * @param context token context
     * @param policy unresolved token policy
     * @param options generation options
     * @param warningCollector warning collector
     * @param location diagnostic location
     */
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

    /**
     * Inserts a table at anchor cell and shifts following rows down when needed.
     *
     * <p>Behavior:
     * <ul>
     *     <li>header row is created at marker row,</li>
     *     <li>data rows are created below header,</li>
     *     <li>marker cell style is reused as baseline for all inserted cells,</li>
     *     <li>table columns are auto-resized after data insertion.</li>
     * </ul>
     */
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
        for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
            Cell cell = getOrCreateCell(headerRow, anchor.colIndex() + columnIndex);
            applyBaselineStyle(cell, anchor.baselineStyle());
            cell.setCellType(CellType.STRING);
            cell.setCellValue(columns.get(columnIndex));
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

    /**
     * Handles unresolved exact token according to missing-value policy.
     *
     * @param cell destination cell
     * @param token unresolved token name
     * @param policy missing-value policy
     * @param warningCollector warning collector
     * @param location diagnostic location
     * @param options generation options
     */
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

    /**
     * Applies width expansion for table columns only.
     *
     * <p>The width is derived from max text length of header/data and clamped to
     * avoid unreadable extremely narrow or huge columns.
     */
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

    /**
     * Calculates desired POI column width from maximum content length.
     *
     * @param maxLength max string length in column
     * @return clamped width in 1/256th character units
     */
    private int calculateDesiredWidth(int maxLength) {
        int width = (maxLength + 2) * 256;
        int min = 8 * 256;
        int max = 100 * 256;
        return Math.max(min, Math.min(width, max));
    }

    /**
     * Converts value to string length for width estimation.
     *
     * @param value source value
     * @return string length or {@code 0} for {@code null}
     */
    private int stringifyLength(Object value) {
        return value == null ? 0 : String.valueOf(value).length();
    }

    /**
     * Builds final column order for table insertion.
     *
     * <p>Configured order is used as base when provided, then unseen keys are appended
     * in row encounter order.
     *
     * @param rows normalized table rows
     * @param configuredColumnOrder optional explicit column order
     * @return ordered column names
     */
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

    /**
     * Resolves optional explicit column order for a table token.
     *
     * <p>Supported keys:
     * <ul>
     *     <li>{@code token__columns}</li>
     *     <li>{@code token_columns}</li>
     *     <li>{@code token.columns}</li>
     * </ul>
     *
     * <p>Example:
     * <pre>{@code
     * Map<String, Object> tokens = Map.of(
     *     "rows", rows,
     *     "rows__columns", List.of("name", "amount", "region")
     * );
     * }</pre>
     */
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

    /**
     * Writes Java value to cell preserving semantic type where possible.
     *
     * <p>Mapping examples:
     * <ul>
     *     <li>{@link Number} -> numeric cell,</li>
     *     <li>{@link Boolean} -> boolean cell,</li>
     *     <li>{@link LocalDate}/{@link LocalDateTime}/{@link Instant}/{@link Date} -> date cell value,</li>
     *     <li>fallback -> string.</li>
     * </ul>
     */
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

    /**
     * Applies baseline style to cell when style is available.
     *
     * @param cell destination cell
     * @param baselineStyle style copied from marker cell
     */
    private void applyBaselineStyle(Cell cell, CellStyle baselineStyle) {
        if (baselineStyle != null) {
            cell.setCellStyle(baselineStyle);
        }
    }

    /**
     * Returns row by index or creates missing row.
     *
     * @param sheet target sheet
     * @param rowIndex zero-based row index
     * @return existing or new row
     */
    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            log.info("getOrCreateRow() - create: rowIndex={}", rowIndex);
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    /**
     * Returns cell by index or creates missing cell.
     *
     * @param row target row
     * @param colIndex zero-based column index
     * @return existing or new cell
     */
    private Cell getOrCreateCell(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            log.info("getOrCreateCell() - create: columnIndex={}", colIndex);
            cell = row.createCell(colIndex);
        }
        return cell;
    }

    /**
     * Builds one-based cell location for diagnostics.
     *
     * @param sheet source sheet
     * @param rowIndex zero-based row index
     * @param colIndex zero-based column index
     * @return location string
     */
    private String cellLocation(Sheet sheet, int rowIndex, int colIndex) {
        return sheet.getSheetName() + "!R" + (rowIndex + 1) + "C" + (colIndex + 1);
    }

    /**
     * Resolves last cell number for logging current sheet geometry.
     *
     * @param sheet source sheet
     * @return last cell number in last row, or {@code 0} when unavailable
     */
    private short resolveSheetLastCellNum(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return 0;
        }
        Row row = sheet.getRow(lastRowNum);
        return row == null ? 0 : row.getLastCellNum();
    }

    /**
     * Mutable aggregation state for single sheet processing pass.
     */
    private static final class SheetProcessingState {
        private int processedCells;
        private int tableTokensFound;
        private int scalarTokensApplied;
        private final List<PoiTableAnchor> anchors = new ArrayList<>();
    }

}
