package io.github.ogbozoyan.processor;

import io.github.ogbozoyan.data.GenerateOptions;
import io.github.ogbozoyan.data.MissingValuePolicy;
import io.github.ogbozoyan.data.PoiTableAnchor;
import io.github.ogbozoyan.data.ResolvedText;
import io.github.ogbozoyan.data.SheetProcessingState;
import io.github.ogbozoyan.data.TemplateScanResult;
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
import java.util.LinkedHashMap;
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
 *     <li>rows-only table insertion when {@link GenerateOptions#rowsOnlyTableTokens()} is {@code true}
 *     and token value is {@code List<Object[]>},</li>
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
     * Creates processor and parses XLS/XLSX bytes via Apache POI.
     *
     * @param bytes source workbook bytes
     * @throws TemplateReadWriteException when workbook cannot be parsed
     */
    public PoiWorkbookProcessor(byte[] bytes) {
        log.debug("PoiWorkbookProcessor() - start: bytesLength={}", bytes.length);
        try {
            this.workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes));
            log.debug("PoiWorkbookProcessor() - end: sheetCount={}", workbook.getNumberOfSheets());
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
        log.trace("scan() - start: sheetCount={}", workbook.getNumberOfSheets());
        TemplateScanResult result = TemplateScanner.scanPoi(workbook);
        log.trace("scan() - end: markers={}, tokens={}", result.markers().size(), result.scalarTokens().size());
        return result;
    }

    /**
     * Applies scalar and table tokens to all sheets in the workbook.
     *
     * <p>Table insertion is triggered when token value resolves to:
     * <ul>
     *     <li>{@code List<Map<String, Object>>} in default mode,</li>
     *     <li>{@code List<Object[]>} when {@link GenerateOptions#rowsOnlyTableTokens()} is {@code true}.</li>
     * </ul>
     * During scan phase table anchors are collected, then applied in reverse order after scalar pass.
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
     *                         in default mode or {@code List<Object[]>} in rows-only mode
     * @param options          generation options
     * @param warningCollector collector for non-fatal issues
     */
    @Override
    public void applyTemplateTokens(Map<String, Object> templateTokens, GenerateOptions options, WarningCollector warningCollector) {
        Map<String, Object> context = templateTokens == null ? Map.of() : templateTokens;
        int sheetCount = workbook.getNumberOfSheets();

        log.trace("applyTemplateTokens() - start: sheetCount={}, tokenCount={}",
            sheetCount, context.size());
        log.trace("applyTemplateTokens() - options: missingValuePolicy={}, zoneId={}, recalculateFormulas={}",
            options.missingValuePolicy(), options.zoneId(), options.recalculateFormulas());
        log.trace("applyTemplateTokens() - tableMode: rowsOnlyTableTokens={}", options.rowsOnlyTableTokens());

        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            processSheetTokens(sheet, sheetIndex, sheetCount, context, options, warningCollector);
        }
        log.trace("applyTemplateTokens() - end: sheetCount={}", sheetCount);
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

        log.trace("processSheetTokens() - start: sheetName={}, sheetIndex={}, sheetCount={}, lastRow={}, lastCell={}",
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

        state.getAnchors()
            .sort(
                Comparator
                    .comparingInt(PoiTableAnchor::rowIndex)
                    .reversed()
                    .thenComparing(
                        Comparator.comparingInt(PoiTableAnchor::colIndex).reversed()
                    )
            );

        log.trace("processSheetTokens() - end: sheetName={}, processedCells={}, tableTokensFound={}, scalarTokensApplied={}, tableInsertions={}",
            sheetName, state.getProcessedCells(), state.getTableTokensFound(), state.getScalarTokensApplied(), state.getAnchors().size());

        for (PoiTableAnchor anchor : state.getAnchors()) {
            insertTableAtAnchor(sheet, anchor, options, warningCollector);
        }
    }

    /**
     * Processes all physical cells in one row.
     *
     * @param sheet            current sheet
     * @param row              current row
     * @param context          token context
     * @param options          generation options
     * @param warningCollector warning collector
     * @param state            mutable sheet processing state
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
        log.trace("processRowTokens() - start: rowIndex={}, cellCount={}", rowIndex, row.getLastCellNum());
        for (Cell cell : row) {
            processCellToken(sheet, row, cell, context, options, warningCollector, state);
        }
        log.trace("processRowTokens() - end: rowIndex={}", rowIndex);
    }

    /**
     * Processes single cell and routes to table-anchor or scalar replacement flow.
     *
     * @param sheet            current sheet
     * @param row              current row
     * @param cell             current cell
     * @param context          token context
     * @param options          generation options
     * @param warningCollector warning collector
     * @param state            mutable sheet processing state
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
        state.incrementProcessedCells();

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

        log.trace("processCellToken() - tokenFound: location={}, tokenText={}", location, original);
        if (tryAddTableAnchor(row, rowIndex, colIndex, cell, original, context, options, warningCollector, location, state)) {
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
        state.incrementScalarTokensApplied();
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
        GenerateOptions options,
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

        boolean rowsOnly = options.rowsOnlyTableTokens();
        Object resolved = TokenResolver.resolvePath(context, tableAnchorToken);
        List<String> configuredColumnOrder = resolveConfiguredColumnOrder(context, tableAnchorToken);
        List<Map<String, Object>> rows = rowsOnly
            ? toRowsOnlyTableRows(resolved, configuredColumnOrder)
            : TokenResolver.toTableRows(resolved);
        if (rows == null) {
            if (rowsOnly && !(resolved instanceof List<?>)) {
                return false;
            }
            if (!rowsOnly && !TokenResolver.isTableValue(resolved)) {
                return false;
            }
            log.warn("tryAddTableAnchor() - invalidTableStructure: token={}, location={}", tableAnchorToken, location);
            warningCollector.add(
                "TABLE_TOKEN_INVALID",
                rowsOnly
                    ? "Rows-only table token must be List<Object[]>: " + tableAnchorToken
                    : "Table token has invalid structure: " + tableAnchorToken,
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

        log.trace("tryAddTableAnchor() - tableTokenFound: token={}, rowCount={}, location={}", tableAnchorToken, rows.size(), location);
        state.incrementTableTokensFound();
        state.getAnchors()
            .add(
                new PoiTableAnchor(
                    rowIndex,
                    colIndex,
                    tableAnchorToken,
                    rows,
                    cell.getCellStyle(),
                    row.getHeight(),
                    configuredColumnOrder
                )
            );
        return true;
    }

    /**
     * Detects formula cells and emits warning when formula contains token syntax.
     *
     * @param cell             source cell
     * @param location         diagnostic location
     * @param warningCollector warning collector
     * @return {@code true} when caller must skip further token replacement
     */
    private boolean isFormulaTokenCell(Cell cell, String location, WarningCollector warningCollector) {
        if (cell.getCellType() != CellType.FORMULA) {
            return false;
        }
        String formula = cell.getCellFormula();
        if (TokenResolver.hasTokens(formula)) {
            log.trace("isFormulaTokenCell() - formulaTokenSkipped: location={}, formula={}", location, formula);
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
            log.trace("recalculateFormulas() - end: recalculated=false");
            return;
        }
        log.trace("recalculateFormulas() - start: recalculated=true");
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
        log.trace("recalculateFormulas() - end: recalculated=true");
    }

    /**
     * Serializes modified workbook into byte array.
     *
     * @return generated report bytes
     */
    @Override
    public byte[] serialize() {
        log.trace("serialize() - start");
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            byte[] result = outputStream.toByteArray();
            log.trace("serialize() - end: bytesLength={}", result.length);
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
        log.trace("close() - start");
        try {
            workbook.close();
            log.trace("close() - end: closed=true");
        } catch (Exception e) {
            log.warn("close() - end with warning: failedToClose=true", e);
        }
    }

    /**
     * Applies scalar token replacement for one cell.
     *
     * <p>Exact token replacement preserves semantic type. Inline replacement writes string value.
     *
     * @param cell             destination cell
     * @param context          token context
     * @param policy           unresolved token policy
     * @param options          generation options
     * @param warningCollector warning collector
     * @param location         diagnostic location
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
                log.trace("applyTokenToCell() - missingExactToken: token={}, location={}", exactToken, location);
                handleMissingExactToken(cell, exactToken, policy, warningCollector, location, options);
                return;
            }
            log.trace("applyTokenToCell() - exactTokenResolved: token={}, location={}, value={}", exactToken, location, resolved);
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
            log.trace("applyTokenToCell() - inlineTokenResolved: location={}, from={}, to={}",
                location, original, resolvedText.value());
            cell.setCellType(CellType.STRING);
            cell.setCellValue(resolvedText.value());
        }
    }

    /**
     * Inserts a table token at anchor cell.
     *
     * <p>The insertion mode depends on {@link GenerateOptions#rowsOnlyTableTokens()}:
     * header+data by default, or rows-only when enabled.
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

        log.trace("insertTableAtAnchor() - start: token={}, location={}, rowCount={}", anchor.token(), location, rows.size());

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

        boolean rowsOnly = options.rowsOnlyTableTokens();
        log.trace("insertTableAtAnchor() - structure: token={}, columnCount={}, columns={}, rowCount={}, rowsOnly={}",
            anchor.token(), columns.size(), String.join(",", columns), rows.size(), rowsOnly);

        if (rowsOnly) {
            insertRowsOnlyAtAnchor(sheet, anchor, options, columns, rows);
        } else {
            insertTableWithHeaderAtAnchor(sheet, anchor, options, columns, rows);
        }

        log.trace("insertTableAtAnchor() - end: token={}, location={}, columnCount={}, rowCount={}",
            anchor.token(), location, columns.size(), rows.size());
    }

    /**
     * Inserts table rows with header at marker position.
     *
     * @param sheet   target sheet
     * @param anchor  insertion anchor
     * @param options generation options
     * @param columns ordered columns
     * @param rows    table rows
     */
    private void insertTableWithHeaderAtAnchor(
        Sheet sheet,
        PoiTableAnchor anchor,
        GenerateOptions options,
        List<String> columns,
        List<Map<String, Object>> rows
    ) {
        int dataRowCount = rows.size();
        if (dataRowCount > 0 && anchor.rowIndex() + 1 <= sheet.getLastRowNum()) {
            log.trace("insertTableWithHeaderAtAnchor() - shiftRows: shiftCount={}, startRow={}", dataRowCount, anchor.rowIndex() + 1);
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

        autoResizeTableColumns(sheet, anchor.colIndex(), columns, rows, true);
    }

    /**
     * Inserts table rows only (without header row) at marker position.
     *
     * @param sheet   target sheet
     * @param anchor  insertion anchor
     * @param options generation options
     * @param columns ordered columns
     * @param rows    table rows
     */
    private void insertRowsOnlyAtAnchor(
        Sheet sheet,
        PoiTableAnchor anchor,
        GenerateOptions options,
        List<String> columns,
        List<Map<String, Object>> rows
    ) {
        int shiftCount = rows.size() - 1;
        if (shiftCount > 0 && anchor.rowIndex() + 1 <= sheet.getLastRowNum()) {
            log.trace("insertRowsOnlyAtAnchor() - shiftRows: shiftCount={}, startRow={}", shiftCount, anchor.rowIndex() + 1);
            sheet.shiftRows(anchor.rowIndex() + 1, sheet.getLastRowNum(), shiftCount, true, false);
        }

        for (int r = 0; r < rows.size(); r++) {
            Map<String, Object> values = rows.get(r);
            Row dataRow = getOrCreateRow(sheet, anchor.rowIndex() + r);
            dataRow.setHeight(anchor.baselineRowHeight());

            for (int c = 0; c < columns.size(); c++) {
                String column = columns.get(c);
                Cell cell = getOrCreateCell(dataRow, anchor.colIndex() + c);
                applyBaselineStyle(cell, anchor.baselineStyle());
                Object value = values.get(column);
                writeValueToCell(cell, value, options.zoneId());
            }
        }

        autoResizeTableColumns(sheet, anchor.colIndex(), columns, rows, false);
    }

    /**
     * Handles unresolved exact token according to missing-value policy.
     *
     * @param cell             destination cell
     * @param token            unresolved token name
     * @param policy           missing-value policy
     * @param warningCollector warning collector
     * @param location         diagnostic location
     * @param options          generation options
     */
    private void handleMissingExactToken(
        Cell cell,
        String token,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location,
        GenerateOptions options
    ) {
        log.trace("handleMissingExactToken() - start: token={}, policy={}, location={}", token, policy, location);
        switch (policy) {
            case EMPTY_AND_LOG -> {
                log.trace("handleMissingExactToken() - action: setBlank, token={}, location={}", token, location);
                warningCollector.add("MISSING_TOKEN", "Token not found: " + token, location);
                writeValueToCell(cell, null, options.zoneId());
            }
            case LEAVE_TOKEN -> {
                log.trace("handleMissingExactToken() - action: leaveToken, token={}, location={}", token, location);
                // no-op
            }
            case FAIL_FAST -> {
                log.error("handleMissingExactToken() - action: failFast, token={}, location={}", token, location);
                throw new TemplateDataBindingException("Token not found: " + token + " at " + location);
            }
        }
        log.trace("handleMissingExactToken() - end: token={}, policy={}, location={}", token, policy, location);
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
        List<Map<String, Object>> rows,
        boolean includeHeader
    ) {
        log.trace("autoResizeTableColumns() - start: columnCount={}, startColumnIndex={}, includeHeader={}",
            columns.size(), startColumnIndex, includeHeader);
        for (int c = 0; c < columns.size(); c++) {
            String column = columns.get(c);
            int maxLength = includeHeader ? column.length() : 0;
            for (Map<String, Object> row : rows) {
                maxLength = Math.max(maxLength, stringifyLength(row.get(column)));
            }

            int desiredWidth = calculateDesiredWidth(maxLength);
            int targetColumn = startColumnIndex + c;
            int currentWidth = sheet.getColumnWidth(targetColumn);

            log.trace("autoResizeTableColumns() - evaluateColumn: column={}, maxContentLength={}, currentWidth={}, desiredWidth={}",
                column, maxLength, currentWidth, desiredWidth);

            if (currentWidth < desiredWidth) {
                log.trace("autoResizeTableColumns() - resizeColumn: column={}, from={}, to={}", column, currentWidth, desiredWidth);
                sheet.setColumnWidth(targetColumn, desiredWidth);
            }
        }
        log.trace("autoResizeTableColumns() - end: columnCount={}", columns.size());
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
     * @param rows                  normalized table rows
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
     * Converts rows-only token payload into internal map-row representation.
     *
     * <p>Expected payload shape is {@code List<Object[]>}. Each array item represents
     * a single row where field order maps to columns by index.
     *
     * <p>Column naming strategy:
     * <ul>
     *     <li>configured column names are used first by index ({@code TOKEN__columns}, etc.),</li>
     *     <li>remaining unnamed fields use synthetic keys {@code __col{index}}.</li>
     * </ul>
     *
     * @param value                 raw token value
     * @param configuredColumnOrder optional configured column names
     * @return normalized rows or {@code null} for unsupported structure
     */
    private List<Map<String, Object>> toRowsOnlyTableRows(Object value, List<String> configuredColumnOrder) {
        if (!(value instanceof List<?> list)) {
            return null;
        }
        if (list.isEmpty()) {
            return List.of();
        }

        List<Map<String, Object>> rows = new ArrayList<>(list.size());
        for (int rowIndex = 0; rowIndex < list.size(); rowIndex++) {
            Object rawRow = list.get(rowIndex);
            if (!(rawRow instanceof Object[] rowArray)) {
                return null;
            }

            LinkedHashMap<String, Object> normalizedRow = new LinkedHashMap<>();
            for (int columnIndex = 0; columnIndex < rowArray.length; columnIndex++) {
                String columnName = resolveRowsOnlyColumnName(configuredColumnOrder, columnIndex);
                normalizedRow.put(columnName, rowArray[columnIndex]);
            }
            rows.add(normalizedRow);
        }
        return rows;
    }

    /**
     * Resolves internal column key for rows-only row arrays.
     *
     * @param configuredColumnOrder optional configured column names
     * @param columnIndex           zero-based column index
     * @return configured name when present, otherwise synthetic key
     */
    private String resolveRowsOnlyColumnName(List<String> configuredColumnOrder, int columnIndex) {
        if (columnIndex < configuredColumnOrder.size()) {
            return configuredColumnOrder.get(columnIndex);
        }
        return "__col" + columnIndex;
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
        if (value instanceof Number number) {
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(number.doubleValue());
        } else if (value instanceof Boolean bool) {
            cell.setCellType(CellType.BOOLEAN);
            cell.setCellValue(bool);
        } else if (value instanceof Date date) {
            cell.setCellValue(date);
        } else if (value instanceof LocalDate localDate) {
            Date date = Date.from(localDate.atStartOfDay(zoneId).toInstant());
            cell.setCellValue(date);
        } else if (value instanceof LocalDateTime localDateTime) {
            Date date = Date.from(localDateTime.atZone(zoneId).toInstant());
            cell.setCellValue(date);
        } else if (value instanceof Instant instant) {
            cell.setCellValue(Date.from(instant));
        } else {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(String.valueOf(value));
        }
    }

    /**
     * Applies baseline style to cell when style is available.
     *
     * @param cell          destination cell
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
     * @param sheet    target sheet
     * @param rowIndex zero-based row index
     * @return existing or new row
     */
    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            log.trace("getOrCreateRow() - create: rowIndex={}", rowIndex);
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    /**
     * Returns cell by index or creates missing cell.
     *
     * @param row      target row
     * @param colIndex zero-based column index
     * @return existing or new cell
     */
    private Cell getOrCreateCell(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            log.trace("getOrCreateCell() - create: columnIndex={}", colIndex);
            cell = row.createCell(colIndex);
        }
        return cell;
    }

    /**
     * Builds one-based cell location for diagnostics.
     *
     * @param sheet    source sheet
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


}
