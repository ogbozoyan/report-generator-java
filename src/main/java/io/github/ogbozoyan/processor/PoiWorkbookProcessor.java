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
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

import static io.github.ogbozoyan.helper.PoiHelper.applyBaselineStyle;
import static io.github.ogbozoyan.helper.PoiHelper.buildColumnOrder;
import static io.github.ogbozoyan.helper.PoiHelper.calculateDesiredWidth;
import static io.github.ogbozoyan.helper.PoiHelper.cellLocation;
import static io.github.ogbozoyan.helper.PoiHelper.getOrCreateCell;
import static io.github.ogbozoyan.helper.PoiHelper.getOrCreateRow;
import static io.github.ogbozoyan.helper.PoiHelper.isEligibleCellType;
import static io.github.ogbozoyan.helper.PoiHelper.resolveConfiguredColumnOrder;
import static io.github.ogbozoyan.helper.PoiHelper.resolveSheetLastCellNum;
import static io.github.ogbozoyan.helper.PoiHelper.stringifyLength;
import static io.github.ogbozoyan.helper.PoiHelper.toRowsOnlyTableRows;
import static io.github.ogbozoyan.helper.PoiHelper.writeValueToCell;

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
 *     <li>run table insertion in multi-pass mode with reverse anchor application, so row shifting
 *     and newly appeared table tokens are handled deterministically,</li>
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

    private static final int MAX_TABLE_PASSES = 50;

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
     * During table phase anchors are collected and applied in reverse order for each pass.
     * Additional passes run until no new table anchors are found, then scalar pass is executed.
     *
     * <p>Example:
     * <pre>{@code
     * Map<String, Object> tokens = Map.of(
     *     "period", "2026-Q1",
     *     "rows", List.of(
     *         Map.of("name", "North", "amount", 1200.25),
     *         Map.of("name", "South", "amount", 900.00)
     *     ),
     *     TagConstants.ROWS_COLUMNS.getValue(), List.of("name", "amount")
     * );
     * io.github.ogbozoyan.processor.applyTemplateTokens(tokens, GenerateOptions.defaults(), warningCollector);
     * }</pre>
     *
     * @param templateTokensMappings token map; table token must be {@code List<Map<String, Object>>}
     *                               in default mode or {@code List<Object[]>} in rows-only mode
     * @param options                generation options
     * @param warningCollector       collector for non-fatal issues
     */
    @Override
    public void process(Map<String, Object> templateTokensMappings, GenerateOptions options, WarningCollector warningCollector) {
        Map<String, Object> context = templateTokensMappings == null ? Map.of() : templateTokensMappings;
        int sheetCount = workbook.getNumberOfSheets();

        log.trace("process() - start: sheetCount={}, tokenCount={}, missingValuePolicy={}, zoneId={}, recalculateFormulas={}, rowsOnlyTableTokens={}",
            sheetCount, context.size(), options.missingValuePolicy(), options.zoneId(), options.recalculateFormulas(), options.rowsOnlyTableTokens());

        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            processSheetTokens(sheet, sheetIndex, sheetCount, context, options, warningCollector);
        }
        log.trace("process() - end: sheetCount={}", sheetCount);
    }

    /**
     * Processes a single sheet using two phases: table passes, then scalar pass.
     *
     * <p>The method is intentionally split into phases:
     * <ol>
     *     <li>repeat table passes until no anchors are found (or {@link #MAX_TABLE_PASSES} reached),</li>
     *     <li>for each table pass: scan physical rows/cells, collect anchors, apply bottom-to-top,</li>
     *     <li>run scalar replacement pass on the stabilized sheet.</li>
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
        String sheetName = sheet.getSheetName();
        int totalTablePasses = 0;
        int totalTableTokensFound = 0;
        int totalTableInsertions = 0;

        log.trace("processSheetTokens() - start: sheetName={}, sheetIndex={}, sheetCount={}, lastRow={}, lastCell={}",
            sheetName,
            sheetIndex + 1,
            sheetCount,
            sheet.getLastRowNum(),
            resolveSheetLastCellNum(sheet));

        for (int pass = 1; pass <= MAX_TABLE_PASSES; pass++) {
            SheetProcessingState tablePassState = new SheetProcessingState();
            for (Row row : sheet) {
                if (row.getLastCellNum() <= 0) {
                    continue;
                }
                processRowTokens(sheet, row, context, options, warningCollector, tablePassState, false);
            }

            tablePassState.getAnchors().sort(
                Comparator
                    .comparingInt(PoiTableAnchor::rowIndex)
                    .reversed()
                    .thenComparing(
                        Comparator.comparingInt(PoiTableAnchor::colIndex).reversed()
                    )
            );

            if (tablePassState.getAnchors().isEmpty()) {
                log.trace("processSheetTokens() - tablePassComplete: sheetName={}, pass={}, anchorsFound=0",
                    sheetName, pass);
                break;
            }

            totalTablePasses++;
            totalTableTokensFound += tablePassState.getTableTokensFound();
            totalTableInsertions += tablePassState.getAnchors().size();

            log.trace("processSheetTokens() - tablePass: sheetName={}, pass={}, processedCells={}, tableTokensFound={}, tableInsertions={}",
                sheetName,
                pass,
                tablePassState.getProcessedCells(),
                tablePassState.getTableTokensFound(),
                tablePassState.getAnchors().size());

            for (PoiTableAnchor anchor : tablePassState.getAnchors()) {
                insertTableAtAnchor(sheet, anchor, options, warningCollector);
            }

            if (pass == MAX_TABLE_PASSES) {
                String location = cellLocation(sheet, 0, 0);
                log.warn("processSheetTokens() - tablePassLimitReached: sheetName={}, maxPasses={}", sheetName, MAX_TABLE_PASSES);
                warningCollector.add(
                    "TABLE_TOKEN_RECURSIVE",
                    "Table token expansion did not stabilize after " + MAX_TABLE_PASSES + " passes",
                    location
                );
            }
        }

        SheetProcessingState scalarPassState = new SheetProcessingState();
        for (Row row : sheet) {
            if (row.getLastCellNum() <= 0) {
                continue;
            }
            processRowTokens(sheet, row, context, options, warningCollector, scalarPassState, true);
        }

        log.trace("processSheetTokens() - end: sheetName={}, tablePasses={}, tableTokensFound={}, tableInsertions={}, scalarProcessedCells={}, scalarTokensApplied={}",
            sheetName,
            totalTablePasses,
            totalTableTokensFound,
            totalTableInsertions,
            scalarPassState.getProcessedCells(),
            scalarPassState.getScalarTokensApplied());
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
        SheetProcessingState state,
        boolean scalarPass
    ) {
        int rowIndex = row.getRowNum();
        log.trace("processRowTokens() - start: rowIndex={}, cellCount={}, scalarPass={}", rowIndex, row.getLastCellNum(), scalarPass);
        for (Cell cell : row) {
            processCellToken(sheet, row, cell, context, options, warningCollector, state, scalarPass);
        }
        log.trace("processRowTokens() - end: rowIndex={}, scalarPass={}", rowIndex, scalarPass);
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
        SheetProcessingState state,
        boolean scalarPass
    ) {
        int rowIndex = row.getRowNum();
        int colIndex = cell.getColumnIndex();
        String location = cellLocation(sheet, rowIndex, colIndex);
        state.incrementProcessedCells();

        if (scalarPass && isFormulaTokenCell(cell, location, warningCollector)) {
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

        if (!scalarPass) {
            return;
        }
        if (isScalarTableTokenCandidate(original, context, options)) {
            log.trace("processCellToken() - scalarSkippedForTableToken: location={}, tokenText={}", location, original);
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
     * Checks whether current token text should be treated as table token during scalar pass.
     *
     * <p>This guard prevents scalar phase from converting unresolved table placeholders to
     * stringified list values when table expansion is expected.
     */
    private boolean isScalarTableTokenCandidate(
        String original,
        Map<String, Object> context,
        GenerateOptions options
    ) {
        String exactToken = TokenResolver.getExactToken(original);
        String singleToken = TokenResolver.getSingleToken(original);
        String token = exactToken != null ? exactToken : singleToken;
        if (token == null || TokenResolver.isItemOrIndexToken(token)) {
            return false;
        }

        Object resolved = TokenResolver.resolvePath(context, token);
        if (resolved == null) {
            return false;
        }

        List<String> configuredColumnOrder = resolveConfiguredColumnOrder(context, token);
        if (options.rowsOnlyTableTokens()) {
            return toRowsOnlyTableRows(resolved, configuredColumnOrder) != null;
        }
        return TokenResolver.toTableRows(resolved) != null;
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

        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            Map<String, Object> values = rows.get(rowIndex);
            Row dataRow = getOrCreateRow(sheet, anchor.rowIndex() + rowIndex);
            dataRow.setHeight(anchor.baselineRowHeight());

            for (int columnIndex = 0; columnIndex < columns.size(); columnIndex++) {
                String column = columns.get(columnIndex);
                Cell cell = getOrCreateCell(dataRow, anchor.colIndex() + columnIndex);
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

}
