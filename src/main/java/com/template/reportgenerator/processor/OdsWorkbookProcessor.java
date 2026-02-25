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
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.doc.table.OdfTableColumn;
import org.odftoolkit.odfdom.doc.table.OdfTableRow;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import javax.xml.parsers.DocumentBuilderFactory;

/**
 * Spreadsheet processor for ODS format based on ODFDOM.
 * <p>
 * Legacy TABLE/COL DSL expansion is intentionally disabled. Table insertion is
 * based on exact-placeholder tokens where token value is {@code List<Map<...>>}.
 */
@Slf4j
public class OdsWorkbookProcessor implements WorkbookProcessor {

    private static final String TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

    private final OdfSpreadsheetDocument document;
    private final byte[] sourceBytes;

    public OdsWorkbookProcessor(byte[] bytes) {
        log.info("Initializing ODS workbook processor with {} bytes", bytes.length);
        try {
            this.sourceBytes = bytes;
            this.document = OdfSpreadsheetDocument.loadDocument(new ByteArrayInputStream(bytes));
            log.info("Successfully loaded ODS document with {} sheets",
                document.getTableList(false).size());
        } catch (Exception e) {
            log.error("Failed to read ODS template: {} bytes", bytes.length, e);
            throw new TemplateReadWriteException("Failed to read ODS template", e);
        }
    }

    @Override
    public TemplateScanResult scan() {
        log.info("Starting ODS template scanning");
        TemplateScanResult result = TemplateScanner.scanOds(document);
        log.info("ODS template scan completed - found tokens across sheets {}", result);
        return result;
    }

    @Override
    public void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector) {
        Map<String, Object> context = scalars == null ? Map.of() : scalars;
        List<OdfTable> sheets = document.getTableList(false);

        log.info("Applying scalar tokens to {} sheets with context size: {}", sheets.size(), context.size());
        log.info("Processing options: missingValuePolicy={}, zoneId={}", options.missingValuePolicy(), options.zoneId());

        for (int sheetIndex = 0; sheetIndex < sheets.size(); sheetIndex++) {
            OdfTable sheet = sheets.get(sheetIndex);
            String sheetName = sheet.getTableName() == null ? ("Sheet" + sheetIndex) : sheet.getTableName();
            List<CellReference> tokenCells = collectTokenCellsFromSource(sheetIndex);
            log.info(
                "Processing sheet '{}' ({}): logical size {}x{}, token cells {}",
                sheetName, sheetIndex, sheet.getRowCount(), sheet.getColumnCount(), tokenCells.size()
            );

            List<TableAnchor> anchors = new ArrayList<>();
            int tableTokensFound = 0;
            int scalarTokensApplied = 0;

            for (CellReference ref : tokenCells) {
                int rowIndex = ref.rowIndex();
                int colIndex = ref.colIndex();
                OdfTableCell cell = sheet.getCellByPosition(colIndex, rowIndex);
                String location = cellLocation(sheetName, rowIndex, colIndex);
                String original = ref.sourceText();
                if (original == null || original.isEmpty()) {
                    original = cell.getStringValue();
                }

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
                            tableTokensFound++;
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
                    original,
                    context,
                    options,
                    warningCollector,
                    location
                );
                scalarTokensApplied++;
            }

            anchors.sort(Comparator.comparingInt(TableAnchor::rowIndex).reversed()
                .thenComparing(Comparator.comparingInt(TableAnchor::colIndex).reversed()));

            log.info(
                "Sheet '{}': found {} table tokens, applied {} scalar tokens, inserting {} tables",
                sheetName, tableTokensFound, scalarTokensApplied, anchors.size()
            );

            for (TableAnchor anchor : anchors) {
                insertTableAtAnchor(sheet, sheetName, anchor, options, warningCollector);
            }
        }
        log.info("Completed scalar token application across all sheets");
    }

    @Override
    public byte[] serialize() {
        log.info("Serializing ODS document");
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            document.save(outputStream);
            byte[] result = outputStream.toByteArray();
            log.info("Successfully serialized ODS document: {} bytes", result.length);
            return result;
        } catch (Exception e) {
            log.error("Failed to serialize ODS document", e);
            throw new TemplateReadWriteException("Failed to serialize ODS document", e);
        }
    }

    @Override
    public void close() {
        log.info("Closing ODS workbook processor");
        try {
            document.close();
            log.info("ODS document closed successfully");
        } catch (Exception e) {
            log.warn("Error closing ODS document", e);
        }
    }

    private void applyTokenToCell(
        OdfTableCell cell,
        String original,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector,
        String location
    ) {
        String exactToken = TokenResolver.getExactToken(original);

        if (exactToken != null && !TokenResolver.isItemOrIndexToken(exactToken)) {
            Object resolved = TokenResolver.resolvePath(context, exactToken);
            if (resolved == null) {
                log.info("Handling missing exact token {} at {}", exactToken, location);
                handleMissingExactToken(cell, exactToken, options.missingValuePolicy(), warningCollector, location);
                return;
            }
            log.info("Writing resolved value for token {} at {}: {}", exactToken, location, resolved);
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
            log.info("Updating cell {} with resolved text: {} -> {}",
                location, original, resolvedText.value());
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

        log.info("Inserting table '{}' at {} with {} rows", anchor.token(), location, rows.size());

        if (rows.isEmpty()) {
            log.warn("Empty table token {} at {}", anchor.token(), location);
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor);
            ValueWriter.writeOdsValue(anchorCell, null, options.zoneId());
            return;
        }

        List<String> columns = buildColumnOrder(rows);
        if (columns.isEmpty()) {
            log.warn("Table token with no columns {} at {}", anchor.token(), location);
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), location);
            applyBaselineStyle(anchorCell, anchor);
            ValueWriter.writeOdsValue(anchorCell, null, options.zoneId());
            return;
        }

        log.info("Table '{}' structure: {} columns [{}], {} rows",
            anchor.token(), columns.size(), String.join(",", columns), rows.size());

        int dataRowCount = rows.size();
        if (dataRowCount > 0) {
            log.info("Inserting {} rows before row {}", dataRowCount, anchor.rowIndex() + 1);
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
        log.info("Completed table insertion for '{}' at {}", anchor.token(), location);
    }

    private void handleMissingExactToken(
        OdfTableCell cell,
        String token,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location
    ) {
        log.info("Handling missing token '{}' with policy {} at {}", token, policy, location);
        switch (policy) {
            case EMPTY_AND_LOG -> {
                log.info("Setting empty value for missing token {} at {}", token, location);
                warningCollector.add("MISSING_TOKEN", "Token not found: " + token, location);
                cell.setStringValue("");
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
        OdfTable table,
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

            long desiredWidth = calculateDesiredWidth(maxLength);
            OdfTableColumn target = table.getColumnByIndex(startColumnIndex + c);
            long currentWidth = target.getWidth();

            log.info("Column '{}': max content length={}, current width={}, desired width={}",
                column, maxLength, currentWidth, desiredWidth);

            if (currentWidth < desiredWidth) {
                log.info("Resizing column '{}' from {} to {}", column, currentWidth, desiredWidth);
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

    private List<CellReference> collectTokenCellsFromSource(int tableIndex) {
        try (ZipInputStream zipInputStream = new ZipInputStream(new ByteArrayInputStream(sourceBytes))) {
            ZipEntry entry;
            while ((entry = zipInputStream.getNextEntry()) != null) {
                if (!"content.xml".equals(entry.getName())) {
                    continue;
                }

                byte[] xmlBytes = zipInputStream.readAllBytes();
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                factory.setNamespaceAware(true);
                Document document = factory.newDocumentBuilder().parse(new ByteArrayInputStream(xmlBytes));
                NodeList tableNodes = document.getElementsByTagNameNS(TABLE_NS, "table");
                if (tableIndex < 0 || tableIndex >= tableNodes.getLength()) {
                    return List.of();
                }

                Element tableElement = (Element) tableNodes.item(tableIndex);
                return extractTokenCellsFromTableElement(tableElement);
            }
            return List.of();
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to inspect ODS content.xml for tokens", e);
        }
    }

    private List<CellReference> extractTokenCellsFromTableElement(Element tableElement) {
        List<CellReference> refs = new ArrayList<>();
        int rowIndex = 0;

        for (Node child = tableElement.getFirstChild(); child != null; child = child.getNextSibling()) {
            if (!(child instanceof Element rowElement)) {
                continue;
            }
            if (!TABLE_NS.equals(rowElement.getNamespaceURI()) || !"table-row".equals(rowElement.getLocalName())) {
                continue;
            }

            int rowRepeat = parsePositiveInt(rowElement.getAttributeNS(TABLE_NS, "number-rows-repeated"), 1);
            int colIndex = 0;

            for (Node rowChild = rowElement.getFirstChild(); rowChild != null; rowChild = rowChild.getNextSibling()) {
                if (!(rowChild instanceof Element cellElement)) {
                    continue;
                }
                if (!TABLE_NS.equals(cellElement.getNamespaceURI())) {
                    continue;
                }

                String localName = cellElement.getLocalName();
                if (!"table-cell".equals(localName) && !"covered-table-cell".equals(localName)) {
                    continue;
                }

                int colRepeat = parsePositiveInt(cellElement.getAttributeNS(TABLE_NS, "number-columns-repeated"), 1);
                if ("table-cell".equals(localName)) {
                    String text = cellElement.getTextContent();
                    String formula = cellElement.getAttributeNS(TABLE_NS, "formula");
                    if (TokenResolver.hasTokens(text) || TokenResolver.hasTokens(formula)) {
                        refs.add(new CellReference(rowIndex, colIndex, text));
                    }
                }

                colIndex += colRepeat;
            }

            rowIndex += rowRepeat;
        }
        return refs;
    }

    private int parsePositiveInt(String value, int fallback) {
        if (value == null || value.isBlank()) {
            return fallback;
        }
        try {
            int parsed = Integer.parseInt(value);
            return parsed > 0 ? parsed : fallback;
        } catch (NumberFormatException ignored) {
            return fallback;
        }
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

    private record CellReference(int rowIndex, int colIndex, String sourceText) {
    }
}
