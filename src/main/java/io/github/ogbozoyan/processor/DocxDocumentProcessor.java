package io.github.ogbozoyan.processor;

import io.github.ogbozoyan.contract.RowBuilder;
import io.github.ogbozoyan.contract.TableBuilder;
import io.github.ogbozoyan.data.DocxTableAnchor;
import io.github.ogbozoyan.data.GenerateOptions;
import io.github.ogbozoyan.data.ParagraphTarget;
import io.github.ogbozoyan.data.ResolvedText;
import io.github.ogbozoyan.data.TemplateScanResult;
import io.github.ogbozoyan.exception.TemplateReadWriteException;
import io.github.ogbozoyan.util.TokenResolver;
import io.github.ogbozoyan.util.WarningCollector;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static io.github.ogbozoyan.helper.DocxHelper.buildColumnOrder;
import static io.github.ogbozoyan.helper.DocxHelper.cloneTableRow;
import static io.github.ogbozoyan.helper.DocxHelper.collectParagraphTargets;
import static io.github.ogbozoyan.helper.DocxHelper.getOrCreateFirstRow;
import static io.github.ogbozoyan.helper.DocxHelper.insertTableAtParagraphCursor;
import static io.github.ogbozoyan.helper.DocxHelper.removeParagraphFromContainer;
import static io.github.ogbozoyan.helper.DocxHelper.replaceParagraphText;
import static io.github.ogbozoyan.helper.DocxHelper.resetRowCells;
import static io.github.ogbozoyan.helper.DocxHelper.setCellText;
import static io.github.ogbozoyan.helper.DocxHelper.setCellTextPreservingParagraphStyle;
import static io.github.ogbozoyan.helper.DocxHelper.setHorizontalSpan;
import static io.github.ogbozoyan.helper.DocxHelper.writeTable;

/**
 * DOCX io.github.ogbozoyan.processor with scalar replacement and table-token insertion.
 *
 * <p>Unlike plain body-only approaches, this implementation recursively traverses
 * document body and nested table cells, so placeholders inside existing tables
 * are handled by the same algorithm as top-level placeholders.
 *
 * <p>Table token payload can be either classic {@code List<Map<String,Object>>}
 * (header + data rows) or declarative {@code io.github.ogbozoyan.contract.TableBuilder}
 * (explicit row/cell model with colspan and cell-level bold) or
 * {@code io.github.ogbozoyan.contract.RowBuilder}
 * (clone styled template row inside an existing DOCX table).
 *
 * <p>Example:
 * <pre>{@code
 * try (DocxDocumentProcessor io.github.ogbozoyan.processor = new DocxDocumentProcessor(bytes)) {
 *     io.github.ogbozoyan.processor.applyTemplateTokens(tokens, GenerateOptions.defaults(), warningCollector);
 *     byte[] output = io.github.ogbozoyan.processor.serialize();
 * }
 * }</pre>
 */
@Slf4j
public class DocxDocumentProcessor implements WorkbookProcessor {

    private final XWPFDocument document;

    /**
     * Creates io.github.ogbozoyan.processor and parses DOCX bytes.
     *
     * @param bytes source DOCX template bytes
     * @throws TemplateReadWriteException when document cannot be parsed
     */
    public DocxDocumentProcessor(byte[] bytes) {
        log.trace("DocxDocumentProcessor() - start: bytesLength={}", bytes == null ? null : bytes.length);
        try {
            this.document = new XWPFDocument(new ByteArrayInputStream(bytes));
            log.trace("DocxDocumentProcessor() - end: paragraphs={}", this.document.getParagraphs().size());
        } catch (Exception e) {
            log.error("DocxDocumentProcessor() - end with error: bytesLength={}", bytes == null ? null : bytes.length, e);
            throw new TemplateReadWriteException("Failed to read DOCX template", e);
        }
    }

    /**
     * Returns empty scan result because DOCX path currently performs direct apply phase.
     *
     * @return empty scan result
     */
    @Override
    public TemplateScanResult scan() {
        log.trace("scan() - start");
        TemplateScanResult result = new TemplateScanResult(List.of(), List.of());
        log.trace("scan() - end: markers={}, tokens={}", result.markers().size(), result.scalarTokens().size());
        return result;
    }

    /**
     * Applies scalar and table tokens to all collected DOCX paragraphs.
     *
     * <p>Table placeholders are collected as anchors first, then applied in reverse
     * order to keep cursor positions stable while body structure mutates.
     *
     * @param templateTokensMappings token map
     * @param options                generation options
     * @param warningCollector       collector for non-fatal warnings
     */
    @Override
    public void process(Map<String, Object> templateTokensMappings, GenerateOptions options, WarningCollector warningCollector) {
        Map<String, Object> context = templateTokensMappings == null ? Map.of() : templateTokensMappings;
        List<ParagraphTarget> paragraphTargets = collectParagraphTargets(document);
        log.debug("process() - start: tokenCount={}, paragraphs={}",
            context.size(),
            paragraphTargets.size());
        List<DocxTableAnchor> anchors = new ArrayList<>();
        int scalarReplacements = 0;

        for (ParagraphTarget paragraphTarget : paragraphTargets) {
            XWPFParagraph paragraph = paragraphTarget.paragraph();
            String text = paragraphTarget.text();
            if (text == null || text.isEmpty()) {
                continue;
            }

            String exactToken = TokenResolver.getExactToken(text);
            if (exactToken != null) {
                Object resolved = TokenResolver.resolvePath(context, exactToken);
                if (resolved instanceof TableBuilder
                    || resolved instanceof RowBuilder
                    || TokenResolver.isTableValue(resolved)) {
                    anchors.add(
                        new DocxTableAnchor(
                            paragraph,
                            exactToken,
                            resolved,
                            paragraphTarget.location(),
                            paragraphTarget.order()
                        )
                    );
                    continue;
                }
            }

            ResolvedText resolvedText = TokenResolver.resolve(
                text,
                context,
                options.missingValuePolicy(),
                warningCollector,
                paragraphTarget.location(),
                false
            );
            if (resolvedText.changed()) {
                replaceParagraphText(paragraph, resolvedText.value());
                scalarReplacements++;
            }
        }

        anchors.sort(
            (a, b) -> Integer.compare(b.order(), a.order())
        );
        for (DocxTableAnchor anchor : anchors) {
            insertTableAtParagraph(anchor, context, options, warningCollector);
        }
        log.trace("process() - end: tableInsertions={}, scalarReplacements={}", anchors.size(), scalarReplacements);
    }

    /**
     * Serializes modified DOCX document.
     *
     * @return generated DOCX bytes
     */
    @Override
    public byte[] serialize() {
        log.trace("serialize() - start");
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            document.write(outputStream);
            byte[] bytes = outputStream.toByteArray();
            log.trace("serialize() - end: bytesLength={}", bytes.length);
            return bytes;
        } catch (Exception e) {
            log.error("serialize() - end with error", e);
            throw new TemplateReadWriteException("Failed to serialize DOCX document", e);
        }
    }

    /**
     * Closes underlying XWPF document.
     */
    @Override
    public void close() {
        log.trace("close() - start");
        try {
            document.close();
            log.trace("close() - end: closed=true");
        } catch (Exception ignored) {
            log.trace("close() - end with warning: failedToClose=true");
            // no-op
        }
    }

    /**
     * Inserts table at placeholder paragraph and removes placeholder paragraph.
     *
     * @param anchor           insertion anchor
     * @param context          token context
     * @param options          generation options
     * @param warningCollector collector for non-fatal warnings
     */
    private void insertTableAtParagraph(
        DocxTableAnchor anchor,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        Object payload = anchor.tablePayload();
        if (payload instanceof RowBuilder rowBuilder) {
            insertTemplateRowsAtParagraph(anchor, rowBuilder, context, options, warningCollector);
            return;
        }
        if (payload instanceof TableBuilder builder) {
            insertDeclarativeTableAtParagraph(anchor, builder, context, options, warningCollector);
            return;
        }
        insertMapTableAtParagraph(anchor, payload, warningCollector);
    }

    private void insertMapTableAtParagraph(
        DocxTableAnchor anchor,
        Object payload,
        WarningCollector warningCollector
    ) {
        List<Map<String, Object>> rows = TokenResolver.toTableRows(payload);
        if (rows == null) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has invalid structure: " + anchor.token(), anchor.location());
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertMapTableAtParagraph() - end: inserted=false, reason=invalidRows");
            return;
        }

        log.trace("insertMapTableAtParagraph() - start: token={}, rowCount={}, location={}",
            anchor.token(), rows.size(), anchor.location());
        if (rows.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), anchor.location());
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertMapTableAtParagraph() - end: inserted=false, reason=emptyRows");
            return;
        }

        List<String> columns = buildColumnOrder(rows);
        if (columns.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), anchor.location());
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertMapTableAtParagraph() - end: inserted=false, reason=emptyColumns");
            return;
        }

        XmlCursor cursor = anchor.paragraph().getCTP().newCursor();
        XWPFTable table = insertTableAtParagraphCursor(anchor.paragraph(), cursor);
        writeTable(table, columns, rows);

        removeParagraphFromContainer(anchor.paragraph());
        log.trace("insertMapTableAtParagraph() - end: inserted=true, columns={}", columns.size());
    }

    private void insertDeclarativeTableAtParagraph(
        DocxTableAnchor anchor,
        TableBuilder builder,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        List<TableBuilder.Row> specRows = builder.rows();
        log.trace("insertDeclarativeTableAtParagraph() - start: token={}, rowCount={}, location={}",
            anchor.token(), specRows.size(), anchor.location());
        if (specRows.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Declarative table has no rows: " + anchor.token(), anchor.location());
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertDeclarativeTableAtParagraph() - end: inserted=false, reason=emptyRows");
            return;
        }

        int columnCount = builder.columnCount();
        if (columnCount < 1) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Declarative table has no columns: " + anchor.token(), anchor.location());
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertDeclarativeTableAtParagraph() - end: inserted=false, reason=emptyColumns");
            return;
        }

        XmlCursor cursor = anchor.paragraph().getCTP().newCursor();
        XWPFTable table = insertTableAtParagraphCursor(anchor.paragraph(), cursor);
        for (int rowIndex = 0; rowIndex < specRows.size(); rowIndex++) {
            TableBuilder.Row rowSpec = specRows.get(rowIndex);
            var row = rowIndex == 0
                ? getOrCreateFirstRow(table)
                : table.createRow();

            int cellCount = rowSpec.cells().size();
            int missingCells = columnCount - rowSpec.width();
            if (missingCells > 0) {
                cellCount += missingCells;
            }
            resetRowCells(row, cellCount);

            int physicalCell = 0;
            for (int colIndex = 0; colIndex < rowSpec.cells().size(); colIndex++) {
                TableBuilder.Cell cellSpec = rowSpec.cells().get(colIndex);
                String cellLocation = anchor.location() + "/row#" + rowIndex + "/cell#" + colIndex;
                ResolvedText resolvedText = TokenResolver.resolve(
                    cellSpec.text(),
                    context,
                    options.missingValuePolicy(),
                    warningCollector,
                    cellLocation,
                    false
                );
                setCellText(row.getCell(physicalCell), resolvedText.value(), cellSpec.bold());
                setHorizontalSpan(row.getCell(physicalCell), cellSpec.colSpan());
                physicalCell++;
            }

            for (int padIndex = 0; padIndex < missingCells; padIndex++) {
                setCellText(row.getCell(physicalCell), "", false);
                physicalCell++;
            }
        }

        removeParagraphFromContainer(anchor.paragraph());
        log.trace("insertDeclarativeTableAtParagraph() - end: inserted=true, columns={}", columnCount);
    }

    private void insertTemplateRowsAtParagraph(
        DocxTableAnchor anchor,
        RowBuilder rowBuilder,
        Map<String, Object> context,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        List<RowBuilder.Row> specRows = rowBuilder.rows();
        log.trace("insertTemplateRowsAtParagraph() - start: token={}, rowCount={}, location={}",
            anchor.token(), specRows.size(), anchor.location());
        if (specRows.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Template row payload has no rows: " + anchor.token(), anchor.location());
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertTemplateRowsAtParagraph() - end: inserted=false, reason=emptyRows");
            return;
        }

        if (!(anchor.paragraph().getBody() instanceof XWPFTableCell tableCell)) {
            warningCollector.add(
                "TABLE_TOKEN_INVALID",
                "Template row token must be placed inside DOCX table cell: " + anchor.token(),
                anchor.location()
            );
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertTemplateRowsAtParagraph() - end: inserted=false, reason=invalidAnchorContainer");
            return;
        }

        XWPFTableRow templateRow = tableCell.getTableRow();
        XWPFTable table = templateRow == null ? null : templateRow.getTable();
        if (templateRow == null || table == null) {
            warningCollector.add(
                "TABLE_TOKEN_INVALID",
                "Template row token has invalid DOCX table context: " + anchor.token(),
                anchor.location()
            );
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertTemplateRowsAtParagraph() - end: inserted=false, reason=missingTableContext");
            return;
        }

        int templateRowIndex = table.getRows().indexOf(templateRow);
        if (templateRowIndex < 0) {
            warningCollector.add(
                "TABLE_TOKEN_INVALID",
                "Template row token row is not attached to table: " + anchor.token(),
                anchor.location()
            );
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertTemplateRowsAtParagraph() - end: inserted=false, reason=detachedRow");
            return;
        }

        boolean truncated = false;
        for (int rowIndex = 0; rowIndex < specRows.size(); rowIndex++) {
            RowBuilder.Row rowSpec = specRows.get(rowIndex);
            XWPFTableRow insertedRow = cloneTableRow(table, templateRow, templateRowIndex + rowIndex);
            int cellCount = insertedRow.getTableCells().size();
            int writableCells = Math.min(cellCount, rowSpec.cells().size());
            if (rowSpec.cells().size() > cellCount) {
                truncated = true;
            }

            for (int cellIndex = 0; cellIndex < writableCells; cellIndex++) {
                String cellLocation = anchor.location() + "/row#" + rowIndex + "/cell#" + cellIndex;
                ResolvedText resolvedText = TokenResolver.resolve(
                    rowSpec.cells().get(cellIndex).text(),
                    context,
                    options.missingValuePolicy(),
                    warningCollector,
                    cellLocation,
                    false
                );
                XWPFTableCell targetCell = insertedRow.getCell(cellIndex);
                if (targetCell == null) {
                    truncated = true;
                    continue;
                }
                setCellTextPreservingParagraphStyle(targetCell, resolvedText.value());
            }
            for (int cellIndex = writableCells; cellIndex < cellCount; cellIndex++) {
                XWPFTableCell targetCell = insertedRow.getCell(cellIndex);
                if (targetCell != null) {
                    setCellTextPreservingParagraphStyle(targetCell, "");
                }
            }
        }

        table.removeRow(templateRowIndex + specRows.size());
        if (truncated) {
            warningCollector.add(
                "TABLE_TOKEN_INVALID",
                "Template row payload has more cells than template row allows: " + anchor.token(),
                anchor.location()
            );
        }
        log.trace("insertTemplateRowsAtParagraph() - end: inserted=true, rows={}", specRows.size());
    }


}
