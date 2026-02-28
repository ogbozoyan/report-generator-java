package io.github.ogbozoyan.processor;

import io.github.ogbozoyan.contract.DocxTableAnchor;
import io.github.ogbozoyan.contract.GenerateOptions;
import io.github.ogbozoyan.contract.ResolvedText;
import io.github.ogbozoyan.contract.TemplateScanResult;
import io.github.ogbozoyan.exception.TemplateReadWriteException;
import io.github.ogbozoyan.util.TokenResolver;
import io.github.ogbozoyan.util.WarningCollector;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

/**
 * DOCX io.github.ogbozoyan.processor with scalar replacement and table-token insertion.
 *
 * <p>Unlike plain body-only approaches, this implementation recursively traverses
 * document body and nested table cells, so placeholders inside existing tables
 * are handled by the same algorithm as top-level placeholders.
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
     * @param templateTokens   token map
     * @param options          generation options
     * @param warningCollector collector for non-fatal warnings
     */
    @Override
    public void applyTemplateTokens(Map<String, Object> templateTokens, GenerateOptions options, WarningCollector warningCollector) {
        List<ParagraphTarget> paragraphTargets = collectParagraphTargets();
        log.debug("applyTemplateTokens() - start: tokenCount={}, paragraphs={}",
            templateTokens == null ? null : templateTokens.size(),
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
                Object resolved = TokenResolver.resolvePath(templateTokens, exactToken);
                if (TokenResolver.isTableValue(resolved)) {
                    List<Map<String, Object>> tableRows = TokenResolver.toTableRows(resolved);
                    if (tableRows == null) {
                        warningCollector.add("TABLE_TOKEN_INVALID", "Table token has invalid structure: " + exactToken, paragraphTarget.location());
                    } else {
                        anchors.add(new DocxTableAnchor(
                            paragraph,
                            exactToken,
                            tableRows,
                            paragraphTarget.location(),
                            paragraphTarget.order()
                        ));
                    }
                    continue;
                }
            }

            ResolvedText resolvedText = TokenResolver.resolve(
                text,
                templateTokens,
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

        anchors.sort((a, b) -> Integer.compare(b.order(), a.order()));
        for (DocxTableAnchor anchor : anchors) {
            insertTableAtParagraph(anchor, warningCollector);
        }
        log.trace("applyTemplateTokens() - end: tableInsertions={}, scalarReplacements={}", anchors.size(), scalarReplacements);
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
     * @param warningCollector collector for non-fatal warnings
     */
    private void insertTableAtParagraph(DocxTableAnchor anchor, WarningCollector warningCollector) {
        log.trace("insertTableAtParagraph() - start: token={}, rowCount={}, location={}",
            anchor.token(), anchor.rows().size(), anchor.location());
        List<Map<String, Object>> rows = anchor.rows();
        if (rows.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), anchor.location());
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertTableAtParagraph() - end: inserted=false, reason=emptyRows");
            return;
        }

        List<String> columns = buildColumnOrder(rows);
        if (columns.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), anchor.location());
            replaceParagraphText(anchor.paragraph(), "");
            log.trace("insertTableAtParagraph() - end: inserted=false, reason=emptyColumns");
            return;
        }

        XmlCursor cursor = anchor.paragraph().getCTP().newCursor();
        XWPFTable table = insertTableAtParagraphCursor(anchor.paragraph(), cursor);
        writeTable(table, columns, rows);

        removeParagraphFromContainer(anchor.paragraph());
        log.trace("insertTableAtParagraph() - end: inserted=true, columns={}", columns.size());
    }

    /**
     * Writes header and data rows into newly created DOCX table.
     *
     * @param table   destination table
     * @param columns ordered columns
     * @param rows    table payload rows
     */
    private void writeTable(XWPFTable table, List<String> columns, List<Map<String, Object>> rows) {
        XWPFTableRow headerRow = getOrCreateFirstRow(table);
        ensureCells(headerRow, columns.size());
        for (int c = 0; c < columns.size(); c++) {
            setCellText(headerRow.getCell(c), columns.get(c));
        }

        for (Map<String, Object> row : rows) {
            XWPFTableRow dataRow = table.createRow();
            ensureCells(dataRow, columns.size());
            for (int c = 0; c < columns.size(); c++) {
                Object value = row.get(columns.get(c));
                setCellText(dataRow.getCell(c), value == null ? "" : String.valueOf(value));
            }
        }
    }

    /**
     * Ensures row has at least requested number of cells.
     *
     * @param row   table row
     * @param count minimum cell count
     */
    private void ensureCells(XWPFTableRow row, int count) {
        while (row.getTableCells().size() < count) {
            row.addNewTableCell();
        }
    }

    /**
     * Returns existing first row or creates one for table header.
     *
     * @param table destination table
     * @return first row
     */
    private XWPFTableRow getOrCreateFirstRow(XWPFTable table) {
        XWPFTableRow headerRow = table.getRow(0);
        if (headerRow != null) {
            return headerRow;
        }
        headerRow = table.insertNewTableRow(0);
        if (headerRow != null) {
            if (headerRow.getCell(0) == null) {
                headerRow.createCell();
            }
            return headerRow;
        }
        throw new TemplateReadWriteException("Failed to create header row for DOCX table insertion");
    }

    /**
     * Replaces cell paragraphs with single paragraph containing provided value.
     *
     * @param cell  destination cell
     * @param value text value
     */
    private void setCellText(XWPFTableCell cell, String value) {
        int paragraphCount = cell.getParagraphs().size();
        for (int i = paragraphCount - 1; i >= 0; i--) {
            cell.removeParagraph(i);
        }
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(value == null ? "" : value);
    }

    /**
     * Builds stable column order: first-row keys, then new keys in encounter order.
     *
     * @param rows table rows
     * @return ordered columns
     */
    private List<String> buildColumnOrder(List<Map<String, Object>> rows) {
        LinkedHashSet<String> ordered = new LinkedHashSet<>();
        if (!rows.isEmpty()) {
            ordered.addAll(rows.get(0).keySet());
        }
        for (Map<String, Object> row : rows) {
            ordered.addAll(row.keySet());
        }
        return List.copyOf(ordered);
    }

    /**
     * Replaces paragraph content with a single run.
     *
     * @param paragraph paragraph to rewrite
     * @param value     replacement text
     */
    private void replaceParagraphText(XWPFParagraph paragraph, String value) {
        int runs = paragraph.getRuns().size();
        for (int i = runs - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
        XWPFRun run = paragraph.createRun();
        run.setText(value == null ? "" : value);
    }

    /**
     * Collects every paragraph from document body and nested table cells.
     *
     * @return ordered paragraph targets
     */
    private List<ParagraphTarget> collectParagraphTargets() {
        List<ParagraphTarget> targets = new ArrayList<>();
        int[] order = {0};
        collectParagraphTargetsFromBody(document, "docx:body", order, targets);
        return targets;
    }

    /**
     * Recursively traverses body elements and records paragraph targets.
     *
     * @param body           body container (document or table cell)
     * @param locationPrefix diagnostic location prefix
     * @param order          mutable traversal counter
     * @param targets        output list
     */
    private void collectParagraphTargetsFromBody(
        IBody body,
        String locationPrefix,
        int[] order,
        List<ParagraphTarget> targets
    ) {
        List<IBodyElement> bodyElements = body.getBodyElements();
        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement bodyElement = bodyElements.get(i);
            if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                targets.add(new ParagraphTarget(
                    paragraph,
                    paragraph.getText(),
                    locationPrefix + "/p#" + i,
                    order[0]++
                ));
                continue;
            }
            if (bodyElement.getElementType() == BodyElementType.TABLE) {
                XWPFTable table = (XWPFTable) bodyElement;
                collectParagraphTargetsFromTable(table, locationPrefix + "/tbl#" + i, order, targets);
            }
        }
    }

    /**
     * Traverses paragraphs inside table cells recursively.
     *
     * @param table          source table
     * @param locationPrefix diagnostic location prefix
     * @param order          mutable traversal counter
     * @param targets        output list
     */
    private void collectParagraphTargetsFromTable(
        XWPFTable table,
        String locationPrefix,
        int[] order,
        List<ParagraphTarget> targets
    ) {
        List<XWPFTableRow> rows = table.getRows();
        for (int r = 0; r < rows.size(); r++) {
            XWPFTableRow row = rows.get(r);
            List<XWPFTableCell> cells = row.getTableCells();
            for (int c = 0; c < cells.size(); c++) {
                XWPFTableCell cell = cells.get(c);
                collectParagraphTargetsFromBody(
                    cell,
                    locationPrefix + "/r#" + r + "/c#" + c,
                    order,
                    targets
                );
            }
        }
    }

    /**
     * Inserts table at paragraph cursor according to paragraph container type.
     *
     * @param paragraph anchor paragraph
     * @param cursor    insertion cursor
     * @return newly inserted table
     */
    private XWPFTable insertTableAtParagraphCursor(XWPFParagraph paragraph, XmlCursor cursor) {
        IBody body = paragraph.getBody();
        if (body instanceof XWPFDocument doc) {
            return doc.insertNewTbl(cursor);
        }
        if (body instanceof XWPFTableCell cell) {
            return cell.insertNewTbl(cursor);
        }
        throw new TemplateReadWriteException(
            "Unsupported DOCX paragraph container for table insertion: " + body.getClass().getName()
        );
    }

    /**
     * Removes placeholder paragraph from its container after table insertion.
     *
     * @param paragraph placeholder paragraph
     */
    private void removeParagraphFromContainer(XWPFParagraph paragraph) {
        IBody body = paragraph.getBody();
        if (body instanceof XWPFDocument doc) {
            int paragraphPos = doc.getPosOfParagraph(paragraph);
            if (paragraphPos >= 0) {
                doc.removeBodyElement(paragraphPos);
                return;
            }
        } else if (body instanceof XWPFTableCell cell) {
            int paragraphPos = cell.getParagraphs().indexOf(paragraph);
            if (paragraphPos >= 0) {
                cell.removeParagraph(paragraphPos);
                return;
            }
        }
        replaceParagraphText(paragraph, "");
    }

    /**
     * Immutable paragraph processing target.
     *
     * @param paragraph paragraph object
     * @param text      paragraph plain text
     * @param location  diagnostic location
     * @param order     traversal order
     */
    private record ParagraphTarget(XWPFParagraph paragraph, String text, String location, int order) {
    }

}
