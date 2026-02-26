package com.template.reportgenerator.processor;

import com.template.reportgenerator.contract.DocxTableAnchor;
import com.template.reportgenerator.contract.GenerateOptions;
import com.template.reportgenerator.contract.ResolvedText;
import com.template.reportgenerator.contract.TemplateScanResult;
import com.template.reportgenerator.contract.TokenOccurrence;
import com.template.reportgenerator.exception.TemplateReadWriteException;
import com.template.reportgenerator.util.TokenResolver;
import com.template.reportgenerator.util.WarningCollector;
import lombok.extern.slf4j.Slf4j;
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
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

/**
 * DOCX processor with scalar replacement and table-token insertion.
 */
@Slf4j
public class DocxDocumentProcessor implements WorkbookProcessor {

    private final XWPFDocument document;

    public DocxDocumentProcessor(byte[] bytes) {
        log.info("DocxDocumentProcessor() - start: bytesLength={}", bytes == null ? null : bytes.length);
        try {
            this.document = new XWPFDocument(new ByteArrayInputStream(bytes));
            log.info("DocxDocumentProcessor() - end: paragraphs={}", this.document.getParagraphs().size());
        } catch (Exception e) {
            log.error("DocxDocumentProcessor() - end with error: bytesLength={}", bytes == null ? null : bytes.length, e);
            throw new TemplateReadWriteException("Failed to read DOCX template", e);
        }
    }

    @Override
    public TemplateScanResult scan() {
        log.info("scan() - start");
        TemplateScanResult result = new TemplateScanResult(List.of(), List.<TokenOccurrence>of());
        log.info("scan() - end: markers={}, tokens={}", result.markers().size(), result.scalarTokens().size());
        return result;
    }

    @Override
    public void applyTemplateTokens(Map<String, Object> templateTokens, GenerateOptions options, WarningCollector warningCollector) {
        log.info("applyTemplateTokens() - start: tokenCount={}, paragraphs={}",
            templateTokens == null ? null : templateTokens.size(),
            document.getParagraphs().size());
        List<DocxTableAnchor> anchors = new ArrayList<>();
        int scalarReplacements = 0;

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            String text = paragraph.getText();
            if (text == null || text.isEmpty()) {
                continue;
            }

            String exactToken = TokenResolver.getExactToken(text);
            if (exactToken != null) {
                Object resolved = TokenResolver.resolvePath(templateTokens, exactToken);
                if (TokenResolver.isTableValue(resolved)) {
                    List<Map<String, Object>> tableRows = TokenResolver.toTableRows(resolved);
                    if (tableRows == null) {
                        warningCollector.add("TABLE_TOKEN_INVALID", "Table token has invalid structure: " + exactToken, "docx:paragraph");
                    } else {
                        anchors.add(new DocxTableAnchor(paragraph, exactToken, tableRows));
                    }
                    continue;
                }
            }

            ResolvedText resolvedText = TokenResolver.resolve(
                text,
                templateTokens,
                options.missingValuePolicy(),
                warningCollector,
                "docx:paragraph",
                false
            );
            if (resolvedText.changed()) {
                replaceParagraphText(paragraph, resolvedText.value());
                scalarReplacements++;
            }
        }

        anchors.sort(Comparator.comparingInt((DocxTableAnchor anchor) -> document.getPosOfParagraph(anchor.paragraph())).reversed());
        for (DocxTableAnchor anchor : anchors) {
            insertTableAtParagraph(anchor, warningCollector);
        }
        log.info("applyTemplateTokens() - end: tableInsertions={}, scalarReplacements={}", anchors.size(), scalarReplacements);
    }

    @Override
    public byte[] serialize() {
        log.info("serialize() - start");
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            document.write(outputStream);
            byte[] bytes = outputStream.toByteArray();
            log.info("serialize() - end: bytesLength={}", bytes.length);
            return bytes;
        } catch (Exception e) {
            log.error("serialize() - end with error", e);
            throw new TemplateReadWriteException("Failed to serialize DOCX document", e);
        }
    }

    @Override
    public void close() {
        log.info("close() - start");
        try {
            document.close();
            log.info("close() - end: closed=true");
        } catch (Exception ignored) {
            log.warn("close() - end with warning: failedToClose=true");
            // no-op
        }
    }

    private void insertTableAtParagraph(DocxTableAnchor anchor, WarningCollector warningCollector) {
        log.info("insertTableAtParagraph() - start: token={}, rowCount={}", anchor.token(), anchor.rows().size());
        List<Map<String, Object>> rows = anchor.rows();
        if (rows.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), "docx:paragraph");
            replaceParagraphText(anchor.paragraph(), "");
            log.info("insertTableAtParagraph() - end: inserted=false, reason=emptyRows");
            return;
        }

        List<String> columns = buildColumnOrder(rows);
        if (columns.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), "docx:paragraph");
            replaceParagraphText(anchor.paragraph(), "");
            log.info("insertTableAtParagraph() - end: inserted=false, reason=emptyColumns");
            return;
        }

        XmlCursor cursor = anchor.paragraph().getCTP().newCursor();
        XWPFTable table = document.insertNewTbl(cursor);
        writeTable(table, columns, rows);

        int paragraphPos = document.getPosOfParagraph(anchor.paragraph());
        if (paragraphPos >= 0) {
            document.removeBodyElement(paragraphPos);
        }
        log.info("insertTableAtParagraph() - end: inserted=true, columns={}", columns.size());
    }

    private void writeTable(XWPFTable table, List<String> columns, List<Map<String, Object>> rows) {
        XWPFTableRow headerRow = table.getRow(0);
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

    private void ensureCells(XWPFTableRow row, int count) {
        while (row.getTableCells().size() < count) {
            row.addNewTableCell();
        }
    }

    private void setCellText(XWPFTableCell cell, String value) {
        int paragraphCount = cell.getParagraphs().size();
        for (int i = paragraphCount - 1; i >= 0; i--) {
            cell.removeParagraph(i);
        }
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(value == null ? "" : value);
    }

    private List<String> buildColumnOrder(List<Map<String, Object>> rows) {
        LinkedHashSet<String> ordered = new LinkedHashSet<>();
        if (!rows.isEmpty()) {
            ordered.addAll(rows.getFirst().keySet());
        }
        for (Map<String, Object> row : rows) {
            ordered.addAll(row.keySet());
        }
        return List.copyOf(ordered);
    }

    private void replaceParagraphText(XWPFParagraph paragraph, String value) {
        int runs = paragraph.getRuns().size();
        for (int i = runs - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
        XWPFRun run = paragraph.createRun();
        run.setText(value == null ? "" : value);
    }

}
