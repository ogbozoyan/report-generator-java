package com.template.reportgenerator.processor;

import com.template.reportgenerator.dto.BlockRegion;
import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.ResolvedText;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.dto.TokenOccurrence;
import com.template.reportgenerator.exception.TemplateReadWriteException;
import com.template.reportgenerator.util.TokenResolver;
import com.template.reportgenerator.util.WarningCollector;
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
public class DocxDocumentProcessor implements WorkbookProcessor {

    private final XWPFDocument document;

    public DocxDocumentProcessor(byte[] bytes) {
        try {
            this.document = new XWPFDocument(new ByteArrayInputStream(bytes));
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to read DOCX template", e);
        }
    }

    @Override
    public TemplateScanResult scan() {
        return new TemplateScanResult(List.of(), List.<TokenOccurrence>of());
    }

    @Override
    public void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector) {
        List<TableAnchor> anchors = new ArrayList<>();

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            String text = paragraph.getText();
            if (text == null || text.isEmpty()) {
                continue;
            }

            String exactToken = TokenResolver.getExactToken(text);
            if (exactToken != null) {
                Object resolved = TokenResolver.resolvePath(scalars, exactToken);
                if (TokenResolver.isTableValue(resolved)) {
                    List<Map<String, Object>> tableRows = TokenResolver.toTableRows(resolved);
                    if (tableRows == null) {
                        warningCollector.add("TABLE_TOKEN_INVALID", "Table token has invalid structure: " + exactToken, "docx:paragraph");
                    } else {
                        anchors.add(new TableAnchor(paragraph, exactToken, tableRows));
                    }
                    continue;
                }
            }

            ResolvedText resolvedText = TokenResolver.resolve(
                text,
                scalars,
                options.missingValuePolicy(),
                warningCollector,
                "docx:paragraph",
                false
            );
            if (resolvedText.changed()) {
                replaceParagraphText(paragraph, resolvedText.value());
            }
        }

        anchors.sort(Comparator.comparingInt((TableAnchor anchor) -> document.getPosOfParagraph(anchor.paragraph())).reversed());
        for (TableAnchor anchor : anchors) {
            insertTableAtParagraph(anchor, warningCollector);
        }
    }

    @Override
    public void expandTableBlocks(
        List<BlockRegion> tableBlocks,
        ReportData data,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        // legacy block expansion is not used anymore.
    }

    @Override
    public void expandColumnBlocks(
        List<BlockRegion> columnBlocks,
        ReportData data,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        // legacy block expansion is not used anymore.
    }

    @Override
    public void clearMarkers(List<BlockRegion> blockRegions) {
        // marker clearing is not required with token-only pipeline.
    }

    @Override
    public void recalculateFormulas(GenerateOptions options) {
        // no formulas in DOCX text flow.
    }

    @Override
    public byte[] serialize() {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            document.write(outputStream);
            return outputStream.toByteArray();
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to serialize DOCX document", e);
        }
    }

    @Override
    public void close() {
        try {
            document.close();
        } catch (Exception ignored) {
            // no-op
        }
    }

    private void insertTableAtParagraph(TableAnchor anchor, WarningCollector warningCollector) {
        List<Map<String, Object>> rows = anchor.rows();
        if (rows.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), "docx:paragraph");
            replaceParagraphText(anchor.paragraph(), "");
            return;
        }

        List<String> columns = buildColumnOrder(rows);
        if (columns.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), "docx:paragraph");
            replaceParagraphText(anchor.paragraph(), "");
            return;
        }

        XmlCursor cursor = anchor.paragraph().getCTP().newCursor();
        XWPFTable table = document.insertNewTbl(cursor);
        writeTable(table, columns, rows);

        int paragraphPos = document.getPosOfParagraph(anchor.paragraph());
        if (paragraphPos >= 0) {
            document.removeBodyElement(paragraphPos);
        }
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
            ordered.addAll(rows.get(0).keySet());
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

    private record TableAnchor(XWPFParagraph paragraph, String token, List<Map<String, Object>> rows) {
    }
}
