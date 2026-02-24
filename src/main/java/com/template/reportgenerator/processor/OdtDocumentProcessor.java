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
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.dom.element.text.TextPElement;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

/**
 * ODT processor with scalar replacement and table-token insertion.
 */
public class OdtDocumentProcessor implements WorkbookProcessor {

    private static final String TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    private final OdfTextDocument document;

    public OdtDocumentProcessor(byte[] bytes) {
        try {
            this.document = OdfTextDocument.loadDocument(new ByteArrayInputStream(bytes));
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to read ODT template", e);
        }
    }

    @Override
    public TemplateScanResult scan() {
        return new TemplateScanResult(List.of(), List.<TokenOccurrence>of());
    }

    @Override
    public void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector) {
        List<TextPElement> paragraphs = collectParagraphs();
        List<TableAnchor> anchors = new ArrayList<>();

        for (TextPElement paragraph : paragraphs) {
            String text = paragraph.getTextContent();
            if (text == null || text.isEmpty()) {
                continue;
            }

            String exactToken = TokenResolver.getExactToken(text);
            if (exactToken != null) {
                Object resolved = TokenResolver.resolvePath(scalars, exactToken);
                if (TokenResolver.isTableValue(resolved)) {
                    List<Map<String, Object>> rows = TokenResolver.toTableRows(resolved);
                    if (rows == null) {
                        warningCollector.add("TABLE_TOKEN_INVALID", "Table token has invalid structure: " + exactToken, "odt:paragraph");
                    } else {
                        anchors.add(new TableAnchor(paragraph, exactToken, rows));
                    }
                    continue;
                }
            }

            ResolvedText resolvedText = TokenResolver.resolve(
                text,
                scalars,
                options.missingValuePolicy(),
                warningCollector,
                "odt:paragraph",
                false
            );
            if (resolvedText.changed()) {
                paragraph.setTextContent(resolvedText.value());
            }
        }

        anchors.sort(Comparator.comparingInt((TableAnchor anchor) -> anchor.paragraph().countPrecedingSiblingElements()).reversed());
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
        // no formulas in ODT text flow.
    }

    @Override
    public byte[] serialize() {
        try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.save(output);
            return output.toByteArray();
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to serialize ODT document", e);
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
        if (anchor.rows().isEmpty()) {
            warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + anchor.token(), "odt:paragraph");
            anchor.paragraph().setTextContent("");
            return;
        }

        List<String> columns = buildColumnOrder(anchor.rows());
        if (columns.isEmpty()) {
            warningCollector.add("TABLE_TOKEN_INVALID", "Table token has no columns: " + anchor.token(), "odt:paragraph");
            anchor.paragraph().setTextContent("");
            return;
        }

        try {
            OdfTable table = OdfTable.newTable(document, anchor.rows().size() + 1, columns.size());
            for (int c = 0; c < columns.size(); c++) {
                table.getCellByPosition(c, 0).setStringValue(columns.get(c));
            }

            for (int r = 0; r < anchor.rows().size(); r++) {
                Map<String, Object> row = anchor.rows().get(r);
                for (int c = 0; c < columns.size(); c++) {
                    Object value = row.get(columns.get(c));
                    table.getCellByPosition(c, r + 1).setStringValue(value == null ? "" : String.valueOf(value));
                }
            }

            Node tableNode = table.getOdfElement();
            Node parent = anchor.paragraph().getParentNode();
            parent.insertBefore(tableNode, anchor.paragraph());
            parent.removeChild(anchor.paragraph());
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to insert table into ODT document", e);
        }
    }

    private List<TextPElement> collectParagraphs() {
        try {
            NodeList nodeList = document.getContentDom().getElementsByTagNameNS(TEXT_NS, "p");
            List<TextPElement> result = new ArrayList<>(nodeList.getLength());
            for (int i = 0; i < nodeList.getLength(); i++) {
                Node node = nodeList.item(i);
                if (node instanceof TextPElement paragraph) {
                    result.add(paragraph);
                }
            }
            return result;
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to read ODT paragraphs", e);
        }
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

    private record TableAnchor(TextPElement paragraph, String token, List<Map<String, Object>> rows) {
    }
}
