package com.template.reportgenerator.processor;

import com.template.reportgenerator.contract.GenerateOptions;
import com.template.reportgenerator.contract.TemplateScanResult;
import com.template.reportgenerator.contract.TokenOccurrence;
import com.template.reportgenerator.exception.TemplateReadWriteException;
import com.template.reportgenerator.util.TokenResolver;
import com.template.reportgenerator.util.WarningCollector;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

/**
 * Basic DOC processor based on HWPF.
 * <p>
 * Table tokens are rendered as tab/newline separated text blocks.
 */
@Slf4j
public class DocDocumentProcessor implements WorkbookProcessor {

    private final HWPFDocument document;

    public DocDocumentProcessor(byte[] bytes) {
        log.info("DocDocumentProcessor() - start: bytesLength={}", bytes == null ? null : bytes.length);
        try {
            this.document = new HWPFDocument(new ByteArrayInputStream(bytes));
            log.info("DocDocumentProcessor() - end: paragraphs={}", this.document.getRange().numParagraphs());
        } catch (Exception e) {
            log.error("DocDocumentProcessor() - end with error: bytesLength={}", bytes == null ? null : bytes.length, e);
            throw new TemplateReadWriteException("Failed to read DOC template", e);
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
    public void applyTemplateTokens(Map<String, Object> templateToken, GenerateOptions options, WarningCollector warningCollector) {
        log.info("applyTemplateTokens() - start: tokenCount={}, missingValuePolicy={}",
            templateToken == null ? null : templateToken.size(),
            options == null ? null : options.missingValuePolicy());
        Range range = document.getRange();
        int tableInsertions = 0;
        int scalarReplacements = 0;

        // Replace exact paragraph placeholders with table blocks.
        for (int i = 0; i < range.numParagraphs(); i++) {
            Paragraph paragraph = range.getParagraph(i);
            if (paragraph == null) {
                continue;
            }
            String paragraphText = normalizeParagraphText(paragraph.text());
            if (paragraphText.isEmpty()) {
                continue;
            }
            String exactToken = TokenResolver.getExactToken(paragraphText);
            if (exactToken == null) {
                continue;
            }

            Object resolved = TokenResolver.resolvePath(templateToken, exactToken);
            if (!TokenResolver.isTableValue(resolved)) {
                continue;
            }

            List<Map<String, Object>> rows = TokenResolver.toTableRows(resolved);
            if (rows == null) {
                warningCollector.add("TABLE_TOKEN_INVALID", "Table token has invalid structure: " + exactToken, "doc:paragraph#" + i);
                continue;
            }
            if (rows.isEmpty()) {
                warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + exactToken, "doc:paragraph#" + i);
                paragraph.replaceText(paragraphText, "");
                continue;
            }

            paragraph.replaceText(paragraphText, renderTableAsDocText(rows));
            tableInsertions++;
        }

        // token replacements for plain tokens.
        for (Map.Entry<String, Object> entry : templateToken.entrySet()) {
            String token = "{{" + entry.getKey() + "}}";
            Object value = entry.getValue();
            if (TokenResolver.isTableValue(value)) {
                continue;
            }
            range.replaceText(token, value == null ? "" : String.valueOf(value));
            scalarReplacements++;
        }
        log.info("applyTemplateTokens() - end: tableInsertions={}, scalarReplacements={}", tableInsertions, scalarReplacements);
    }

    @Override
    public byte[] serialize() {
        log.info("serialize() - start");
        try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.write(output);
            byte[] bytes = output.toByteArray();
            log.info("serialize() - end: bytesLength={}", bytes.length);
            return bytes;
        } catch (Exception e) {
            log.error("serialize() - end with error", e);
            throw new TemplateReadWriteException("Failed to serialize DOC document", e);
        }
    }

    @Override
    public void close() {
        log.info("close() - start");
        try {
            document.close();
            log.info("close() - end: closed=true");
        } catch (Exception e) {
            log.warn("close() - end with warning: failedToClose=true", e);
            // no-op
        }
    }

    private String normalizeParagraphText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace("\u0007", "").replace("\r", "").trim();
    }

    private String renderTableAsDocText(List<Map<String, Object>> rows) {
        List<String> columns = buildColumnOrder(rows);
        StringBuilder sb = new StringBuilder();
        sb.append(String.join("\t", columns)).append("\r");
        for (Map<String, Object> row : rows) {
            for (int c = 0; c < columns.size(); c++) {
                if (c > 0) {
                    sb.append('\t');
                }
                Object value = row.get(columns.get(c));
                sb.append(value == null ? "" : value);
            }
            sb.append("\r");
        }
        return sb.toString();
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
}
