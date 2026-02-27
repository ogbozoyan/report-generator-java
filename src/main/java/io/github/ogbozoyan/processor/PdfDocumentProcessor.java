package io.github.ogbozoyan.processor;

import io.github.ogbozoyan.contract.GenerateOptions;
import io.github.ogbozoyan.contract.ResolvedText;
import io.github.ogbozoyan.contract.TemplateScanResult;
import io.github.ogbozoyan.exception.TemplateReadWriteException;
import io.github.ogbozoyan.util.TokenResolver;
import io.github.ogbozoyan.util.WarningCollector;
import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

/**
 * Processor for PDF templates.
 * <p>
 * Implementation note: PDFs are immutable for in-place text editing in this io.github.ogbozoyan.service.
 * The io.github.ogbozoyan.processor extracts text, replaces scalar tokens, and writes a new text PDF.
 *
 * <p>Table tokens are rendered as ASCII grid text blocks.
 */
@Slf4j
public class PdfDocumentProcessor implements WorkbookProcessor {

    private String extractedText;

    /**
     * Creates io.github.ogbozoyan.processor and extracts text from source PDF.
     *
     * @param bytes source PDF bytes
     * @throws TemplateReadWriteException when source PDF cannot be parsed
     */
    public PdfDocumentProcessor(byte[] bytes) {
        log.info("PdfDocumentProcessor() - start: bytesLength={}", bytes == null ? null : bytes.length);
        try (PDDocument document = Loader.loadPDF(bytes)) {
            PDFTextStripper textStripper = new PDFTextStripper();
            extractedText = textStripper.getText(document);
            log.info("PdfDocumentProcessor() - end: extractedTextLength={}, pageCount={}",
                extractedText == null ? 0 : extractedText.length(),
                document.getNumberOfPages());
        } catch (Exception e) {
            log.error("PdfDocumentProcessor() - end with error: bytesLength={}", bytes == null ? null : bytes.length, e);
            throw new TemplateReadWriteException("Failed to read PDF template", e);
        }
    }

    /**
     * Splits text into logical paragraphs by newline boundaries.
     *
     * @param text source text
     * @return paragraph lines preserving empty lines
     */
    private static List<String> splitParagraphs(String text) {
        if (text == null || text.isEmpty()) {
            return List.of("");
        }
        return List.of(text.replace("\r\n", "\n").replace('\r', '\n').split("\n", -1));
    }

    /**
     * Wraps single line to fit page width in points.
     *
     * @param text     source line
     * @param font     target font
     * @param fontSize font size in points
     * @param maxWidth max line width in points
     * @return wrapped lines
     * @throws Exception when font metrics cannot be evaluated
     */
    private static List<String> wrapLine(String text, PDType1Font font, float fontSize, float maxWidth) throws Exception {
        if (text == null || text.isEmpty()) {
            return List.of("");
        }

        List<String> lines = new ArrayList<>();
        String[] words = text.split("\\s+");
        StringBuilder current = new StringBuilder();

        for (String word : words) {
            String candidate = current.isEmpty() ? word : current + " " + word;
            float width = font.getStringWidth(sanitizeForPdf(candidate)) / 1000f * fontSize;

            if (width <= maxWidth) {
                current.setLength(0);
                current.append(candidate);
                continue;
            }

            if (!current.isEmpty()) {
                lines.add(current.toString());
                current.setLength(0);
            }

            // hard-wrap very long words
            String remainder = word;
            while (!remainder.isEmpty()) {
                int splitAt = findSplitIndex(remainder, font, fontSize, maxWidth);
                lines.add(remainder.substring(0, splitAt));
                remainder = remainder.substring(splitAt);
            }
        }

        if (!current.isEmpty()) {
            lines.add(current.toString());
        }
        return lines;
    }

    /**
     * Finds longest prefix that fits max width for hard-wrap scenario.
     *
     * @param text     source text
     * @param font     target font
     * @param fontSize font size
     * @param maxWidth max width in points
     * @return split index, always at least {@code 1}
     * @throws Exception when font metrics cannot be evaluated
     */
    private static int findSplitIndex(String text, PDType1Font font, float fontSize, float maxWidth) throws Exception {
        int split = Math.max(1, text.length());
        while (split > 1) {
            String candidate = sanitizeForPdf(text.substring(0, split));
            float width = font.getStringWidth(candidate) / 1000f * fontSize;
            if (width <= maxWidth) {
                return split;
            }
            split--;
        }
        return 1;
    }

    /**
     * Converts unsupported glyphs to printable ASCII subset.
     *
     * @param text source text
     * @return sanitized text
     */
    private static String sanitizeForPdf(String text) {
        StringBuilder sanitized = new StringBuilder(text.length());
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch >= 32 && ch <= 126) {
                sanitized.append(ch);
            } else if (Character.isWhitespace(ch)) {
                sanitized.append(' ');
            } else {
                sanitized.append('?');
            }
        }
        return sanitized.toString();
    }

    /**
     * Returns empty scan result because PDF path currently performs direct apply phase.
     *
     * @return empty scan result
     */
    @Override
    public TemplateScanResult scan() {
        log.info("scan() - start");
        TemplateScanResult result = new TemplateScanResult(List.of(), List.of());
        log.info("scan() - end: markers={}, tokens={}", result.markers().size(), result.scalarTokens().size());
        return result;
    }

    /**
     * Applies token replacement to extracted text representation.
     *
     * @param templateToken    token map
     * @param options          generation options
     * @param warningCollector collector for non-fatal warnings
     */
    @Override
    public void applyTemplateTokens(Map<String, Object> templateToken, GenerateOptions options, WarningCollector warningCollector) {
        log.info("applyTemplateTokens() - start: tokenCount={}, extractedTextLength={}",
            templateToken == null ? null : templateToken.size(),
            extractedText == null ? 0 : extractedText.length());
        extractedText = replaceTokensWithTables(extractedText == null ? "" : extractedText, templateToken, options, warningCollector);
        log.info("applyTemplateTokens() - end: extractedTextLength={}", extractedText == null ? 0 : extractedText.length());
    }

    /**
     * Serializes processed text into a newly constructed PDF.
     *
     * @return generated PDF bytes
     */
    @Override
    public byte[] serialize() {
        log.info("serialize() - start: extractedTextLength={}", extractedText == null ? 0 : extractedText.length());
        try (PDDocument document = new PDDocument();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            PDType1Font font = new PDType1Font(Standard14Fonts.FontName.HELVETICA);
            float fontSize = 11f;
            float leading = 1.35f * fontSize;
            float margin = 48f;

            PDRectangle pageSize = PDRectangle.A4;
            PDPage page = new PDPage(pageSize);
            document.addPage(page);

            PDPageContentStream contentStream = new PDPageContentStream(document, page);
            contentStream.setFont(font, fontSize);
            contentStream.beginText();

            float y = pageSize.getHeight() - margin;
            contentStream.newLineAtOffset(margin, y);
            float availableWidth = pageSize.getWidth() - 2 * margin;

            for (String paragraph : splitParagraphs(extractedText)) {
                List<String> wrappedLines = wrapLine(paragraph, font, fontSize, availableWidth);
                if (wrappedLines.isEmpty()) {
                    wrappedLines = List.of("");
                }

                for (String line : wrappedLines) {
                    if (y - leading < margin) {
                        contentStream.endText();
                        contentStream.close();

                        page = new PDPage(pageSize);
                        document.addPage(page);
                        contentStream = new PDPageContentStream(document, page);
                        contentStream.setFont(font, fontSize);
                        contentStream.beginText();
                        y = pageSize.getHeight() - margin;
                        contentStream.newLineAtOffset(margin, y);
                    }

                    contentStream.showText(sanitizeForPdf(line));
                    contentStream.newLineAtOffset(0, -leading);
                    y -= leading;
                }
            }

            contentStream.endText();
            contentStream.close();

            document.save(outputStream);
            byte[] bytes = outputStream.toByteArray();
            log.info("serialize() - end: bytesLength={}", bytes.length);
            return bytes;
        } catch (Exception e) {
            log.error("serialize() - end with error", e);
            throw new TemplateReadWriteException("Failed to serialize PDF document", e);
        }
    }

    /**
     * Replaces scalar/table tokens in extracted text representation.
     *
     * <p>For exact-placeholder table tokens this method renders ASCII grid table and replaces
     * whole line with rendered block.
     *
     * @param source           extracted source text
     * @param templateTokens   token map
     * @param options          generation options
     * @param warningCollector warning collector
     * @return transformed text
     */
    private String replaceTokensWithTables(
        String source,
        Map<String, Object> templateTokens,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        log.info("replaceTokensWithTables() - start: sourceLength={}, tokenCount={}",
            source == null ? 0 : source.length(),
            templateTokens == null ? null : templateTokens.size());
        String normalized = source.replace("\r\n", "\n").replace('\r', '\n');
        String[] lines = normalized.split("\n", -1);
        StringBuilder result = new StringBuilder();
        int tableInsertions = 0;
        int scalarReplacements = 0;

        for (int i = 0; i < lines.length; i++) {
            String line = lines[i];
            String exactToken = TokenResolver.getExactToken(line.trim());
            if (exactToken != null) {
                Object resolved = TokenResolver.resolvePath(templateTokens, exactToken);
                if (TokenResolver.isTableValue(resolved)) {
                    List<Map<String, Object>> rows = TokenResolver.toTableRows(resolved);
                    if (rows == null) {
                        warningCollector.add("TABLE_TOKEN_INVALID", "Table token has invalid structure: " + exactToken, "pdf:line#" + (i + 1));
                    } else if (rows.isEmpty()) {
                        warningCollector.add("TABLE_TOKEN_EMPTY", "Table token has no rows: " + exactToken, "pdf:line#" + (i + 1));
                    } else {
                        result.append(renderAsciiTable(rows));
                        tableInsertions++;
                    }
                    if (i < lines.length - 1) {
                        result.append('\n');
                    }
                    continue;
                }
            }

            ResolvedText resolvedText = TokenResolver.resolve(
                line,
                templateTokens,
                options.missingValuePolicy(),
                warningCollector,
                "pdf:line#" + (i + 1),
                false
            );
            result.append(resolvedText.value());
            if (resolvedText.changed()) {
                scalarReplacements++;
            }
            if (i < lines.length - 1) {
                result.append('\n');
            }
        }
        String output = result.toString();
        log.info("replaceTokensWithTables() - end: outputLength={}, tableInsertions={}, scalarReplacements={}",
            output.length(), tableInsertions, scalarReplacements);
        return output;
    }

    /**
     * Renders rows as ASCII grid table.
     *
     * @param rows normalized table rows
     * @return rendered table text
     */
    private String renderAsciiTable(List<Map<String, Object>> rows) {
        List<String> columns = buildColumnOrder(rows);
        int[] widths = new int[columns.size()];
        for (int c = 0; c < columns.size(); c++) {
            widths[c] = columns.get(c).length();
        }
        for (Map<String, Object> row : rows) {
            for (int c = 0; c < columns.size(); c++) {
                Object value = row.get(columns.get(c));
                widths[c] = Math.max(widths[c], value == null ? 0 : String.valueOf(value).length());
            }
        }

        StringBuilder sb = new StringBuilder();
        sb.append(buildRow(columns, widths)).append('\n');
        sb.append(buildSeparator(widths)).append('\n');

        for (int r = 0; r < rows.size(); r++) {
            Map<String, Object> row = rows.get(r);
            List<String> values = new ArrayList<>(columns.size());
            for (String column : columns) {
                Object value = row.get(column);
                values.add(value == null ? "" : String.valueOf(value));
            }
            sb.append(buildRow(values, widths));
            if (r < rows.size() - 1) {
                sb.append('\n');
            }
        }
        return sb.toString();
    }

    /**
     * Builds separator line for ASCII table.
     *
     * @param widths per-column widths
     * @return separator line
     */
    private String buildSeparator(int[] widths) {
        StringBuilder sb = new StringBuilder();
        for (int width : widths) {
            if (sb.length() > 0) {
                sb.append("-+-");
            }
            sb.append("-".repeat(Math.max(1, width)));
        }
        return sb.toString();
    }

    /**
     * Builds padded row line for ASCII table.
     *
     * @param cells  row cell values
     * @param widths per-column widths
     * @return row line
     */
    private String buildRow(List<String> cells, int[] widths) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < cells.size(); i++) {
            if (i > 0) {
                sb.append(" | ");
            }
            sb.append(padRight(cells.get(i), widths[i]));
        }
        return sb.toString();
    }

    /**
     * Right-pads value to fixed length.
     *
     * @param value  source value
     * @param length target length
     * @return padded string
     */
    private String padRight(String value, int length) {
        String source = value == null ? "" : value;
        if (source.length() >= length) {
            return source;
        }
        return source + " ".repeat(length - source.length());
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
}
