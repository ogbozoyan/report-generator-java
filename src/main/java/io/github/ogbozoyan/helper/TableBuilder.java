package io.github.ogbozoyan.helper;

import io.github.ogbozoyan.data.GenerateOptions;
import io.github.ogbozoyan.data.MissingValuePolicy;
import io.github.ogbozoyan.data.ParagraphTarget;
import io.github.ogbozoyan.data.ResolvedText;
import io.github.ogbozoyan.util.TokenResolver;
import io.github.ogbozoyan.util.WarningCollector;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static io.github.ogbozoyan.helper.DocxHelper.collectParagraphTargets;
import static io.github.ogbozoyan.helper.DocxHelper.replaceParagraphText;
import static io.github.ogbozoyan.helper.DocxHelper.setCellText;

/**
 * Builds and fills DOCX tables based on token placeholders ({@code {{token}}}).
 *
 * <p>Supports two stages:
 * <ul>
 *     <li>row-template expansion in existing tables for {@code List<?>} payloads,</li>
 *     <li>scalar replacement in all paragraphs for remaining {@code {{token}}} placeholders.</li>
 * </ul>
 */
public class TableBuilder extends CommonHelper {

    private static final Pattern DOCX_TOKEN_PATTERN = Pattern.compile("\\{\\{\\s*([a-zA-Z0-9_.-]+)\\s*}}");

    /**
     * Expands row templates in tables and resolves scalar tokens in document paragraphs.
     *
     * @param document               destination DOCX document
     * @param templateTokensMappings runtime token map
     * @param options                generation options
     * @param warningCollector       collector for non-fatal warnings
     */
    public void apply(
        XWPFDocument document,
        Map<String, Object> templateTokensMappings,
        GenerateOptions options,
        WarningCollector warningCollector
    ) {
        if (document == null) {
            return;
        }

        Map<String, Object> context = templateTokensMappings == null ? Map.of() : templateTokensMappings;
        MissingValuePolicy policy = options == null ? MissingValuePolicy.EMPTY_AND_LOG : options.missingValuePolicy();

        List<XWPFTable> tables = new ArrayList<>();
        collectTables(document, tables);
        for (int tableIndex = 0; tableIndex < tables.size(); tableIndex++) {
            processTable(tables.get(tableIndex), tableIndex, context, policy, warningCollector);
        }

        List<ParagraphTarget> paragraphTargets = collectParagraphTargets(document);
        for (ParagraphTarget paragraphTarget : paragraphTargets) {
            XWPFParagraph paragraph = paragraphTarget.paragraph();
            String text = paragraphTarget.text();
            if (text == null || text.isEmpty()) {
                continue;
            }
            ResolvedText resolved = resolveDocxTokens(
                text,
                context,
                policy,
                warningCollector,
                paragraphTarget.location()
            );
            if (resolved.changed()) {
                replaceParagraphText(paragraph, resolved.value());
            }
        }
    }

    private void processTable(
        XWPFTable table,
        int tableIndex,
        Map<String, Object> context,
        MissingValuePolicy policy,
        WarningCollector warningCollector
    ) {
        for (int rowIndex = 0; rowIndex < table.getRows().size(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            LinkedHashSet<String> rowTokens = extractDocxTokens(row);
            if (rowTokens.isEmpty()) {
                continue;
            }

            MatchedRows matchedRows = findRowsForTemplate(rowTokens, context);
            if (matchedRows.ambiguous()) {
                warningCollector.add(
                    "TABLE_TOKEN_INVALID",
                    "Ambiguous table rows payload for template row tokens: " + rowTokens,
                    "docx:table#" + tableIndex + "/row#" + rowIndex
                );
                continue;
            }
            if (matchedRows.rows() == null) {
                continue;
            }

            int templateRowIndex = rowIndex;
            removePlaceholderTailRows(table, templateRowIndex + 1);

            List<Map<String, Object>> rows = matchedRows.rows();
            if (rows.isEmpty()) {
                table.removeRow(templateRowIndex);
                warningCollector.add(
                    "TABLE_TOKEN_EMPTY",
                    "Table row template has empty rows payload: " + matchedRows.key(),
                    "docx:table#" + tableIndex + "/row#" + templateRowIndex
                );
                rowIndex = Math.max(-1, templateRowIndex - 1);
                continue;
            }

            XWPFTableRow templateRow = table.getRow(templateRowIndex);
            List<String> templateCellTexts = snapshotTemplateCellTexts(templateRow);
            fillRowWithContext(
                templateRow,
                context,
                rows.get(0),
                policy,
                warningCollector,
                "docx:table#" + tableIndex + "/row#" + templateRowIndex
            );

            for (int r = 1; r < rows.size(); r++) {
                int insertPos = templateRowIndex + r;
                XWPFTableRow newRow = table.insertNewTableRow(insertPos);
                fillInsertedRowFromTemplate(
                    templateRow,
                    newRow,
                    templateCellTexts,
                    context,
                    rows.get(r),
                    policy,
                    warningCollector,
                    "docx:table#" + tableIndex + "/row#" + insertPos
                );
            }

            rowIndex = templateRowIndex + rows.size() - 1;
        }
    }

    private List<String> snapshotTemplateCellTexts(XWPFTableRow templateRow) {
        List<String> cellTexts = new ArrayList<>();
        for (XWPFTableCell cell : templateRow.getTableCells()) {
            cellTexts.add(cell.getText());
        }
        return cellTexts;
    }

    private void fillInsertedRowFromTemplate(
        XWPFTableRow templateRow,
        XWPFTableRow newRow,
        List<String> templateCellTexts,
        Map<String, Object> globalContext,
        Map<String, Object> rowContext,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String locationPrefix
    ) {
        if (templateRow == null || newRow == null) {
            return;
        }

        if (templateRow.getCtRow().getTrPr() != null) {
            CTTrPr trPrCopy = (CTTrPr) templateRow.getCtRow().getTrPr().copy();
            newRow.getCtRow().setTrPr(trPrCopy);
        }

        LinkedHashMap<String, Object> context = new LinkedHashMap<>(globalContext);
        context.putAll(rowContext);

        for (int c = 0; c < templateCellTexts.size(); c++) {
            XWPFTableCell newCell = newRow.addNewTableCell();
            XWPFTableCell templateCell = templateRow.getCell(c);
            if (templateCell != null && templateCell.getCTTc().getTcPr() != null) {
                CTTcPr tcPrCopy = (CTTcPr) templateCell.getCTTc().getTcPr().copy();
                newCell.getCTTc().setTcPr(tcPrCopy);
            }

            String sourceText = templateCellTexts.get(c);
            if (sourceText == null || sourceText.isEmpty()) {
                setCellText(newCell, "");
                continue;
            }
            ResolvedText resolved = resolveDocxTokens(
                sourceText,
                context,
                policy,
                warningCollector,
                locationPrefix + "/cell#" + c
            );
            setCellText(newCell, resolved.value());
        }
    }

    private void fillRowWithContext(
        XWPFTableRow row,
        Map<String, Object> globalContext,
        Map<String, Object> rowContext,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String locationPrefix
    ) {
        if (row == null) {
            return;
        }
        LinkedHashMap<String, Object> context = new LinkedHashMap<>(globalContext);
        context.putAll(rowContext);

        List<XWPFTableCell> cells = row.getTableCells();
        for (int c = 0; c < cells.size(); c++) {
            XWPFTableCell cell = cells.get(c);
            String text = cell.getText();
            if (text == null || text.isEmpty()) {
                continue;
            }
            ResolvedText resolved = resolveDocxTokens(
                text,
                context,
                policy,
                warningCollector,
                locationPrefix + "/cell#" + c
            );
            if (resolved.changed()) {
                setCellText(cell, resolved.value());
            }
        }
    }

    private MatchedRows findRowsForTemplate(LinkedHashSet<String> rowTokens, Map<String, Object> context) {
        String bestKey = null;
        List<Map<String, Object>> bestRows = null;
        int bestScore = 0;
        boolean ambiguous = false;

        for (Map.Entry<String, Object> entry : context.entrySet()) {
            List<Map<String, Object>> candidateRows = toObjectRows(entry.getValue());
            if (candidateRows == null) {
                continue;
            }

            int score = scoreRowTemplate(rowTokens, candidateRows);
            if (score == 0) {
                continue;
            }
            if (score > bestScore) {
                bestScore = score;
                bestKey = entry.getKey();
                bestRows = candidateRows;
                ambiguous = false;
                continue;
            }
            if (score == bestScore && !Objects.equals(bestKey, entry.getKey())) {
                ambiguous = true;
            }
        }

        if (bestRows != null) {
            return new MatchedRows(bestKey, bestRows, ambiguous);
        }
        return new MatchedRows(null, null, false);
    }

    private int scoreRowTemplate(LinkedHashSet<String> rowTokens, List<Map<String, Object>> rows) {
        if (rows.isEmpty()) {
            return 0;
        }
        Map<String, Object> firstRow = rows.get(0);
        int score = 0;
        for (String token : rowTokens) {
            if (firstRow.containsKey(token)) {
                score++;
            }
        }
        return score;
    }

    private LinkedHashSet<String> extractDocxTokens(XWPFTableRow row) {
        LinkedHashSet<String> tokens = new LinkedHashSet<>();
        List<XWPFTableCell> cells = row.getTableCells();
        for (XWPFTableCell cell : cells) {
            String text = cell.getText();
            if (text == null || text.isBlank()) {
                continue;
            }
            Matcher matcher = DOCX_TOKEN_PATTERN.matcher(text);
            while (matcher.find()) {
                tokens.add(matcher.group(1));
            }
        }
        return tokens;
    }

    private int removePlaceholderTailRows(XWPFTable table, int startRow) {
        int removed = 0;
        while (startRow < table.getRows().size()) {
            XWPFTableRow row = table.getRow(startRow);
            if (!isPlaceholderRow(row)) {
                break;
            }
            table.removeRow(startRow);
            removed++;
        }
        return removed;
    }

    private boolean isPlaceholderRow(XWPFTableRow row) {
        if (row == null) {
            return false;
        }

        List<String> nonBlankCellTexts = new ArrayList<>();
        for (XWPFTableCell cell : row.getTableCells()) {
            String normalized = normalizeCellText(cell.getText());
            if (!normalized.isBlank()) {
                nonBlankCellTexts.add(normalized);
            }
        }

        if (nonBlankCellTexts.isEmpty()) {
            return true;
        }
        if (nonBlankCellTexts.size() == 1) {
            String onlyValue = nonBlankCellTexts.get(0);
            return "…".equals(onlyValue)
                || "...".equals(onlyValue)
                || "n.".equalsIgnoreCase(onlyValue)
                || "n".equalsIgnoreCase(onlyValue);
        }
        return false;
    }

    private String normalizeCellText(String value) {
        if (value == null) {
            return "";
        }
        return value.replace('\u00A0', ' ').trim();
    }

    private ResolvedText resolveDocxTokens(
        String text,
        Map<String, Object> context,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location
    ) {
        Matcher matcher = DOCX_TOKEN_PATTERN.matcher(text);
        StringBuilder sb = new StringBuilder();
        boolean changed = false;

        while (matcher.find()) {
            String token = matcher.group(1);
            Object resolved = TokenResolver.resolvePath(context, token);

            if (resolved == null) {
                String replacement = switch (policy) {
                    case EMPTY_AND_LOG -> {
                        warningCollector.add(
                            "MISSING_TOKEN",
                            "Token not found in file but was present in template: " + token,
                            location
                        );
                        yield "";
                    }
                    case LEAVE_TOKEN -> matcher.group(0);
                    case FAIL_FAST -> throw new io.github.ogbozoyan.exception.TemplateDataBindingException(
                        "Token not found: " + token + " at " + location
                    );
                };
                matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
                changed = changed || !Objects.equals(matcher.group(0), replacement);
                continue;
            }

            if (toObjectRows(resolved) != null) {
                String exactToken = TokenResolver.getExactToken(text);
                if (exactToken == null || !Objects.equals(exactToken, token)) {
                    warningCollector.add(
                        "TABLE_TOKEN_INLINE_IGNORED",
                        "Table token can be inserted only as exact placeholder",
                        location
                    );
                }
                matcher.appendReplacement(sb, Matcher.quoteReplacement(matcher.group(0)));
                continue;
            }

            String replacement = String.valueOf(resolved);
            matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
            changed = true;
        }

        matcher.appendTail(sb);
        return new ResolvedText(sb.toString(), changed);
    }

    private void collectTables(IBody body, List<XWPFTable> destination) {
        List<IBodyElement> elements = body.getBodyElements();
        for (IBodyElement element : elements) {
            if (!(element instanceof XWPFTable table)) {
                continue;
            }
            destination.add(table);
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    collectTables(cell, destination);
                }
            }
        }
    }

    private List<Map<String, Object>> toObjectRows(Object value) {
        if (!(value instanceof List<?> list)) {
            return null;
        }
        List<Map<String, Object>> rows = new ArrayList<>(list.size());
        for (Object item : list) {
            if (item == null) {
                rows.add(Map.of());
                continue;
            }
            if (item instanceof Map<?, ?> map) {
                LinkedHashMap<String, Object> row = new LinkedHashMap<>();
                for (Map.Entry<?, ?> entry : map.entrySet()) {
                    String key = entry.getKey() == null ? "" : String.valueOf(entry.getKey());
                    row.put(key, entry.getValue());
                }
                rows.add(row);
                continue;
            }

            LinkedHashMap<String, Object> beanMap = beanToMap(item);
            if (beanMap.isEmpty()) {
                return null;
            }
            rows.add(beanMap);
        }
        return rows;
    }

    private LinkedHashMap<String, Object> beanToMap(Object bean) {
        LinkedHashMap<String, Object> row = new LinkedHashMap<>();
        Method[] methods = bean.getClass().getMethods();
        for (Method method : methods) {
            if (method.getParameterCount() != 0) {
                continue;
            }
            String name = method.getName();
            if ("getClass".equals(name)) {
                continue;
            }

            String property = null;
            if (name.startsWith("get") && name.length() > 3) {
                property = decapitalize(name.substring(3));
            } else if (name.startsWith("is") && name.length() > 2) {
                property = decapitalize(name.substring(2));
            }
            if (property == null || property.isBlank()) {
                continue;
            }

            try {
                if (!method.canAccess(bean) && !method.trySetAccessible()) {
                    continue;
                }
                Object value = method.invoke(bean);
                row.put(property, value);
                String snakeCase = toSnakeCase(property);
                if (!snakeCase.equals(property) && !row.containsKey(snakeCase)) {
                    row.put(snakeCase, value);
                }
            } catch (Exception ignored) {
                // no-op
            }
        }
        return row;
    }

    private String decapitalize(String value) {
        if (value == null || value.isEmpty()) {
            return value;
        }
        if (value.length() == 1) {
            return value.toLowerCase(Locale.ROOT);
        }
        return value.substring(0, 1).toLowerCase(Locale.ROOT) + value.substring(1);
    }

    private String toSnakeCase(String value) {
        StringBuilder sb = new StringBuilder(value.length() + 8);
        for (int i = 0; i < value.length(); i++) {
            char ch = value.charAt(i);
            if (Character.isUpperCase(ch)) {
                if (i > 0) {
                    sb.append('_');
                }
                sb.append(Character.toLowerCase(ch));
            } else {
                sb.append(ch);
            }
        }
        return sb.toString();
    }

    private record MatchedRows(String key, List<Map<String, Object>> rows, boolean ambiguous) {
    }
}
