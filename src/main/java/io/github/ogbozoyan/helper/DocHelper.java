package io.github.ogbozoyan.helper;

import lombok.experimental.Helper;
import lombok.extern.slf4j.Slf4j;

import java.util.List;
import java.util.Map;

@Slf4j
@Helper
public class DocHelper extends CommonHelper {

    /**
     * Normalizes HWPF paragraph text by removing control markers and trimming spaces.
     *
     * @param text source paragraph text
     * @return normalized text
     */
    public static String normalizeParagraphText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace("\u0007", "").replace("\r", "").trim();
    }

    /**
     * Renders table payload as DOC-compatible text grid.
     *
     * <p>First row is header, followed by data rows.
     *
     * @param rows normalized table rows
     * @return grid text with {@code \\t} and {@code \\r} separators
     */
    public static String renderTableAsDocText(List<Map<String, Object>> rows) {
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
}
