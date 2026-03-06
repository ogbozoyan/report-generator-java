package io.github.ogbozoyan.helper;

import io.github.ogbozoyan.contract.TableBuilder;
import io.github.ogbozoyan.data.MissingValuePolicy;
import io.github.ogbozoyan.data.ResolvedText;
import io.github.ogbozoyan.util.TokenResolver;
import io.github.ogbozoyan.util.WarningCollector;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

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

    /**
     * Renders declarative table payload as DOC-compatible text grid.
     *
     * <p>Because HWPF does not expose robust table construction APIs for this flow,
     * the table is represented as text rows separated by {@code \\r} and columns by {@code \\t}.
     * Colspan is flattened into additional empty columns.
     *
     * @param table            declarative table payload
     * @param tokenContext     token context for scalar placeholders inside table cells
     * @param policy           missing token policy
     * @param warningCollector collector for non-fatal warnings
     * @param location         base diagnostic location
     * @return DOC-compatible text representation
     */
    public static String renderDeclarativeTableAsDocText(
        TableBuilder table,
        Map<String, Object> tokenContext,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location
    ) {
        Map<String, Object> context = tokenContext == null ? Map.of() : tokenContext;
        int columnCount = table.columnCount();
        StringBuilder sb = new StringBuilder();
        List<TableBuilder.Row> rows = table.rows();
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            TableBuilder.Row row = rows.get(rowIndex);
            List<String> flattenedCells = new ArrayList<>();
            for (int colIndex = 0; colIndex < row.cells().size(); colIndex++) {
                TableBuilder.Cell cell = row.cells().get(colIndex);
                String cellLocation = location + "/row#" + rowIndex + "/cell#" + colIndex;
                ResolvedText resolvedText = TokenResolver.resolve(
                    cell.text(),
                    context,
                    policy,
                    warningCollector,
                    cellLocation,
                    false
                );
                flattenedCells.add(resolvedText.value());
                for (int span = 1; span < cell.colSpan(); span++) {
                    flattenedCells.add("");
                }
            }

            while (flattenedCells.size() < columnCount) {
                flattenedCells.add("");
            }
            sb.append(String.join("\t", flattenedCells)).append("\r");
        }
        return sb.toString();
    }
}
