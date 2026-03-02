package io.github.ogbozoyan.helper;

import lombok.extern.slf4j.Slf4j;

import java.util.List;

@Slf4j
public class PDFHelper {
    /**
     * Builds separator line for ASCII table.
     *
     * @param widths per-column widths
     * @return separator line
     */
    public static String buildSeparator(int[] widths) {
        StringBuilder sb = new StringBuilder();
        for (int width : widths) {
            if (!sb.isEmpty()) {
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
    public static String buildRow(List<String> cells, int[] widths) {
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
    public static String padRight(String value, int length) {
        String source = value == null ? "" : value;
        if (source.length() >= length) {
            return source;
        }
        return source + " ".repeat(length - source.length());
    }
}
