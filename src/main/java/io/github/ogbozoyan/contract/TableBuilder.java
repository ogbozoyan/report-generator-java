package io.github.ogbozoyan.contract;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * Declarative table builder for DOC/DOCX table-token insertion.
 *
 * <p>Use this type as value in {@code ReportData.templateTokens()} for an exact
 * placeholder like {@code {{TABLE_HERE}}}. The processor creates a table at token
 * position without requiring a pre-existing table in the template.
 *
 * <p>Example:
 * <pre>{@code
 * TableBuilder schedule = TableBuilder.create()
 *     .row(TableBuilder.boldCell("Payment schedule", 4))
 *     .row(
 *         TableBuilder.boldCell("No"),
 *         TableBuilder.boldCell("Month"),
 *         TableBuilder.boldCell("Amount"),
 *         TableBuilder.boldCell("Balance")
 *     )
 *     .row(
 *         TableBuilder.cell("1."),
 *         TableBuilder.cell("{{payment_date}}"),
 *         TableBuilder.cell("{{amount}}"),
 *         TableBuilder.cell("{{balance}}")
 *     );
 * }</pre>
 */
public final class TableBuilder {

    private final List<Row> rows = new ArrayList<>();

    private TableBuilder() {
    }

    /**
     * Creates an empty builder.
     *
     * @return new builder instance
     */
    public static TableBuilder create() {
        return new TableBuilder();
    }

    /**
     * Creates plain cell with span 1.
     *
     * @param text cell text
     * @return cell spec
     */
    public static Cell cell(String text) {
        return new Cell(text, false, 1);
    }

    /**
     * Creates plain cell with explicit span.
     *
     * @param text    cell text
     * @param colSpan colspan value (>=1)
     * @return cell spec
     */
    public static Cell cell(String text, int colSpan) {
        return new Cell(text, false, colSpan);
    }

    /**
     * Creates bold cell with span 1.
     *
     * @param text cell text
     * @return cell spec
     */
    public static Cell boldCell(String text) {
        return new Cell(text, true, 1);
    }

    /**
     * Creates bold cell with explicit span.
     *
     * @param text    cell text
     * @param colSpan colspan value (>=1)
     * @return cell spec
     */
    public static Cell boldCell(String text, int colSpan) {
        return new Cell(text, true, colSpan);
    }

    /**
     * Appends a row to the end of table definition.
     *
     * @param cells row cells
     * @return this builder
     */
    public TableBuilder row(Cell... cells) {
        rows.add(Row.of(cells));
        return this;
    }

    /**
     * Inserts row at index.
     *
     * @param index insertion index
     * @param cells row cells
     * @return this builder
     */
    public TableBuilder insertRow(int index, Cell... cells) {
        rows.add(index, Row.of(cells));
        return this;
    }

    /**
     * Replaces row at index.
     *
     * @param index row index
     * @param cells replacement row cells
     * @return this builder
     */
    public TableBuilder replaceRow(int index, Cell... cells) {
        rows.set(index, Row.of(cells));
        return this;
    }

    /**
     * Removes row at index.
     *
     * @param index row index
     * @return this builder
     */
    public TableBuilder removeRow(int index) {
        rows.remove(index);
        return this;
    }

    /**
     * Returns immutable row list snapshot.
     *
     * @return rows snapshot
     */
    public List<Row> rows() {
        return List.copyOf(rows);
    }

    /**
     * Resolves logical column count using max row width (sum of colspans).
     *
     * @return column count
     */
    public int columnCount() {
        int maxColumns = 0;
        for (Row row : rows) {
            maxColumns = Math.max(maxColumns, row.width());
        }
        return maxColumns;
    }

    /**
     * Declarative row specification.
     *
     * @param cells ordered row cells
     */
    public record Row(List<Cell> cells) {
        /**
         * Validates and normalizes row specification.
         */
        public Row {
            Objects.requireNonNull(cells, "cells must not be null");
            if (cells.isEmpty()) {
                throw new IllegalArgumentException("row must contain at least one cell");
            }
            cells = List.copyOf(cells);
        }

        /**
         * Creates row from varargs cells.
         *
         * @param cells row cells
         * @return row spec
         */
        public static Row of(Cell... cells) {
            Objects.requireNonNull(cells, "cells must not be null");
            return new Row(Arrays.asList(cells));
        }

        /**
         * Returns logical row width (sum of colspans).
         *
         * @return width
         */
        public int width() {
            int width = 0;
            for (Cell cell : cells) {
                width += cell.colSpan();
            }
            return width;
        }
    }

    /**
     * Declarative cell specification.
     *
     * @param text    cell text (can contain scalar placeholders)
     * @param bold    bold style flag
     * @param colSpan horizontal colspan (>=1)
     */
    public record Cell(String text, boolean bold, int colSpan) {
        /**
         * Validates and normalizes cell specification.
         */
        public Cell {
            if (colSpan < 1) {
                throw new IllegalArgumentException("colSpan must be >= 1");
            }
            text = text == null ? "" : text;
        }

        /**
         * Returns copy with bold style enabled.
         *
         * @return styled cell
         */
        public Cell withBold() {
            return new Cell(text, true, colSpan);
        }

        /**
         * Returns copy with different colspan.
         *
         * @param value new colspan (>=1)
         * @return resized cell
         */
        public Cell withColSpan(int value) {
            return new Cell(text, bold, value);
        }
    }
}
