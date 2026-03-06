package io.github.ogbozoyan.contract;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * Declarative table builder for XLS/XLSX insertion in {@code PoiWorkbookProcessor}.
 *
 * <p>Use this type as table token payload in {@code ReportData.templateTokens()} for an exact
 * placeholder like {@code {{TABLE_HERE}}}.
 *
 * <p>Example:
 * <pre>{@code
 * TableXlsxBuilder schedule = TableXlsxBuilder.create()
 *     .row(TableXlsxBuilder.boldCell("Payment schedule", 4))
 *     .row(
 *         TableXlsxBuilder.boldCell("No"),
 *         TableXlsxBuilder.boldCell("Month"),
 *         TableXlsxBuilder.boldCell("Amount"),
 *         TableXlsxBuilder.boldCell("Balance")
 *     )
 *     .row(
 *         TableXlsxBuilder.cell("1."),
 *         TableXlsxBuilder.cell("{{payment_date}}"),
 *         TableXlsxBuilder.cell("{{amount}}"),
 *         TableXlsxBuilder.cell("{{balance}}")
 *     );
 * }</pre>
 */
public final class TableXlsxBuilder {

    private final List<Row> rows = new ArrayList<>();

    private TableXlsxBuilder() {
    }

    /**
     * Creates an empty builder.
     *
     * @return new builder instance
     */
    public static TableXlsxBuilder create() {
        return new TableXlsxBuilder();
    }

    /**
     * Creates plain cell with span 1.
     *
     * @param value cell value
     * @return cell spec
     */
    public static Cell cell(Object value) {
        return new Cell(value, false, 1);
    }

    /**
     * Creates plain cell with explicit colspan.
     *
     * @param value   cell value
     * @param colSpan horizontal span
     * @return cell spec
     */
    public static Cell cell(Object value, int colSpan) {
        return new Cell(value, false, colSpan);
    }

    /**
     * Creates bold cell with span 1.
     *
     * @param value cell value
     * @return cell spec
     */
    public static Cell boldCell(Object value) {
        return new Cell(value, true, 1);
    }

    /**
     * Creates bold cell with explicit colspan.
     *
     * @param value   cell value
     * @param colSpan horizontal span
     * @return cell spec
     */
    public static Cell boldCell(Object value, int colSpan) {
        return new Cell(value, true, colSpan);
    }

    /**
     * Appends a row.
     *
     * @param cells row cells
     * @return this builder
     */
    public TableXlsxBuilder row(Cell... cells) {
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
    public TableXlsxBuilder insertRow(int index, Cell... cells) {
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
    public TableXlsxBuilder replaceRow(int index, Cell... cells) {
        rows.set(index, Row.of(cells));
        return this;
    }

    /**
     * Removes row at index.
     *
     * @param index row index
     * @return this builder
     */
    public TableXlsxBuilder removeRow(int index) {
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
     * @param value   cell value
     * @param bold    bold flag
     * @param colSpan horizontal span, {@code >= 1}
     */
    public record Cell(Object value, boolean bold, int colSpan) {
        /**
         * Validates cell specification.
         */
        public Cell {
            if (colSpan < 1) {
                throw new IllegalArgumentException("colSpan must be >= 1");
            }
        }

        /**
         * Returns copy with bold enabled.
         *
         * @return bold cell
         */
        public Cell withBold() {
            return new Cell(value, true, colSpan);
        }

        /**
         * Returns copy with changed colspan.
         *
         * @param value new colspan
         * @return resized cell
         */
        public Cell withColSpan(int value) {
            return new Cell(this.value, bold, value);
        }
    }
}
