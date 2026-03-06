package io.github.ogbozoyan.contract;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * Declarative string constructor for extending DOCX template strings inside existing tables.
 *
 * <p>Use this type as the value in {@code ReportData.templateTokens()} for accurate
 * placeholder type {@code {{PAYMENT_ROWS}}}, located inside and at the beginning of the table row.
 * The processor clones the entire row of the template and writes the specified cell values,
 * While maintaining the original formatting of rows/cells.
 *
 * <p>Example:
 * <pre>{@code
 * RowBuilder rows = RowBuilder.create()
 *     .row(
 *         RowBuilder.cell("1"),
 *         RowBuilder.cell("2026-03"),
 *         RowBuilder.cell("250000"),
 *         RowBuilder.cell("750000")
 *     )
 *     .row(
 *         RowBuilder.cell("2"),
 *         RowBuilder.cell("2026-04"),
 *         RowBuilder.cell("250000"),
 *         RowBuilder.cell("500000")
 *     );
 * }</pre>
 */
public final class RowBuilder {

    private final List<Row> rows = new ArrayList<>();

    private RowBuilder() {
    }

    /**
     * Creates an empty builder.
     *
     * @return new builder instance
     */
    public static RowBuilder create() {
        return new RowBuilder();
    }

    /**
     * Creates plain cell value.
     *
     * @param text cell text
     * @return cell spec
     */
    public static Cell cell(String text) {
        return new Cell(text);
    }

    /**
     * Appends a row to the end of definition.
     *
     * @param cells row cells
     * @return this builder
     */
    public RowBuilder row(Cell... cells) {
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
    public RowBuilder insertRow(int index, Cell... cells) {
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
    public RowBuilder replaceRow(int index, Cell... cells) {
        rows.set(index, Row.of(cells));
        return this;
    }

    /**
     * Removes row at index.
     *
     * @param index row index
     * @return this builder
     */
    public RowBuilder removeRow(int index) {
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
    }

    /**
     * Declarative cell specification.
     *
     * @param text cell text (can contain scalar placeholders)
     */
    public record Cell(String text) {
        /**
         * Validates and normalizes cell specification.
         */
        public Cell {
            text = text == null ? "" : text;
        }
    }
}
