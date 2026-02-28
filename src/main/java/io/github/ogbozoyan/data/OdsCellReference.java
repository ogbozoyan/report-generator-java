package io.github.ogbozoyan.data;

/**
 * Lightweight ODS cell reference used by ODS-specific scanner code.
 *
 * @param rowIndex   zero-based row index
 * @param colIndex   zero-based column index
 * @param sourceText original cell text
 */
public record OdsCellReference(int rowIndex, int colIndex, String sourceText) {
}
