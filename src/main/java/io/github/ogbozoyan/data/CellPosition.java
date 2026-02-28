package io.github.ogbozoyan.data;

/**
 * Position of a token/marker inside spreadsheet coordinates.
 *
 * @param sheetIndex  zero-based sheet index
 * @param sheetName   sheet name
 * @param rowIndex    zero-based row index
 * @param columnIndex zero-based column index
 */
public record CellPosition(int sheetIndex, String sheetName, int rowIndex, int columnIndex) {

    /**
     * Returns a diagnostic location string in {@code Sheet!R{row}C{col}} format.
     *
     * @return one-based location representation
     */
    public String asLocation() {
        return sheetName + "!R" + (rowIndex + 1) + "C" + (columnIndex + 1);
    }
}
