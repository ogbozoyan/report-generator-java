package contract;

/**
 * Validated rectangular region for a legacy TABLE/COL block.
 *
 * @param blockType block type
 * @param key logical block key
 * @param sheetIndex zero-based sheet index
 * @param sheetName sheet name
 * @param startRow zero-based start row (inclusive)
 * @param startCol zero-based start column (inclusive)
 * @param endRow zero-based end row (inclusive)
 * @param endCol zero-based end column (inclusive)
 */
public record BlockRegion(
    BlockType blockType,
    String key,
    int sheetIndex,
    String sheetName,
    int startRow,
    int startCol,
    int endRow,
    int endCol
) {
    /**
     * Returns first row inside block bounds (excluding marker row).
     *
     * @return zero-based row index
     */
    public int innerStartRow() {
        return startRow + 1;
    }

    /**
     * Returns last row inside block bounds (excluding marker row).
     *
     * @return zero-based row index
     */
    public int innerEndRow() {
        return endRow - 1;
    }

    /**
     * Returns first column inside block bounds (excluding marker column).
     *
     * @return zero-based column index
     */
    public int innerStartCol() {
        return startCol + 1;
    }

    /**
     * Returns last column inside block bounds (excluding marker column).
     *
     * @return zero-based column index
     */
    public int innerEndCol() {
        return endCol - 1;
    }

    /**
     * Builds a human-readable location string for diagnostics.
     *
     * @return formatted block location
     */
    public String asLocation() {
        return sheetName + "!R" + (startRow + 1) + "C" + (startCol + 1)
               + "..R" + (endRow + 1) + "C" + (endCol + 1);
    }
}
