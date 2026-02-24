package com.template.reportgenerator.dto;

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
    public int innerStartRow() {
        return startRow + 1;
    }

    public int innerEndRow() {
        return endRow - 1;
    }

    public int innerStartCol() {
        return startCol + 1;
    }

    public int innerEndCol() {
        return endCol - 1;
    }

    public String asLocation() {
        return sheetName + "!R" + (startRow + 1) + "C" + (startCol + 1)
               + "..R" + (endRow + 1) + "C" + (endCol + 1);
    }
}
