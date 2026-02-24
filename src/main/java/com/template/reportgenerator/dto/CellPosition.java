package com.template.reportgenerator.dto;

public record CellPosition(int sheetIndex, String sheetName, int rowIndex, int columnIndex) {

    public String asLocation() {
        return sheetName + "!R" + (rowIndex + 1) + "C" + (columnIndex + 1);
    }
}
