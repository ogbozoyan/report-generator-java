package com.template.reportgenerator.dto;

public record BlockMarker(BlockType blockType, String markerRole, String key, CellPosition position) {

    public boolean isStart() {
        return "START".equals(markerRole);
    }

    public boolean isEnd() {
        return "END".equals(markerRole);
    }
}
