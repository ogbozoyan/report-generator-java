package io.github.ogbozoyan.data;

/**
 * Raw legacy block marker discovered during scan phase.
 *
 * @param blockType  marker block type ({@link BlockType#TABLE} or {@link BlockType#COL})
 * @param markerRole marker role, usually {@code START} or {@code END}
 * @param key        logical block key used for marker pairing
 * @param position   marker location in template coordinates
 */
public record BlockMarker(BlockType blockType, String markerRole, String key, CellPosition position) {

    /**
     * Checks whether marker is a block start marker.
     *
     * @return {@code true} for {@code START}
     */
    public boolean isStart() {
        return "START".equals(markerRole);
    }

    /**
     * Checks whether marker is a block end marker.
     *
     * @return {@code true} for {@code END}
     */
    public boolean isEnd() {
        return "END".equals(markerRole);
    }
}
