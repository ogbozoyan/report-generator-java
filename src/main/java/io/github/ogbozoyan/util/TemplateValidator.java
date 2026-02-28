package io.github.ogbozoyan.util;


import io.github.ogbozoyan.contract.BlockMarker;
import io.github.ogbozoyan.contract.BlockRegion;
import io.github.ogbozoyan.contract.BlockType;
import io.github.ogbozoyan.contract.TemplateScanResult;
import io.github.ogbozoyan.exception.TemplateStructureException;
import io.github.ogbozoyan.exception.TemplateSyntaxException;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Validates block markers and builds executable block regions.
 */
@UtilityClass
@Slf4j
public class TemplateValidator {

    /**
     * Validates scanned markers and builds block regions.
     *
     * @param scanResult scanner output
     * @return validated non-overlapping block regions
     * @throws TemplateSyntaxException    when markers are unpaired or form invalid rectangles
     * @throws TemplateStructureException when regions overlap
     */
    public static List<BlockRegion> validateAndBuildRegions(TemplateScanResult scanResult) {
        log.debug("validateAndBuildRegions() - start: markers={}",
            scanResult == null || scanResult.markers() == null ? null : scanResult.markers().size());
        Map<MarkerKey, List<BlockMarker>> grouped = new HashMap<>();

        for (BlockMarker marker : scanResult.markers()) {
            MarkerKey key = new MarkerKey(
                marker.position().sheetIndex(),
                marker.blockType(),
                marker.key()
            );
            grouped.computeIfAbsent(key, ignored -> new ArrayList<>()).add(marker);
        }

        List<BlockRegion> regions = new ArrayList<>();

        for (Map.Entry<MarkerKey, List<BlockMarker>> entry : grouped.entrySet()) {
            List<BlockMarker> markers = entry.getValue();
            BlockMarker start = null;
            BlockMarker end = null;

            for (BlockMarker marker : markers) {
                if (marker.isStart()) {
                    if (start != null) {
                        throw new TemplateSyntaxException("More than one START marker for block " + entry.getKey());
                    }
                    start = marker;
                } else if (marker.isEnd()) {
                    if (end != null) {
                        throw new TemplateSyntaxException("More than one END marker for block " + entry.getKey());
                    }
                    end = marker;
                }
            }

            if (start == null || end == null) {
                throw new TemplateSyntaxException("Unpaired block markers for block " + entry.getKey());
            }

            if (start.position().rowIndex() >= end.position().rowIndex()
                || start.position().columnIndex() >= end.position().columnIndex()) {
                throw new TemplateSyntaxException(
                    "Invalid block rectangle for key '" + entry.getKey().key + "' at "
                    + start.position().asLocation() + " and " + end.position().asLocation()
                );
            }

            BlockRegion region = new BlockRegion(
                entry.getKey().blockType,
                entry.getKey().key,
                entry.getKey().sheetIndex,
                start.position().sheetName(),
                start.position().rowIndex(),
                start.position().columnIndex(),
                end.position().rowIndex(),
                end.position().columnIndex()
            );

            if (region.innerStartRow() > region.innerEndRow() || region.innerStartCol() > region.innerEndCol()) {
                throw new TemplateSyntaxException("Block has empty internal area: " + region.asLocation());
            }

            regions.add(region);
        }

        validateNoOverlaps(regions);
        log.trace("validateAndBuildRegions() - end: regions={}", regions.size());
        return regions;
    }

    /**
     * Ensures validated regions do not overlap on same sheet.
     *
     * @param regions validated regions
     */
    private void validateNoOverlaps(List<BlockRegion> regions) {
        for (int i = 0; i < regions.size(); i++) {
            for (int j = i + 1; j < regions.size(); j++) {
                BlockRegion a = regions.get(i);
                BlockRegion b = regions.get(j);

                if (a.sheetIndex() != b.sheetIndex()) {
                    continue;
                }

                boolean intersects = a.startRow() <= b.endRow()
                                     && b.startRow() <= a.endRow()
                                     && a.startCol() <= b.endCol()
                                     && b.startCol() <= a.endCol();

                if (intersects) {
                    throw new TemplateStructureException(
                        "Blocks overlap or nest: " + a.asLocation() + " and " + b.asLocation()
                    );
                }
            }
        }
    }

    /**
     * Compound key for grouping markers by logical block identity.
     *
     * @param sheetIndex sheet index
     * @param blockType  block type
     * @param key        block key
     */
    private record MarkerKey(int sheetIndex, BlockType blockType, String key) {

        @Override
        public boolean equals(Object obj) {
            if (this == obj) {
                return true;
            }

            if (!(obj instanceof MarkerKey other)) {
                return false;
            }
            return sheetIndex == other.sheetIndex
                   && blockType == other.blockType
                   && key.equals(other.key);
        }

        @Override
        public int hashCode() {
            int result = Integer.hashCode(sheetIndex);
            result = 31 * result + blockType.hashCode();
            result = 31 * result + key.hashCode();
            return result;
        }

        @Override
        public String toString() {
            return "sheet=" + sheetIndex + ", type=" + blockType + ", key='" + key + "'";
        }
    }
}
