package io.github.ogbozoyan.util;

import io.github.ogbozoyan.BaseTest;
import io.github.ogbozoyan.data.BlockMarker;
import io.github.ogbozoyan.data.BlockRegion;
import io.github.ogbozoyan.data.BlockType;
import io.github.ogbozoyan.data.CellPosition;
import io.github.ogbozoyan.data.TemplateScanResult;
import io.github.ogbozoyan.exception.TemplateStructureException;
import io.github.ogbozoyan.exception.TemplateSyntaxException;
import org.junit.jupiter.api.Test;

import java.util.List;

import static io.github.ogbozoyan.util.TemplateValidator.validateAndBuildRegions;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class TemplateValidatorTest extends BaseTest {

    @Test
    void shouldBuildRegionsForValidMarkers() {
        TemplateScanResult scan = new TemplateScanResult(List.of(
            new BlockMarker(BlockType.TABLE, "START", "rows", new CellPosition(0, "S", 0, 0)),
            new BlockMarker(BlockType.TABLE, "END", "rows", new CellPosition(0, "S", 2, 2))
        ), List.of());

        List<BlockRegion> regions = validateAndBuildRegions(scan);
        assertEquals(1, regions.size());
        assertEquals("rows", regions.get(0).key());
        assertEquals(BlockType.TABLE, regions.get(0).blockType());
    }

    @Test
    void shouldFailOnUnpairedMarkers() {
        TemplateScanResult scan = new TemplateScanResult(List.of(
            new BlockMarker(BlockType.TABLE, "START", "rows", new CellPosition(0, "S", 0, 0))
        ), List.of());

        assertThrows(TemplateSyntaxException.class, () -> validateAndBuildRegions(scan));
    }

    @Test
    void shouldFailOnOverlappingBlocks() {
        TemplateScanResult scan = new TemplateScanResult(List.of(
            new BlockMarker(BlockType.TABLE, "START", "a", new CellPosition(0, "S", 0, 0)),
            new BlockMarker(BlockType.TABLE, "END", "a", new CellPosition(0, "S", 3, 3)),
            new BlockMarker(BlockType.COL, "START", "b", new CellPosition(0, "S", 1, 1)),
            new BlockMarker(BlockType.COL, "END", "b", new CellPosition(0, "S", 4, 4))
        ), List.of());

        assertThrows(TemplateStructureException.class, () -> validateAndBuildRegions(scan));
    }
}
