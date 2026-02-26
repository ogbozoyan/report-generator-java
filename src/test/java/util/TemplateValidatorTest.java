package util;

import contract.BlockMarker;
import contract.BlockRegion;
import contract.BlockType;
import contract.CellPosition;
import contract.TemplateScanResult;
import exception.TemplateStructureException;
import exception.TemplateSyntaxException;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static util.TemplateValidator.validateAndBuildRegions;

class TemplateValidatorTest {

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
