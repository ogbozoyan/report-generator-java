package io.github.ogbozoyan.contract;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * Mutable aggregation state for single sheet processing pass.
 */
@Data
public final class SheetProcessingState {
    private final List<PoiTableAnchor> anchors = new ArrayList<>();
    private int processedCells;
    private int tableTokensFound;
    private int scalarTokensApplied;

    public void incrementProcessedCells() {
        processedCells++;
    }

    public void incrementScalarTokensApplied() {
        scalarTokensApplied++;
    }

    public void incrementTableTokensFound() {
        tableTokensFound++;
    }
}
