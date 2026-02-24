package com.template.reportgenerator.processor;

import com.template.reportgenerator.dto.BlockRegion;
import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.util.WarningCollector;

import java.util.List;
import java.util.Map;

/**
 * Format-specific processor abstraction used by the service pipeline.
 */
public interface WorkbookProcessor extends AutoCloseable {

    /**
     * Scans template for markers/tokens needed by validation stage.
     */
    TemplateScanResult scan();

    /**
     * Applies scalar and table tokens.
     */
    void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector);

    /**
     * Legacy TABLE DSL hook. Implementations keep it as no-op.
     */
    void expandTableBlocks(List<BlockRegion> tableBlocks, ReportData data, GenerateOptions options, WarningCollector warningCollector);

    /**
     * Legacy COL DSL hook. Implementations keep it as no-op.
     */
    void expandColumnBlocks(List<BlockRegion> columnBlocks, ReportData data, GenerateOptions options, WarningCollector warningCollector);

    /**
     * Legacy marker cleanup hook. Implementations keep it as no-op.
     */
    void clearMarkers(List<BlockRegion> blockRegions);

    /**
     * Recalculates formulas when the format supports it.
     */
    void recalculateFormulas(GenerateOptions options);

    /**
     * Serializes processed document into output bytes.
     */
    byte[] serialize();

    @Override
    default void close() {
        // nothing by default
    }
}
