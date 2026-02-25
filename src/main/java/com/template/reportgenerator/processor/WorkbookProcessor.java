package com.template.reportgenerator.processor;

import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.util.WarningCollector;

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
     * Recalculates formulas when the format supports it.
     */
    default void recalculateFormulas(GenerateOptions options) {

    }

    /**
     * Serializes processed document into output bytes.
     */
    byte[] serialize();

    @Override
    default void close() {
        // nothing by default
    }
}
