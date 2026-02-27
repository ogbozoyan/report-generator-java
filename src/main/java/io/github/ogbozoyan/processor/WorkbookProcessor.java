package io.github.ogbozoyan.processor;

import io.github.ogbozoyan.contract.GenerateOptions;
import io.github.ogbozoyan.contract.TemplateScanResult;
import io.github.ogbozoyan.util.WarningCollector;

import java.util.Map;

/**
 * Format-specific io.github.ogbozoyan.processor abstraction used by io.github.ogbozoyan.service pipeline.
 *
 * <p>Each implementation owns one document instance and applies the same lifecycle:
 * <ol>
 *     <li>optional {@link #scan()} phase,</li>
 *     <li>{@link #applyTemplateTokens(Map, GenerateOptions, WarningCollector)},</li>
 *     <li>optional {@link #recalculateFormulas(GenerateOptions)},</li>
 *     <li>{@link #serialize()},</li>
 *     <li>{@link #close()}.</li>
 * </ol>
 */
public interface WorkbookProcessor extends AutoCloseable {

    /**
     * Scans template for markers/tokens needed by validation stage.
     *
     * @return scan result with markers and scalar token occurrences
     */
    TemplateScanResult scan();

    /**
     * Applies scalar and table tokens.
     *
     * @param scalars unified token map; table token values are expected as
     *                {@code List<Map<String, Object>>}
     * @param options generation options
     * @param warningCollector collector for non-fatal generation warnings
     */
    void applyTemplateTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector);

    /**
     * Recalculates formulas when the format supports it.
     *
     * @param options generation options
     */
    default void recalculateFormulas(GenerateOptions options) {

    }

    /**
     * Serializes processed document into output bytes.
     *
     * @return generated document bytes
     */
    byte[] serialize();

    /**
     * Releases io.github.ogbozoyan.processor resources.
     */
    @Override
    default void close() {
        // nothing by default
    }
}
