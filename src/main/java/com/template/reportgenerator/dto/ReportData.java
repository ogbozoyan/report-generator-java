package com.template.reportgenerator.dto;

import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Runtime data model used for token resolution.
 * <p>
 * Table tokens are passed via {@link #scalars()} as {@code List<Map<String, Object>>}
 * values and rendered when the template contains an exact placeholder like
 * {@code {{TABLE_TOKEN}}}.
 */
public record ReportData(
    Map<String, Object> scalars,
    /**
     * Legacy DSL payload, kept for binary compatibility and ignored by generation pipeline.
     */
    Map<String, List<Map<String, Object>>> tables,
    /**
     * Legacy DSL payload, kept for binary compatibility and ignored by generation pipeline.
     */
    Map<String, List<Map<String, Object>>> columns
) {
    public ReportData {
        scalars = scalars == null ? Collections.emptyMap() : scalars;
        tables = tables == null ? Collections.emptyMap() : tables;
        columns = columns == null ? Collections.emptyMap() : columns;
    }
}
