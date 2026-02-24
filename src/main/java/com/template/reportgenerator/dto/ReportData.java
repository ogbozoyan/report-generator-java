package com.template.reportgenerator.dto;

import java.util.Collections;
import java.util.Map;

/**
 * Runtime data model used for token resolution.
 * <p>
 * Table tokens are passed via {@link #scalars()} as {@code List<Map<String, Object>>}
 * values and rendered when the template contains an exact placeholder like
 * {@code {{TABLE_TOKEN}}}.
 */
public record ReportData(
    Map<String, Object> scalars
) {
    public ReportData {
        scalars = scalars == null ? Collections.emptyMap() : scalars;
    }
}
