package com.template.reportgenerator.contract;

import java.util.Collections;
import java.util.Map;

/**
 * Runtime data model used for token resolution.
 * <p>
 * Table tokens are passed via {@link #templateTokens()} as {@code List<Map<String, Object>>}
 * values and rendered when the template contains an exact placeholder like
 * {@code {{TABLE_TOKEN}}}.
 */
public record ReportData(
    Map<String, Object> templateTokens
) {
    public ReportData {
        templateTokens = templateTokens == null ? Collections.emptyMap() : templateTokens;
    }
}
