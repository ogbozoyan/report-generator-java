package io.github.ogbozoyan.contract;

import java.util.Collections;
import java.util.Map;

/**
 * Runtime data model used for token resolution.
 * <p>
 * Table tokens are passed via {@link #templateTokens()} as {@code List<Map<String, Object>>}
 * values and rendered when the template contains an exact placeholder like
 * {@code {{TABLE_TOKEN}}}.
 *
 * @param templateTokens unified token context map
 */
public record ReportData(
    Map<String, Object> templateTokens
) {
    /**
     * Replaces {@code null} token map with immutable empty map.
     */
    public ReportData {
        templateTokens = templateTokens == null ? Collections.emptyMap() : templateTokens;
    }
}
