package io.github.ogbozoyan.data;

import java.util.Collections;
import java.util.Map;

/**
 * Runtime data model used for token resolution.
 * <p>
 * Table tokens are passed via {@link #templateTokens()} and rendered when the template
 * contains an exact placeholder like
 * <ul>
 * <li>{@code {{TABLE_TOKEN}}} / {@code {{TOKEN}}} / {@code {{TOKEN_1}}} / {@code {{token}}}.</li>
 * </ul>
 * <p>
 * Supported payload shapes:
 * <ul>
 *     <li>{@code List<Map<String, Object>>} for default header+data table insertion.</li>
 *     <li>{@code List<Object[]>} when {@code GenerateOptions.rowsOnlyTableTokens=true}
 *     (XLS/XLSX rows-only insertion without header).</li>
 *     <li>{@code io.github.ogbozoyan.contract.TableBuilder} for declarative DOC/DOCX table insertion
 *     at exact placeholder token.</li>
 *     <li>{@code io.github.ogbozoyan.contract.TableXlsxBuilder} for declarative XLS/XLSX table insertion
 *     at exact placeholder token.</li>
 * </ul>
 * <p>
 * Optional table meta-keys can be provided in the same map:
 * <ul>
 *     <li>{@code TOKEN__columns} / {@code TOKEN_columns} / {@code TOKEN.columns}: explicit column order.</li>
 * </ul>
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
