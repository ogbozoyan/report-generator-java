package io.github.ogbozoyan.data;

import java.time.ZoneId;
import java.util.Locale;

/**
 * Generation options controlling token fallback behavior and localization.
 *
 * @param missingValuePolicy  policy for unresolved tokens
 * @param recalculateFormulas {@code true} to evaluate spreadsheet formulas after token application
 * @param rowsOnlyTableTokens {@code true} to insert XLS/XLSX table tokens as data rows only (without header row)
 * @param locale              locale hint for locale-sensitive formatting
 * @param zoneId              time-zone used for date/time conversions
 */
public record GenerateOptions(
    MissingValuePolicy missingValuePolicy,
    boolean recalculateFormulas,
    Locale locale,
    ZoneId zoneId,
    boolean rowsOnlyTableTokens
) {
    /**
     * Normalizes nullable options into safe defaults.
     */
    public GenerateOptions {
        missingValuePolicy = missingValuePolicy == null ? MissingValuePolicy.EMPTY_AND_LOG : missingValuePolicy;
        locale = locale == null ? Locale.getDefault() : locale;
        zoneId = zoneId == null ? ZoneId.systemDefault() : zoneId;
    }

    /**
     * Returns default generation options.
     *
     * <p>Equivalent to:
     * <pre>{@code
     * new GenerateOptions(
     *     MissingValuePolicy.EMPTY_AND_LOG,
     *     true,
     *     Locale.getDefault(),
     *     ZoneId.systemDefault(),
     *     false
     * );
     * }</pre>
     *
     * @return default options
     */
    public static GenerateOptions defaults() {
        return new GenerateOptions(MissingValuePolicy.EMPTY_AND_LOG,
            true,
            Locale.getDefault(),
            ZoneId.systemDefault(),
            false
        );
    }
}
