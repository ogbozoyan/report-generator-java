package com.template.reportgenerator.util;

import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.ResolvedText;
import lombok.experimental.UtilityClass;

import java.util.Map;
import java.util.regex.Pattern;

/**
 * Shared text replacement utility for non-spreadsheet formats.
 */
@UtilityClass
public class TextTemplateEngine {

    private static final Pattern BLOCK_MARKER_PATTERN = Pattern.compile(
        "\\[\\[\\s*(TABLE_START|TABLE_END|COL_START|COL_END)\\s*:[a-zA-Z0-9_.-]+\\s*]]"
    );

    /**
     * Replaces scalar tokens in arbitrary text and removes unsupported TABLE/COL markers.
     */
    public static String replaceText(
        String source,
        Map<String, Object> scalars,
        GenerateOptions options,
        WarningCollector warningCollector,
        String location
    ) {
        if (source == null || source.isEmpty()) {
            return source;
        }

        String withoutMarkers = source;
        var markerMatcher = BLOCK_MARKER_PATTERN.matcher(source);
        if (markerMatcher.find()) {
            warningCollector.add(
                "BLOCK_MARKER_IGNORED",
                "TABLE/COL markers are ignored for this template format",
                location
            );
            withoutMarkers = markerMatcher.replaceAll("");
        }

        ResolvedText resolvedText = TokenResolver.resolve(
            withoutMarkers,
            scalars,
            options.missingValuePolicy(),
            warningCollector,
            location,
            false
        );

        if (containsItemLevelTokens(resolvedText.value())) {
            warningCollector.add(
                "ITEM_TOKEN_IGNORED",
                "item/index tokens are ignored for this template format",
                location
            );
        }

        return resolvedText.value();
    }

    private static boolean containsItemLevelTokens(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }
        return value.contains("{{item.") || value.contains("{{ index }}") || value.contains("{{index}}");
    }
}
