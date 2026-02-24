package com.template.reportgenerator.dto;

import java.util.List;

/**
 * Result of template scan containing markers and scalar token occurrences.
 */
public record TemplateScanResult(
    List<BlockMarker> markers,
    List<TokenOccurrence> scalarTokens
) {
}
