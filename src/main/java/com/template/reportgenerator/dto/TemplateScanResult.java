package com.template.reportgenerator.dto;

import java.util.List;

public record TemplateScanResult(
    List<BlockMarker> markers,
    List<TokenOccurrence> scalarTokens
) {
}
