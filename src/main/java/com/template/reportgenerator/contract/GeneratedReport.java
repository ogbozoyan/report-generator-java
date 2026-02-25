package com.template.reportgenerator.contract;

import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * Generated artifact with output metadata and warnings.
 */
public record GeneratedReport(
    String fileName,
    String contentType,
    byte[] bytes,
    List<GenerationWarning> warnings
) {
    public GeneratedReport {
        Objects.requireNonNull(bytes, "bytes must not be null");
        warnings = warnings == null ? Collections.emptyList() : List.copyOf(warnings);
    }
}
