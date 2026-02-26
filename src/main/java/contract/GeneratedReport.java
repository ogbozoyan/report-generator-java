package contract;

import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * Generated artifact with output metadata and warnings.
 *
 * @param fileName output file name
 * @param contentType output MIME type
 * @param bytes generated file bytes
 * @param warnings non-fatal generation warnings
 */
public record GeneratedReport(
    String fileName,
    String contentType,
    byte[] bytes,
    List<GenerationWarning> warnings
) {
    /**
     * Validates mandatory fields and normalizes warning list.
     */
    public GeneratedReport {
        Objects.requireNonNull(bytes, "bytes must not be null");
        warnings = warnings == null ? Collections.emptyList() : List.copyOf(warnings);
    }
}
