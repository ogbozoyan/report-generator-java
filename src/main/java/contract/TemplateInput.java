package contract;

import java.util.Objects;

/**
 * Input template descriptor.
 *
 * @param fileName    original template filename (optional)
 * @param contentType template content type (optional)
 * @param bytes       raw template bytes
 */
public record TemplateInput(
    String fileName,
    String contentType,
    byte[] bytes
) {
    public TemplateInput {
        Objects.requireNonNull(bytes, "bytes must not be null");
    }
}
