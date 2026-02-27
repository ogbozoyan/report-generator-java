package io.github.ogbozoyan.exception;

/**
 * Raised when template format cannot be recognized.
 */
public class UnsupportedTemplateFormatException extends RuntimeException {
    /**
     * Creates unsupported-format exception with detailed message.
     *
     * @param message format validation details
     */
    public UnsupportedTemplateFormatException(String message) {
        super(message);
    }
}
