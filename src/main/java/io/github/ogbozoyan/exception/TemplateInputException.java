package io.github.ogbozoyan.exception;

/**
 * Raised when template input is null
 */
public class TemplateInputException extends RuntimeException {
    /**
     * Creates unsupported-format exception with detailed message.
     *
     * @param message format validation details
     */
    public TemplateInputException(String message) {
        super(message);
    }
}
