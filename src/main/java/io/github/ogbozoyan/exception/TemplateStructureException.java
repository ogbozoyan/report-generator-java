package io.github.ogbozoyan.exception;

/**
 * Raised when template blocks overlap or nest in unsupported way.
 */
public class TemplateStructureException extends RuntimeException {
    /**
     * Creates structure exception with detailed message.
     *
     * @param message structure violation details
     */
    public TemplateStructureException(String message) {
        super(message);
    }
}
