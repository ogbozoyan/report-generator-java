package com.template.reportgenerator.exception;

/**
 * Raised when source template cannot be read or output cannot be serialized.
 */
public class TemplateReadWriteException extends RuntimeException {
    /**
     * Creates exception with message only.
     *
     * @param message error description
     */
    public TemplateReadWriteException(String message) {
        super(message);
    }

    /**
     * Creates exception with message and original cause.
     *
     * @param message error description
     * @param cause   original failure
     */
    public TemplateReadWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
