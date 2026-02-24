package com.template.reportgenerator.exception;

/**
 * Raised when source template cannot be read or output cannot be serialized.
 */
public class TemplateReadWriteException extends RuntimeException {
    public TemplateReadWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
