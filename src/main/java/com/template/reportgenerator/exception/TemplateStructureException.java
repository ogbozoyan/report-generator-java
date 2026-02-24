package com.template.reportgenerator.exception;

/**
 * Raised when template blocks overlap or nest in unsupported way.
 */
public class TemplateStructureException extends RuntimeException {
    public TemplateStructureException(String message) {
        super(message);
    }
}
