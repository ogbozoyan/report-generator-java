package com.template.reportgenerator.exception;

/**
 * Raised when template DSL markers are syntactically invalid.
 */
public class TemplateSyntaxException extends RuntimeException {
    public TemplateSyntaxException(String message) {
        super(message);
    }
}
