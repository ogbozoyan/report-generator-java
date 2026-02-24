package com.template.reportgenerator.exception;

/**
 * Raised when template format cannot be recognized.
 */
public class UnsupportedTemplateFormatException extends RuntimeException {
    public UnsupportedTemplateFormatException(String message) {
        super(message);
    }
}
