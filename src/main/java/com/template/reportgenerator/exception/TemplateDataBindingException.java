package com.template.reportgenerator.exception;

/**
 * Raised on data binding failures (for example missing token in fail-fast mode).
 */
public class TemplateDataBindingException extends RuntimeException {
    public TemplateDataBindingException(String message) {
        super(message);
    }
}
