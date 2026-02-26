package com.template.reportgenerator.exception;

/**
 * Raised on data binding failures (for example missing token in fail-fast mode).
 */
public class TemplateDataBindingException extends RuntimeException {
    /**
     * Creates binding exception with detailed message.
     *
     * @param message error description
     */
    public TemplateDataBindingException(String message) {
        super(message);
    }
}
