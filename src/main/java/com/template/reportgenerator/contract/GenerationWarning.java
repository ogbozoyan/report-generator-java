package com.template.reportgenerator.contract;

/**
 * Non-fatal generation warning.
 */
public record GenerationWarning(String code, String message, String location) {
}
