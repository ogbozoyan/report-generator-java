package com.template.reportgenerator.dto;

/**
 * Non-fatal generation warning.
 */
public record GenerationWarning(String code, String message, String location) {
}
