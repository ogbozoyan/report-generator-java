package com.template.reportgenerator.dto;

/**
 * Scalar token found in template and its location.
 */
public record TokenOccurrence(String token, CellPosition position) {
}
