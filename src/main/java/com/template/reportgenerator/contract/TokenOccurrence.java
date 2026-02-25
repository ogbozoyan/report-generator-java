package com.template.reportgenerator.contract;

/**
 * Scalar token found in template and its location.
 */
public record TokenOccurrence(String token, CellPosition position) {
}
