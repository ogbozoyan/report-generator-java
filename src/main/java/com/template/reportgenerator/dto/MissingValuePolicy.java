package com.template.reportgenerator.dto;

/**
 * Behavior when a template token cannot be resolved from input data.
 */
public enum MissingValuePolicy {
    EMPTY_AND_LOG,
    LEAVE_TOKEN,
    FAIL_FAST
}
