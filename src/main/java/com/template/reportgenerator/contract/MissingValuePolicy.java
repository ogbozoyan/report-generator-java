package com.template.reportgenerator.contract;

/**
 * Behavior when a template token cannot be resolved from input data.
 */
public enum MissingValuePolicy {
    EMPTY_AND_LOG,
    LEAVE_TOKEN,
    FAIL_FAST
}
