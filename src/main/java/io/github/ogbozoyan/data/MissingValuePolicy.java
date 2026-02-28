package io.github.ogbozoyan.data;

/**
 * Behavior when a template token cannot be resolved from input data.
 */
public enum MissingValuePolicy {
    /**
     * Replace unresolved token with empty string and collect warning.
     */
    EMPTY_AND_LOG,
    /**
     * Keep unresolved token text unchanged in output.
     */
    LEAVE_TOKEN,
    /**
     * Stop generation immediately with {@code TemplateDataBindingException}.
     */
    FAIL_FAST
}
