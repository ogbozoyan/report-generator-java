package io.github.ogbozoyan.data;

/**
 * Non-fatal generation warning.
 *
 * @param code     stable warning code
 * @param message  human-readable message
 * @param location template location where warning occurred
 */
public record GenerationWarning(String code, String message, String location) {
}
