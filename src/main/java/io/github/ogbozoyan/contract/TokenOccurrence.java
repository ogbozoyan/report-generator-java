package io.github.ogbozoyan.contract;

/**
 * Scalar token found in template and its location.
 *
 * @param token    token name without braces
 * @param position token location
 */
public record TokenOccurrence(String token, CellPosition position) {
}
