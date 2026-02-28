package io.github.ogbozoyan.data;

/**
 * Text replacement result flagging whether source changed.
 *
 * @param value   resolved text value
 * @param changed {@code true} when replacement modified source
 */
public record ResolvedText(String value, boolean changed) {
}
