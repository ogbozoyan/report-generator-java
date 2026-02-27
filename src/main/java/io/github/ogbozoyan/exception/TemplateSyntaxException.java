package io.github.ogbozoyan.exception;

/**
 * Raised when template DSL markers are syntactically invalid.
 */
public class TemplateSyntaxException extends RuntimeException {
    /**
     * Creates syntax exception with detailed message.
     *
     * @param message syntax violation details
     */
    public TemplateSyntaxException(String message) {
        super(message);
    }
}
