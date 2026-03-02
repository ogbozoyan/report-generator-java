package io.github.ogbozoyan.data;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 * Immutable paragraph processing target.
 *
 * @param paragraph paragraph object
 * @param text      paragraph plain text
 * @param location  diagnostic location
 * @param order     traversal order
 */
public record ParagraphTarget(XWPFParagraph paragraph, String text, String location, int order) {
}
