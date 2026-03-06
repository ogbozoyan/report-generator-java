package io.github.ogbozoyan.data;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 * Deferred DOCX table insertion anchor.
 *
 * @param paragraph paragraph that contains table placeholder
 * @param token     table token name
 * @param tablePayload table payload object:
 *                     {@code List<Map<String,Object>>} or declarative
 *                     {@code io.github.ogbozoyan.contract.TableBuilder} /
 *                     {@code io.github.ogbozoyan.contract.RowBuilder}
 * @param location  diagnostic location
 * @param order     traversal order, used for reverse application
 */
public record DocxTableAnchor(
    XWPFParagraph paragraph,
    String token,
    Object tablePayload,
    String location,
    int order
) {
}
