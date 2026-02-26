package com.template.reportgenerator.contract;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.List;
import java.util.Map;

/**
 * Deferred DOCX table insertion anchor.
 *
 * @param paragraph paragraph that contains table placeholder
 * @param token     table token name
 * @param rows      table payload rows
 * @param location  diagnostic location
 * @param order     traversal order, used for reverse application
 */
public record DocxTableAnchor(
    XWPFParagraph paragraph,
    String token,
    List<Map<String, Object>> rows,
    String location,
    int order
) {
}
