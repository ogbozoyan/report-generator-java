package com.template.reportgenerator.contract;

import java.util.List;
import java.util.Map;

public record OdsTableAnchor(
    int rowIndex,
    int colIndex,
    String token,
    List<Map<String, Object>> rows,
    String styleName,
    String horizontalAlignment,
    String verticalAlignment,
    boolean wrapped,
    long rowHeight,
    boolean rowOptimalHeight
) {
}
