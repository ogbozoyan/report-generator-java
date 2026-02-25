package com.template.reportgenerator.contract;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.List;
import java.util.Map;

public record PoiTableAnchor(
    int rowIndex,
    int colIndex,
    String token,
    List<Map<String, Object>> rows,
    CellStyle baselineStyle,
    short baselineRowHeight
) {
}
