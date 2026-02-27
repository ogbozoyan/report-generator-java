package io.github.ogbozoyan.contract;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.List;
import java.util.Map;

/**
 * Deferred POI table insertion anchor.
 *
 * @param rowIndex              zero-based anchor row
 * @param colIndex              zero-based anchor column
 * @param token                 table token name
 * @param rows                  table payload rows
 * @param baselineStyle         style copied from marker cell
 * @param baselineRowHeight     row height copied from marker row
 * @param configuredColumnOrder optional explicitly configured column order
 */
public record PoiTableAnchor(
    int rowIndex,
    int colIndex,
    String token,
    List<Map<String, Object>> rows,
    CellStyle baselineStyle,
    short baselineRowHeight,
    List<String> configuredColumnOrder
) {
}
