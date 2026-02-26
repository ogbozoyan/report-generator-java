package contract;

import java.util.List;
import java.util.Map;

/**
 * ODS table insertion anchor (retained for compatibility with older ODS path).
 *
 * @param rowIndex            zero-based row index of placeholder
 * @param colIndex            zero-based column index of placeholder
 * @param token               table token name
 * @param rows                normalized table rows
 * @param styleName           baseline style name
 * @param horizontalAlignment baseline horizontal alignment
 * @param verticalAlignment   baseline vertical alignment
 * @param wrapped             baseline wrap flag
 * @param rowHeight           baseline row height
 * @param rowOptimalHeight    baseline optimal-height flag
 */
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
