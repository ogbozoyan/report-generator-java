package io.github.ogbozoyan.helper;

import lombok.experimental.Helper;
import lombok.extern.slf4j.Slf4j;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

@Slf4j
@Helper
public class CommonHelper {
    /**
     * Builds stable column order: first-row keys, then new keys in encounter order.
     *
     * @param rows table rows
     * @return ordered columns
     */
    public static List<String> buildColumnOrder(List<Map<String, Object>> rows) {
        LinkedHashSet<String> ordered = new LinkedHashSet<>();
        if (!rows.isEmpty()) {
            ordered.addAll(rows.get(0).keySet());
        }
        for (Map<String, Object> row : rows) {
            ordered.addAll(row.keySet());
        }
        return List.copyOf(ordered);
    }
}
