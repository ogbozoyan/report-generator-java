package com.template.reportgenerator.dto;

import java.util.Collections;
import java.util.List;
import java.util.Map;

public record ReportData(
    Map<String, Object> scalars,
    Map<String, List<Map<String, Object>>> tables,
    Map<String, List<Map<String, Object>>> columns
) {
    public ReportData {
        scalars = scalars == null ? Collections.emptyMap() : scalars;
        tables = tables == null ? Collections.emptyMap() : tables;
        columns = columns == null ? Collections.emptyMap() : columns;
    }
}
