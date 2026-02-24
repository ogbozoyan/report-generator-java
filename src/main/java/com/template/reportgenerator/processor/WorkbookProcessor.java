package com.template.reportgenerator.processor;

import com.template.reportgenerator.dto.BlockRegion;
import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.util.WarningCollector;

import java.util.List;
import java.util.Map;

public interface WorkbookProcessor extends AutoCloseable {

    TemplateScanResult scan();

    void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector);

    void expandTableBlocks(List<BlockRegion> tableBlocks, ReportData data, GenerateOptions options, WarningCollector warningCollector);

    void expandColumnBlocks(List<BlockRegion> columnBlocks, ReportData data, GenerateOptions options, WarningCollector warningCollector);

    void clearMarkers(List<BlockRegion> blockRegions);

    void recalculateFormulas(GenerateOptions options);

    byte[] serialize();

    @Override
    default void close() {
        // nothing by default
    }
}
