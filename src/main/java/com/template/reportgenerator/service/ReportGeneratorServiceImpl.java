package com.template.reportgenerator.service;

import com.template.reportgenerator.dto.BlockRegion;
import com.template.reportgenerator.dto.BlockType;
import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.GeneratedReport;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.TemplateFormat;
import com.template.reportgenerator.dto.TemplateInput;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.processor.OdsWorkbookProcessor;
import com.template.reportgenerator.processor.PoiWorkbookProcessor;
import com.template.reportgenerator.processor.WorkbookProcessor;
import com.template.reportgenerator.util.ReportSerializer;
import com.template.reportgenerator.util.TemplateFormatDetector;
import com.template.reportgenerator.util.TemplateValidator;
import com.template.reportgenerator.util.WarningCollector;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ReportGeneratorServiceImpl implements ReportGeneratorService {

    @Override
    public GeneratedReport generate(TemplateInput template, ReportData data, GenerateOptions options) {
        if (template == null) {
            throw new IllegalArgumentException("template must not be null");
        }

        ReportData resolvedData = data == null
            ? new ReportData(null, null, null)
            : data;

        GenerateOptions resolvedOptions = options == null
            ? GenerateOptions.defaults()
            : options;

        TemplateFormat format = TemplateFormatDetector.detect(template);
        WarningCollector warningCollector = new WarningCollector();

        try (WorkbookProcessor processor = createProcessor(format, template.bytes())) {
            TemplateScanResult scanResult = processor.scan();
            List<BlockRegion> blockRegions = TemplateValidator.validateAndBuildRegions(scanResult);

            List<BlockRegion> tableBlocks = blockRegions.stream()
                .filter(region -> region.blockType() == BlockType.TABLE)
                .toList();

            List<BlockRegion> columnBlocks = blockRegions.stream()
                .filter(region -> region.blockType() == BlockType.COL)
                .toList();

            processor.applyScalarTokens(resolvedData.scalars(), resolvedOptions, warningCollector);
            processor.expandTableBlocks(tableBlocks, resolvedData, resolvedOptions, warningCollector);
            processor.expandColumnBlocks(columnBlocks, resolvedData, resolvedOptions, warningCollector);
            processor.clearMarkers(blockRegions);
            processor.recalculateFormulas(resolvedOptions);

            byte[] output = processor.serialize();
            return ReportSerializer.serialize(template, format, output, warningCollector);
        }
    }

    private WorkbookProcessor createProcessor(TemplateFormat format, byte[] bytes) {
        return switch (format) {
            case XLS, XLSX -> new PoiWorkbookProcessor(bytes);
            case ODS -> new OdsWorkbookProcessor(bytes);
        };
    }
}
