package com.template.reportgenerator.service;

import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.GeneratedReport;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.TemplateFormat;
import com.template.reportgenerator.dto.TemplateInput;
import com.template.reportgenerator.processor.DocDocumentProcessor;
import com.template.reportgenerator.processor.DocxDocumentProcessor;
import com.template.reportgenerator.processor.OdsWorkbookProcessor;
import com.template.reportgenerator.processor.OdtDocumentProcessor;
import com.template.reportgenerator.processor.PdfDocumentProcessor;
import com.template.reportgenerator.processor.PoiWorkbookProcessor;
import com.template.reportgenerator.processor.WorkbookProcessor;
import com.template.reportgenerator.util.ReportSerializer;
import com.template.reportgenerator.util.TemplateFormatDetector;
import com.template.reportgenerator.util.WarningCollector;
import org.springframework.stereotype.Service;

/**
 * Default implementation of the report generation pipeline.
 */
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
            processor.applyScalarTokens(resolvedData.scalars(), resolvedOptions, warningCollector);
            processor.recalculateFormulas(resolvedOptions);

            byte[] output = processor.serialize();
            return ReportSerializer.serialize(template, format, output, warningCollector);
        }
    }

    private WorkbookProcessor createProcessor(TemplateFormat format, byte[] bytes) {
        return switch (format) {
            case XLS, XLSX -> new PoiWorkbookProcessor(bytes);
            case ODS -> new OdsWorkbookProcessor(bytes);
            case DOC -> new DocDocumentProcessor(bytes);
            case DOCX -> new DocxDocumentProcessor(bytes);
            case ODT -> new OdtDocumentProcessor(bytes);
            case PDF -> new PdfDocumentProcessor(bytes);
        };
    }
}
