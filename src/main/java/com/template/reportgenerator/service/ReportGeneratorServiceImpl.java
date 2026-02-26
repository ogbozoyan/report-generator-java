package com.template.reportgenerator.service;

import com.template.reportgenerator.contract.GenerateOptions;
import com.template.reportgenerator.contract.GeneratedReport;
import com.template.reportgenerator.contract.ReportData;
import com.template.reportgenerator.contract.TemplateFormat;
import com.template.reportgenerator.contract.TemplateInput;
import com.template.reportgenerator.exception.UnsupportedTemplateFormatException;
import com.template.reportgenerator.processor.DocDocumentProcessor;
import com.template.reportgenerator.processor.DocxDocumentProcessor;
import com.template.reportgenerator.processor.PdfDocumentProcessor;
import com.template.reportgenerator.processor.PoiWorkbookProcessor;
import com.template.reportgenerator.processor.WorkbookProcessor;
import com.template.reportgenerator.util.DocumentFormatConverter;
import com.template.reportgenerator.util.LibreOfficeDocumentFormatConverter;
import com.template.reportgenerator.util.ReportSerializer;
import com.template.reportgenerator.util.TemplateFormatDetector;
import com.template.reportgenerator.util.WarningCollector;
import org.springframework.stereotype.Service;

import java.util.Objects;

/**
 * Default implementation of the report generation pipeline.
 */
@Service
public class ReportGeneratorServiceImpl implements ReportGeneratorService {

    private final DocumentFormatConverter formatConverter;

    public ReportGeneratorServiceImpl() {
        this(new LibreOfficeDocumentFormatConverter());
    }

    public ReportGeneratorServiceImpl(DocumentFormatConverter formatConverter) {
        this.formatConverter = Objects.requireNonNull(formatConverter, "formatConverter must not be null");
    }

    @Override
    public GeneratedReport generate(TemplateInput template, ReportData data, GenerateOptions options) {
        if (template == null) {
            throw new IllegalArgumentException("template must not be null");
        }
        if (data == null) {
            throw new IllegalArgumentException("data must not be null");
        }

        GenerateOptions resolvedOptions = options == null
            ? GenerateOptions.defaults()
            : options;

        TemplateFormat sourceFormat = TemplateFormatDetector.detectFormat(template);
        validateSourceTemplateFormat(sourceFormat);

        TemplateFormat requestedOutputFormat = TemplateFormatDetector.detectRequestedOutputFormat(template);
        TemplateFormat outputFormat = requestedOutputFormat == null ? sourceFormat : requestedOutputFormat;
        validateOutputFormatConversion(sourceFormat, outputFormat);

        WarningCollector warningCollector = new WarningCollector();

        try (WorkbookProcessor processor = createProcessor(sourceFormat, template.bytes())) {
            processor.applyTemplateTokens(data.templateTokens(), resolvedOptions, warningCollector);
            processor.recalculateFormulas(resolvedOptions);

            byte[] generatedBytes = processor.serialize();
            byte[] outputBytes = outputFormat == sourceFormat
                ? generatedBytes
                : formatConverter.convert(generatedBytes, sourceFormat, outputFormat);

            return ReportSerializer.serialize(template, outputFormat, outputBytes, warningCollector);
        }
    }

    private WorkbookProcessor createProcessor(TemplateFormat format, byte[] bytes) {
        return switch (format) {
            case XLS, XLSX -> new PoiWorkbookProcessor(bytes);
            case DOC -> new DocDocumentProcessor(bytes);
            case DOCX -> new DocxDocumentProcessor(bytes);
            case PDF -> new PdfDocumentProcessor(bytes);
            case ODS, ODT -> throw new UnsupportedTemplateFormatException(
                "ODS/ODT templates are not supported as input. " +
                "Use XLS/XLSX or DOC/DOCX templates and request ODS/ODT on output."
            );
        };
    }

    private void validateSourceTemplateFormat(TemplateFormat sourceFormat) {
        if (sourceFormat == TemplateFormat.ODS || sourceFormat == TemplateFormat.ODT) {
            throw new UnsupportedTemplateFormatException(
                "ODS/ODT templates are not supported as input. " +
                "Use XLS/XLSX or DOC/DOCX templates and request ODS/ODT on output."
            );
        }
    }

    private void validateOutputFormatConversion(TemplateFormat sourceFormat, TemplateFormat outputFormat) {
        if (sourceFormat == outputFormat) {
            return;
        }

        boolean spreadsheetToOds = isSpreadsheet(sourceFormat) && outputFormat == TemplateFormat.ODS;
        boolean wordToOdt = isWord(sourceFormat) && outputFormat == TemplateFormat.ODT;

        if (spreadsheetToOds || wordToOdt) {
            return;
        }

        throw new UnsupportedTemplateFormatException(
            "Unsupported output conversion: " + sourceFormat + " -> " + outputFormat + ". " +
            "Supported conversions: XLS/XLSX -> ODS and DOC/DOCX -> ODT."
        );
    }

    private boolean isSpreadsheet(TemplateFormat format) {
        return format == TemplateFormat.XLS || format == TemplateFormat.XLSX;
    }

    private boolean isWord(TemplateFormat format) {
        return format == TemplateFormat.DOC || format == TemplateFormat.DOCX;
    }
}
