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
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.Objects;

/**
 * Default implementation of the report generation pipeline.
 */
@Service
@Slf4j
public class ReportGeneratorServiceImpl implements ReportGeneratorService {

    private final DocumentFormatConverter formatConverter;

    public ReportGeneratorServiceImpl() {
        this(new LibreOfficeDocumentFormatConverter());
        log.info("ReportGeneratorServiceImpl() - end: converter={}", this.formatConverter.getClass().getSimpleName());
    }

    public ReportGeneratorServiceImpl(DocumentFormatConverter formatConverter) {
        log.info("ReportGeneratorServiceImpl(DocumentFormatConverter) - start: converterClass={}",
            formatConverter == null ? null : formatConverter.getClass().getName());
        this.formatConverter = Objects.requireNonNull(formatConverter, "formatConverter must not be null");
        log.info("ReportGeneratorServiceImpl(DocumentFormatConverter) - end: converterClass={}",
            this.formatConverter.getClass().getName());
    }

    @Override
    public GeneratedReport generate(TemplateInput template, ReportData data, GenerateOptions options) {
        log.info("generate() - start: fileName={}, contentType={}, bytesLength={}, scalarTokens={}, requestedOptionsPresent={}",
            template == null ? null : template.fileName(),
            template == null ? null : template.contentType(),
            template == null || template.bytes() == null ? null : template.bytes().length,
            data == null || data.templateTokens() == null ? null : data.templateTokens().size(),
            options != null);
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
        log.info("generate() - formatResolved: sourceFormat={}, requestedOutputFormat={}, outputFormat={}",
            sourceFormat, requestedOutputFormat, outputFormat);

        WarningCollector warningCollector = new WarningCollector();

        try (WorkbookProcessor processor = createProcessor(sourceFormat, template.bytes())) {
            processor.applyTemplateTokens(data.templateTokens(), resolvedOptions, warningCollector);
            processor.recalculateFormulas(resolvedOptions);

            byte[] generatedBytes = processor.serialize();
            byte[] outputBytes = outputFormat == sourceFormat
                ? generatedBytes
                : formatConverter.convert(generatedBytes, sourceFormat, outputFormat);

            GeneratedReport report = ReportSerializer.serialize(template, outputFormat, outputBytes, warningCollector);
            log.info("generate() - end: outputFileName={}, outputContentType={}, outputBytesLength={}, warnings={}",
                report.fileName(), report.contentType(), report.bytes().length, report.warnings().size());
            return report;
        }
    }

    private WorkbookProcessor createProcessor(TemplateFormat format, byte[] bytes) {
        log.info("createProcessor() - start: format={}, bytesLength={}", format, bytes == null ? null : bytes.length);
        WorkbookProcessor processor = switch (format) {
            case XLS, XLSX -> new PoiWorkbookProcessor(bytes);
            case DOC -> new DocDocumentProcessor(bytes);
            case DOCX -> new DocxDocumentProcessor(bytes);
            case PDF -> new PdfDocumentProcessor(bytes);
            case ODS, ODT -> throw new UnsupportedTemplateFormatException(
                "ODS/ODT templates are not supported as input. " +
                "Use XLS/XLSX or DOC/DOCX templates and request ODS/ODT on output."
            );
        };
        log.info("createProcessor() - end: processorClass={}", processor.getClass().getName());
        return processor;
    }

    private void validateSourceTemplateFormat(TemplateFormat sourceFormat) {
        log.info("validateSourceTemplateFormat() - start: sourceFormat={}", sourceFormat);
        if (sourceFormat == TemplateFormat.ODS || sourceFormat == TemplateFormat.ODT) {
            throw new UnsupportedTemplateFormatException(
                "ODS/ODT templates are not supported as input. " +
                "Use XLS/XLSX or DOC/DOCX templates and request ODS/ODT on output."
            );
        }
        log.info("validateSourceTemplateFormat() - end: sourceFormat={}", sourceFormat);
    }

    private void validateOutputFormatConversion(TemplateFormat sourceFormat, TemplateFormat outputFormat) {
        log.info("validateOutputFormatConversion() - start: sourceFormat={}, outputFormat={}", sourceFormat, outputFormat);
        if (sourceFormat == outputFormat) {
            log.info("validateOutputFormatConversion() - end: sameFormat=true");
            return;
        }

        boolean spreadsheetToOds = isSpreadsheet(sourceFormat) && outputFormat == TemplateFormat.ODS;
        boolean wordToOdt = isWord(sourceFormat) && outputFormat == TemplateFormat.ODT;

        if (spreadsheetToOds || wordToOdt) {
            log.info("validateOutputFormatConversion() - end: conversionAllowed=true");
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
