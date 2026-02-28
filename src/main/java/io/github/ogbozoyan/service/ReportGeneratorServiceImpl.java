package io.github.ogbozoyan.service;

import io.github.ogbozoyan.data.GenerateOptions;
import io.github.ogbozoyan.data.GeneratedReport;
import io.github.ogbozoyan.data.ReportData;
import io.github.ogbozoyan.data.TemplateFormat;
import io.github.ogbozoyan.data.TemplateInput;
import io.github.ogbozoyan.exception.UnsupportedTemplateFormatException;
import io.github.ogbozoyan.processor.DocDocumentProcessor;
import io.github.ogbozoyan.processor.DocxDocumentProcessor;
import io.github.ogbozoyan.processor.PdfDocumentProcessor;
import io.github.ogbozoyan.processor.PoiWorkbookProcessor;
import io.github.ogbozoyan.processor.WorkbookProcessor;
import io.github.ogbozoyan.util.DocumentFormatConverter;
import io.github.ogbozoyan.util.LibreOfficeDocumentFormatConverter;
import io.github.ogbozoyan.util.ReportSerializer;
import io.github.ogbozoyan.util.TemplateFormatDetector;
import io.github.ogbozoyan.util.WarningCollector;
import lombok.extern.slf4j.Slf4j;

import java.util.Objects;

/**
 * Default implementation of report generation pipeline.
 *
 * <p>Lifecycle:
 * <ol>
 *     <li>detect source template format,</li>
 *     <li>resolve requested output format,</li>
 *     <li>validate conversion route,</li>
 *     <li>apply tokens in format-specific io.github.ogbozoyan.processor,</li>
 *     <li>recalculate formulas where supported,</li>
 *     <li>optionally convert output bytes to requested format.</li>
 * </ol>
 */
@Slf4j
public class ReportGeneratorServiceImpl implements ReportGeneratorService {

    private final DocumentFormatConverter formatConverter;

    /**
     * Creates io.github.ogbozoyan.service with default LibreOffice-based converter.
     */
    public ReportGeneratorServiceImpl() {
        this(new LibreOfficeDocumentFormatConverter());
        log.debug("ReportGeneratorServiceImpl() - end: converter={}", this.formatConverter.getClass().getSimpleName());
    }

    /**
     * Creates io.github.ogbozoyan.service with explicit output converter.
     *
     * @param formatConverter converter for post-processing output formats
     */
    public ReportGeneratorServiceImpl(DocumentFormatConverter formatConverter) {
        log.debug("ReportGeneratorServiceImpl(DocumentFormatConverter) - start: converterClass={}",
            formatConverter == null ? null : formatConverter.getClass().getName());
        this.formatConverter = Objects.requireNonNull(formatConverter, "formatConverter must not be null");
        log.debug("ReportGeneratorServiceImpl(DocumentFormatConverter) - end: converterClass={}",
            this.formatConverter.getClass().getName());
    }

    /**
     * Generates final report bytes from template and token data.
     *
     * <p>When {@code options} is {@code null}, {@link GenerateOptions#defaults()} is used.
     *
     * @param template template descriptor with source bytes
     * @param data     unified token map
     * @param options  generation options, nullable
     * @return generated report with output metadata and warnings
     * @throws IllegalArgumentException           when template or data is {@code null}
     * @throws UnsupportedTemplateFormatException when input/output format route is not supported
     */
    @Override
    public GeneratedReport generate(TemplateInput template, ReportData data, GenerateOptions options) {
        log.debug("generate() - start: fileName={}, contentType={}, bytesLength={}, scalarTokens={}, requestedOptionsPresent={}",
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
        log.debug("generate() - formatResolved: sourceFormat={}, requestedOutputFormat={}, outputFormat={}",
            sourceFormat, requestedOutputFormat, outputFormat);
        log.debug("generate() - optionsResolved: missingValuePolicy={}, recalculateFormulas={}, rowsOnlyTableTokens={}, zoneId={}",
            resolvedOptions.missingValuePolicy(),
            resolvedOptions.recalculateFormulas(),
            resolvedOptions.rowsOnlyTableTokens(),
            resolvedOptions.zoneId());

        WarningCollector warningCollector = new WarningCollector();

        try (WorkbookProcessor processor = createProcessor(sourceFormat, template.bytes())) {
            processor.applyTemplateTokens(data.templateTokens(), resolvedOptions, warningCollector);
            processor.recalculateFormulas(resolvedOptions);

            byte[] generatedBytes = processor.serialize();
            byte[] outputBytes = outputFormat == sourceFormat
                ? generatedBytes
                : formatConverter.convert(generatedBytes, sourceFormat, outputFormat);

            GeneratedReport report = ReportSerializer.serialize(template, outputFormat, outputBytes, warningCollector);
            log.debug("generate() - end: outputFileName={}, outputContentType={}, outputBytesLength={}, warnings={}",
                report.fileName(), report.contentType(), report.bytes().length, report.warnings().size());
            return report;
        }
    }

    /**
     * Creates format-specific io.github.ogbozoyan.processor for input template format.
     *
     * @param format detected source format
     * @param bytes  source template bytes
     * @return io.github.ogbozoyan.processor implementation for the format
     */
    private WorkbookProcessor createProcessor(TemplateFormat format, byte[] bytes) {
        if (bytes == null || bytes.length == 0) {
            throw new IllegalArgumentException("Template bytes must not be null or empty");
        }
        log.debug("createProcessor() - start: format={}, bytesLength={}", format, bytes.length);
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
        log.debug("createProcessor() - end: processorClass={}", processor.getClass().getName());
        return processor;
    }

    /**
     * Validates that source format can be used as template input.
     *
     * @param sourceFormat detected source template format
     */
    private void validateSourceTemplateFormat(TemplateFormat sourceFormat) {
        log.trace("validateSourceTemplateFormat() - start: sourceFormat={}", sourceFormat);
        if (sourceFormat == TemplateFormat.ODS || sourceFormat == TemplateFormat.ODT) {
            throw new UnsupportedTemplateFormatException(
                "ODS/ODT templates are not supported as input. " +
                "Use XLS/XLSX or DOC/DOCX templates and request ODS/ODT on output."
            );
        }
        log.trace("validateSourceTemplateFormat() - end: sourceFormat={}", sourceFormat);
    }

    /**
     * Validates requested output conversion route.
     *
     * <p>Allowed routes:
     * <ul>
     *     <li>same format (no conversion),</li>
     *     <li>{@code XLS/XLSX -> ODS},</li>
     *     <li>{@code DOC/DOCX -> ODT}.</li>
     * </ul>
     *
     * @param sourceFormat source processing format
     * @param outputFormat requested output format
     */
    private void validateOutputFormatConversion(TemplateFormat sourceFormat, TemplateFormat outputFormat) {
        log.trace("validateOutputFormatConversion() - start: sourceFormat={}, outputFormat={}", sourceFormat, outputFormat);
        if (sourceFormat == outputFormat) {
            log.trace("validateOutputFormatConversion() - end: sameFormat=true");
            return;
        }

        boolean spreadsheetToOds = isSpreadsheet(sourceFormat) && outputFormat == TemplateFormat.ODS;
        boolean wordToOdt = isWord(sourceFormat) && outputFormat == TemplateFormat.ODT;

        if (spreadsheetToOds || wordToOdt) {
            log.trace("validateOutputFormatConversion() - end: conversionAllowed=true");
            return;
        }

        throw new UnsupportedTemplateFormatException(
            "Unsupported output conversion: " + sourceFormat + " -> " + outputFormat + ". " +
            "Supported conversions: XLS/XLSX -> ODS and DOC/DOCX -> ODT."
        );
    }

    /**
     * Checks whether format belongs to spreadsheet family.
     *
     * @param format format to inspect
     * @return {@code true} for {@code XLS}/{@code XLSX}
     */
    private boolean isSpreadsheet(TemplateFormat format) {
        return format == TemplateFormat.XLS || format == TemplateFormat.XLSX;
    }

    /**
     * Checks whether format belongs to word-processing family.
     *
     * @param format format to inspect
     * @return {@code true} for {@code DOC}/{@code DOCX}
     */
    private boolean isWord(TemplateFormat format) {
        return format == TemplateFormat.DOC || format == TemplateFormat.DOCX;
    }
}
