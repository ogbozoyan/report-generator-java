package io.github.ogbozoyan.service;

import io.github.ogbozoyan.data.GenerateOptions;
import io.github.ogbozoyan.data.GeneratedReport;
import io.github.ogbozoyan.data.ReportData;
import io.github.ogbozoyan.data.TemplateInput;

/**
 * Public io.github.ogbozoyan.service contract for template-based report generation.
 */
public interface ReportGeneratorService {

    /**
     * Generates report bytes from a template and input data.
     *
     * @param template template file metadata and bytes
     * @param data     token data model; table tokens are provided in {@link ReportData#templateTokens()} as
     *                 {@code List<Map<String, Object>>}, {@code List<Object[]>} (rows-only XLS/XLSX mode),
     *                 declarative {@code io.github.ogbozoyan.contract.TableBuilder} for DOC/DOCX,
     *                 or declarative {@code io.github.ogbozoyan.contract.TableXlsxBuilder} for XLS/XLSX
     * @param options  generation options; if {@code null}, defaults are used
     * @return generated report
     */
    GeneratedReport generate(TemplateInput template, ReportData data, GenerateOptions options);
}
