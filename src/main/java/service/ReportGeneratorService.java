package service;

import contract.GenerateOptions;
import contract.GeneratedReport;
import contract.ReportData;
import contract.TemplateInput;

/**
 * Public service contract for template-based report generation.
 */
public interface ReportGeneratorService {

    /**
     * Generates report bytes from a template and input data.
     *
     * @param template template file metadata and bytes
     * @param data     token data model; table tokens are provided as
     *                 {@code List<Map<String, Object>>} values in {@link ReportData#templateTokens()}
     * @param options  generation options; if {@code null}, defaults are used
     * @return generated report
     */
    GeneratedReport generate(TemplateInput template, ReportData data, GenerateOptions options);
}
