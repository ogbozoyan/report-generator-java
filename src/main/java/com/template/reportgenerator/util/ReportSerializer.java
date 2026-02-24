package com.template.reportgenerator.util;

import com.template.reportgenerator.dto.GeneratedReport;
import com.template.reportgenerator.dto.TemplateFormat;
import com.template.reportgenerator.dto.TemplateInput;
import lombok.experimental.UtilityClass;

/**
 * Serializes generation result with normalized filename and content type.
 */
@UtilityClass
public class ReportSerializer {

    public static GeneratedReport serialize(
        TemplateInput input,
        TemplateFormat format,
        byte[] bytes,
        WarningCollector warningCollector
    ) {
        String fileName = normalizeFileName(input.fileName(), format);
        String contentType = format.contentType();
        return new GeneratedReport(fileName, contentType, bytes, warningCollector.asList());
    }

    private static String normalizeFileName(String originalFileName, TemplateFormat format) {
        if (originalFileName == null || originalFileName.isBlank()) {
            return "generated" + format.extension();
        }

        String lower = originalFileName.toLowerCase();
        if (lower.endsWith(".xls")
            || lower.endsWith(".xlsx")
            || lower.endsWith(".ods")
            || lower.endsWith(".doc")
            || lower.endsWith(".docx")
            || lower.endsWith(".odt")
            || lower.endsWith(".pdf")) {
            return originalFileName;
        }

        return originalFileName + format.extension();
    }
}
