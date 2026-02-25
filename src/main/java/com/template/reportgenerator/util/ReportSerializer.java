package com.template.reportgenerator.util;

import com.template.reportgenerator.contract.GeneratedReport;
import com.template.reportgenerator.contract.TemplateFormat;
import com.template.reportgenerator.contract.TemplateInput;
import lombok.experimental.UtilityClass;
import org.apache.commons.lang3.StringUtils;

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

    private static String normalizeFileName(String fileName, TemplateFormat format) {
        if (StringUtils.isEmpty(fileName)) {
            return "generated" + format.extension();
        }

        String lower = fileName.toLowerCase();
        if (lower.endsWith(".xls")
            || lower.endsWith(".xlsx")
            || lower.endsWith(".ods")
            || lower.endsWith(".doc")
            || lower.endsWith(".docx")
            || lower.endsWith(".odt")
            || lower.endsWith(".pdf")) {
            return fileName;
        }

        return fileName + format.extension();
    }
}
