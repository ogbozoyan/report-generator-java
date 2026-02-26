package com.template.reportgenerator.util;

import com.template.reportgenerator.contract.GeneratedReport;
import com.template.reportgenerator.contract.TemplateFormat;
import com.template.reportgenerator.contract.TemplateInput;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

/**
 * Serializes generation result with normalized filename and content type.
 */
@UtilityClass
@Slf4j
public class ReportSerializer {

    public static GeneratedReport serialize(
        TemplateInput input,
        TemplateFormat format,
        byte[] bytes,
        WarningCollector warningCollector
    ) {
        log.info("serialize() - start: inputFileName={}, format={}, bytesLength={}, warnings={}",
            input == null ? null : input.fileName(),
            format,
            bytes == null ? null : bytes.length,
            warningCollector == null ? null : warningCollector.asList().size());
        String fileName = normalizeFileName(input.fileName(), format);
        String contentType = format.contentType();
        GeneratedReport report = new GeneratedReport(fileName, contentType, bytes, warningCollector.asList());
        log.info("serialize() - end: outputFileName={}, contentType={}, warnings={}",
            report.fileName(), report.contentType(), report.warnings().size());
        return report;
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
            return replaceExtension(fileName, format.extension());
        }

        return fileName + format.extension();
    }

    private static String replaceExtension(String fileName, String extension) {
        int dot = fileName.lastIndexOf('.');
        if (dot < 0) {
            return fileName + extension;
        }
        return fileName.substring(0, dot) + extension;
    }
}
