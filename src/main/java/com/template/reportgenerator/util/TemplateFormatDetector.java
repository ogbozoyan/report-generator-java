package com.template.reportgenerator.util;

import com.template.reportgenerator.dto.TemplateFormat;
import com.template.reportgenerator.dto.TemplateInput;
import com.template.reportgenerator.exception.UnsupportedTemplateFormatException;
import lombok.experimental.UtilityClass;

import java.util.Locale;

@UtilityClass
public class TemplateFormatDetector {

    public static TemplateFormat detect(TemplateInput input) {
        if (input.fileName() != null) {
            String name = input.fileName().toLowerCase(Locale.ROOT);
            if (name.endsWith(".xlsx")) {
                return TemplateFormat.XLSX;
            }
            if (name.endsWith(".xls")) {
                return TemplateFormat.XLS;
            }
            if (name.endsWith(".ods")) {
                return TemplateFormat.ODS;
            }
        }

        if (input.contentType() != null) {
            String contentType = input.contentType().toLowerCase(Locale.ROOT);
            if (contentType.contains("spreadsheetml")) {
                return TemplateFormat.XLSX;
            }
            if (contentType.contains("ms-excel")) {
                return TemplateFormat.XLS;
            }
            if (contentType.contains("oasis.opendocument.spreadsheet")) {
                return TemplateFormat.ODS;
            }
        }

        byte[] bytes = input.bytes();
        if (bytes.length >= 4) {
            int b0 = bytes[0] & 0xFF;
            int b1 = bytes[1] & 0xFF;
            int b2 = bytes[2] & 0xFF;
            int b3 = bytes[3] & 0xFF;

            // ZIP signature (xlsx / ods)
            if (b0 == 0x50 && b1 == 0x4B && b2 == 0x03 && b3 == 0x04) {
                String asText = new String(bytes, 0, Math.min(bytes.length, 512));
                if (asText.contains("mimetypeapplication/vnd.oasis.opendocument.spreadsheet")) {
                    return TemplateFormat.ODS;
                }
                return TemplateFormat.XLSX;
            }

            // OLE2 signature (xls)
            if (b0 == 0xD0 && b1 == 0xCF && b2 == 0x11 && b3 == 0xE0) {
                return TemplateFormat.XLS;
            }
        }

        throw new UnsupportedTemplateFormatException("Unsupported template format for file: " + input.fileName());
    }
}
