package com.template.reportgenerator.util;

import com.template.reportgenerator.dto.TemplateFormat;
import com.template.reportgenerator.dto.TemplateInput;
import com.template.reportgenerator.exception.UnsupportedTemplateFormatException;
import org.junit.jupiter.api.Test;

import static com.template.reportgenerator.util.TemplateFormatDetector.detect;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class TemplateFormatDetectorTest {

    @Test
    void shouldDetectByExtension() {
        assertEquals(TemplateFormat.XLSX, detect(new TemplateInput("a.xlsx", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.XLS, detect(new TemplateInput("a.xls", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.ODS, detect(new TemplateInput("a.ods", null, new byte[] {1, 2, 3, 4})));
    }

    @Test
    void shouldDetectByMagicBytes() {
        byte[] xls = {(byte) 0xD0, (byte) 0xCF, 0x11, (byte) 0xE0};
        assertEquals(TemplateFormat.XLS, detect(new TemplateInput(null, null, xls)));

        byte[] zip = {0x50, 0x4B, 0x03, 0x04};
        assertEquals(TemplateFormat.XLSX, detect(new TemplateInput(null, null, zip)));
    }

    @Test
    void shouldThrowOnUnknown() {
        TemplateInput input = new TemplateInput("report.abc", "text/plain", new byte[] {1, 2, 3, 4});
        assertThrows(UnsupportedTemplateFormatException.class, () -> detect(input));
    }
}
