package com.template.reportgenerator.util;

import com.template.reportgenerator.dto.TemplateFormat;
import com.template.reportgenerator.dto.TemplateInput;
import com.template.reportgenerator.exception.UnsupportedTemplateFormatException;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static com.template.reportgenerator.util.TemplateFormatDetector.detect;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class TemplateFormatDetectorTest {

    @Test
    void shouldDetectByExtension() {
        assertEquals(TemplateFormat.XLSX, detect(new TemplateInput("a.xlsx", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.XLS, detect(new TemplateInput("a.xls", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.ODS, detect(new TemplateInput("a.ods", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.DOC, detect(new TemplateInput("a.doc", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.DOCX, detect(new TemplateInput("a.docx", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.ODT, detect(new TemplateInput("a.odt", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.PDF, detect(new TemplateInput("a.pdf", null, new byte[] {1, 2, 3, 4})));
    }

    @Test
    void shouldDetectDocByContentType() {
        TemplateInput input = new TemplateInput(null, "application/msword", new byte[] {1, 2, 3, 4});
        assertEquals(TemplateFormat.DOC, detect(input));
    }

    @Test
    void shouldDetectByMagicBytes() {
        byte[] xls = {(byte) 0xD0, (byte) 0xCF, 0x11, (byte) 0xE0};
        assertEquals(TemplateFormat.XLS, detect(new TemplateInput(null, null, xls)));

        byte[] xlsxZip = zipWithEntries("xl/workbook.xml");
        assertEquals(TemplateFormat.XLSX, detect(new TemplateInput(null, null, xlsxZip)));

        byte[] pdf = {0x25, 0x50, 0x44, 0x46};
        assertEquals(TemplateFormat.PDF, detect(new TemplateInput(null, null, pdf)));
    }

    @Test
    void shouldThrowOnUnknown() {
        TemplateInput input = new TemplateInput("report.abc", "text/plain", new byte[] {1, 2, 3, 4});
        assertThrows(UnsupportedTemplateFormatException.class, () -> detect(input));
    }

    private byte[] zipWithEntries(String... entryNames) {
        try (ByteArrayOutputStream output = new ByteArrayOutputStream();
             ZipOutputStream zipOutputStream = new ZipOutputStream(output, StandardCharsets.UTF_8)) {
            for (String entryName : entryNames) {
                zipOutputStream.putNextEntry(new ZipEntry(entryName));
                zipOutputStream.write("<x/>".getBytes(StandardCharsets.UTF_8));
                zipOutputStream.closeEntry();
            }
            zipOutputStream.finish();
            return output.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
