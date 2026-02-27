package io.github.ogbozoyan.util;

import io.github.ogbozoyan.contract.TemplateFormat;
import io.github.ogbozoyan.contract.TemplateInput;
import io.github.ogbozoyan.exception.UnsupportedTemplateFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static io.github.ogbozoyan.util.TemplateFormatDetector.detectFormat;
import static io.github.ogbozoyan.util.TemplateFormatDetector.detectRequestedOutputFormat;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

class TemplateFormatDetectorTest {

    @Test
    void shouldDetectFormatByExtension() {
        assertEquals(TemplateFormat.XLSX, detectFormat(new TemplateInput("a.xlsx", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.XLS, detectFormat(new TemplateInput("a.xls", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.ODS, detectFormat(new TemplateInput("a.ods", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.DOC, detectFormat(new TemplateInput("a.doc", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.DOCX, detectFormat(new TemplateInput("a.docx", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.ODT, detectFormat(new TemplateInput("a.odt", null, new byte[] {1, 2, 3, 4})));
        assertEquals(TemplateFormat.PDF, detectFormat(new TemplateInput("a.pdf", null, new byte[] {1, 2, 3, 4})));
    }

    @Test
    void shouldDetectFormatDocByContentType() {
        TemplateInput input = new TemplateInput(null, "application/msword", new byte[] {1, 2, 3, 4});
        assertEquals(TemplateFormat.DOC, detectFormat(input));
    }

    @Test
    void shouldDetectFormatByMagicBytes() {
        byte[] xls = {(byte) 0xD0, (byte) 0xCF, 0x11, (byte) 0xE0};
        assertEquals(TemplateFormat.XLS, detectFormat(new TemplateInput(null, null, xls)));

        byte[] xlsxZip = zipWithEntries("xl/workbook.xml");
        assertEquals(TemplateFormat.XLSX, detectFormat(new TemplateInput(null, null, xlsxZip)));

        byte[] pdf = {0x25, 0x50, 0x44, 0x46};
        assertEquals(TemplateFormat.PDF, detectFormat(new TemplateInput(null, null, pdf)));
    }

    @Test
    void shouldDetectOle2DocAndXlsByContainerEntries() throws Exception {
        byte[] docBytes = loadResourceBytes("/fixtures/doc-table-template.doc");
        assertEquals(TemplateFormat.DOC, detectFormat(new TemplateInput("unknown.bin", null, docBytes)));

        byte[] xlsBytes;
        try (Workbook workbook = new HSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            workbook.createSheet("S");
            workbook.write(output);
            xlsBytes = output.toByteArray();
        }
        assertEquals(TemplateFormat.XLS, detectFormat(new TemplateInput("unknown.bin", null, xlsBytes)));
    }

    @Test
    void shouldDetectRequestedOutputFormatWithoutReadingMagicBytes() {
        assertEquals(
            TemplateFormat.ODS,
            detectRequestedOutputFormat(new TemplateInput("report.ods", null, new byte[] {1, 2, 3, 4}))
        );
        assertEquals(
            TemplateFormat.ODT,
            detectRequestedOutputFormat(new TemplateInput("report.docx", TemplateFormat.ODT.contentType(), new byte[] {1, 2, 3, 4}))
        );
    }

    @Test
    void shouldThrowOnUnknown() {
        TemplateInput input = new TemplateInput("report.abc", "text/plain", new byte[] {1, 2, 3, 4});
        assertThrows(UnsupportedTemplateFormatException.class, () -> detectFormat(input));
    }

    private byte[] loadResourceBytes(String path) throws Exception {
        try (InputStream stream = getClass().getResourceAsStream(path)) {
            assertNotNull(stream, "Missing test resource: " + path);
            return stream.readAllBytes();
        }
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
