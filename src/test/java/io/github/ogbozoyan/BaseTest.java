package io.github.ogbozoyan;

import io.github.ogbozoyan.data.GenerateOptions;
import io.github.ogbozoyan.data.MissingValuePolicy;
import io.github.ogbozoyan.data.TemplateFormat;
import io.github.ogbozoyan.service.ReportGeneratorService;
import io.github.ogbozoyan.service.ReportGeneratorServiceImpl;
import io.github.ogbozoyan.util.DocumentFormatConverter;
import lombok.Data;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.time.ZoneId;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertNotNull;

public class BaseTest {
    protected final ReportGeneratorService service = new ReportGeneratorServiceImpl();

    protected byte[] createXlsScalarTemplate() throws Exception {
        try (Workbook workbook = new HSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("S");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("{{name}}");

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            cell.setCellStyle(style);

            workbook.write(output);
            return output.toByteArray();
        }
    }

    protected byte[] createXlsxTableTemplate() throws Exception {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("S");
            Row markerRow = sheet.createRow(0);
            Cell markerCell = markerRow.createCell(0);
            markerCell.setCellValue("{{rows}}");

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            style.setWrapText(true);
            style.setAlignment(HorizontalAlignment.CENTER);
            markerCell.setCellStyle(style);

            sheet.setColumnWidth(0, 1200);
            sheet.setColumnWidth(1, 1200);
            sheet.createRow(1).createCell(0).setCellValue("after");

            workbook.write(output);
            return output.toByteArray();
        }
    }

    protected byte[] createXlsxInlineTableTemplate() throws Exception {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("S");
            Row markerRow = sheet.createRow(0);
            Cell markerCell = markerRow.createCell(0);
            markerCell.setCellValue("{{rows}} год");

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            style.setWrapText(true);
            style.setAlignment(HorizontalAlignment.CENTER);
            markerCell.setCellStyle(style);

            sheet.setColumnWidth(0, 1200);
            sheet.setColumnWidth(1, 1200);
            sheet.createRow(1).createCell(0).setCellValue("after");

            workbook.write(output);
            return output.toByteArray();
        }
    }

    protected byte[] createDocxTableTemplate() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText("{{rows}}");
            document.createParagraph().createRun().setText("tail");
            document.write(output);
            return output.toByteArray();
        }
    }

    protected byte[] createDocxNestedTableTokenTemplate() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            XWPFTable table = document.createTable(1, 2);
            putCellToken(table.getRow(0).getCell(0), "{{inner_table}}");
            putCellToken(table.getRow(0).getCell(1), "{{mega_test}}");
            document.write(output);
            return output.toByteArray();
        }
    }

    private void putCellToken(XWPFTableCell cell, String token) {
        for (int i = cell.getParagraphs().size() - 1; i >= 0; i--) {
            cell.removeParagraph(i);
        }
        cell.addParagraph().createRun().setText(token);
    }

    protected byte[] createDocxScalarTemplate(String value) throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText(value);
            document.write(output);
            return output.toByteArray();
        }
    }

    protected byte[] createPdfTableTemplate() throws Exception {
        try (PDDocument document = new PDDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            PDPage page = new PDPage(PDRectangle.A4);
            document.addPage(page);

            try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                contentStream.setFont(new PDType1Font(Standard14Fonts.FontName.HELVETICA), 12);
                contentStream.beginText();
                contentStream.newLineAtOffset(40, 780);
                contentStream.showText("{{rows}}");
                contentStream.endText();
            }

            document.save(output);
            return output.toByteArray();
        }
    }

    protected byte[] createSimpleXlsx(String value, boolean withStyleAndWidth) throws Exception {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("S");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(value);

            if (withStyleAndWidth) {
                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setBold(true);
                style.setFont(font);
                style.setWrapText(true);
                cell.setCellStyle(style);
                sheet.setColumnWidth(0, 4200);
            }

            workbook.write(output);
            return output.toByteArray();
        }
    }

    protected GenerateOptions rowsOnlyOptions() {
        return new GenerateOptions(
            MissingValuePolicy.EMPTY_AND_LOG,
            true,
            Locale.getDefault(), ZoneId.systemDefault(), true
        );
    }

    protected GenerateOptions headerModeOptions() {
        return new GenerateOptions(
            MissingValuePolicy.EMPTY_AND_LOG,
            true,
            Locale.getDefault(), ZoneId.systemDefault(), false
        );
    }

    protected byte[] loadResourceBytes(String path) throws Exception {
        try (InputStream stream = getClass().getResourceAsStream(path)) {
            assertNotNull(stream, "Missing test resource: " + path);
            return stream.readAllBytes();
        }
    }

    protected byte[] createXlsxFormattingTemplate() throws Exception {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("S");

            Row markerRow = sheet.createRow(0);
            markerRow.setHeight((short) 680);
            markerRow.createCell(1).setCellValue("{{rows}}");

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontHeightInPoints((short) 13);
            style.setFont(font);
            style.setWrapText(true);
            style.setAlignment(HorizontalAlignment.CENTER);
            markerRow.getCell(1).setCellStyle(style);
            sheet.setColumnWidth(1, 1700);

            Row staticRow = sheet.createRow(1);
            staticRow.createCell(1).setCellValue("static");
            staticRow.createCell(2).setCellValue("");
            staticRow.getCell(1).setCellStyle(style);
            staticRow.getCell(2).setCellStyle(style);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 2));

            workbook.write(output);
            return output.toByteArray();
        }
    }

    protected Map<String, Object> row(Object... values) {
        LinkedHashMap<String, Object> row = new LinkedHashMap<>();
        for (int i = 0; i < values.length; i += 2) {
            row.put(String.valueOf(values[i]), values[i + 1]);
        }
        return row;
    }

    protected boolean hasMergedRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.getFirstRow() == firstRow
                && region.getLastRow() == lastRow
                && region.getFirstColumn() == firstCol
                && region.getLastColumn() == lastCol) {
                return true;
            }
        }
        return false;
    }

    protected byte[] zipWithEntries(String... entryNames) {
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

    @Data
    public static final class RecordingConverter implements DocumentFormatConverter {
        private final byte[] convertedBytes;
        private int calls;
        private TemplateFormat sourceFormat;
        private TemplateFormat targetFormat;
        private byte[] sourceBytes;

        public RecordingConverter(byte[] convertedBytes) {
            this.convertedBytes = convertedBytes;
        }

        @Override
        public byte[] convert(byte[] sourceBytes, TemplateFormat sourceFormat, TemplateFormat targetFormat) {
            this.calls++;
            this.sourceBytes = sourceBytes;
            this.sourceFormat = sourceFormat;
            this.targetFormat = targetFormat;
            return convertedBytes;
        }
    }
}
