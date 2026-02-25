package com.template.reportgenerator;

import com.template.reportgenerator.contract.GeneratedReport;
import com.template.reportgenerator.contract.ReportData;
import com.template.reportgenerator.contract.TemplateInput;
import com.template.reportgenerator.service.ReportGeneratorService;
import com.template.reportgenerator.service.ReportGeneratorServiceImpl;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReportGeneratorServiceImplTest {

    private final ReportGeneratorService service = new ReportGeneratorServiceImpl();

    @Test
    void shouldGenerateXlsxAndKeepCellStyleAndWidth() throws Exception {
        byte[] template = createSimpleXlsx("{{name}}", true);

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("name", "Alice")),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            Cell outCell = sheet.getRow(0).getCell(0);

            assertEquals("Alice", outCell.getStringCellValue());
            assertEquals(4200, sheet.getColumnWidth(0));

            Font outFont = workbook.getFontAt(outCell.getCellStyle().getFontIndexAsInt());
            assertTrue(outFont.getBold());
            assertTrue(outCell.getCellStyle().getWrapText());
        }
    }

    @Test
    void shouldGenerateXls() throws Exception {
        byte[] template = createXlsScalarTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("report.xls", null, template),
            new ReportData(Map.of("name", "Bob")),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            assertEquals("Bob", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldGenerateOds() throws Exception {
        byte[] template = createOdsScalarTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("report.ods", null, template),
            new ReportData(Map.of("name", "Carol")),
            null
        );

        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.loadDocument(new ByteArrayInputStream(result.bytes()))) {
            OdfTable table = document.getTableList(false).get(0);
            assertEquals("Carol", table.getCellByPosition(0, 0).getStringValue());
            assertTrue(Math.abs(3000 - table.getColumnByIndex(0).getWidth()) <= 1);
            assertEquals("center", table.getCellByPosition(0, 0).getHorizontalAlignment());
        }
    }

    @Test
    void shouldInsertTableTokenInXlsx() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "North", "amount", 1200.25),
            row("name", "South", "amount", 900.00)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("amount", sheet.getRow(0).getCell(1).getStringCellValue());

            assertEquals("North", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(1200.25, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.0001);
            assertEquals("South", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("after", sheet.getRow(3).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldAppendNewColumnsFromFollowingRows() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Map<String, Object>> rows = new ArrayList<>();
        rows.add(row("name", "North", "amount", 1200.25));
        rows.add(row("name", "South", "amount", 900.00, "region", "RU"));

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("amount", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("region", sheet.getRow(0).getCell(2).getStringCellValue());
            assertEquals("RU", sheet.getRow(2).getCell(2).getStringCellValue());
        }
    }

    @Test
    void shouldAutoResizeColumnsForInsertedTableInXlsx() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "Very very long region name for width expansion", "amount", 1200.25)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertTrue(sheet.getColumnWidth(0) > 1200);
            assertTrue(sheet.getColumnWidth(1) >= 1200);
        }
    }

    @Test
    void shouldKeepMarkerStyleAsBaselineForInsertedTableInXlsx() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Map<String, Object>> rows = List.of(row("name", "North", "amount", 1200.25));

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            CellStyle headerStyle = sheet.getRow(0).getCell(0).getCellStyle();
            CellStyle dataStyle = sheet.getRow(1).getCell(0).getCellStyle();

            Font headerFont = workbook.getFontAt(headerStyle.getFontIndexAsInt());
            Font dataFont = workbook.getFontAt(dataStyle.getFontIndexAsInt());

            assertTrue(headerFont.getBold());
            assertEquals(headerFont.getBold(), dataFont.getBold());
            assertEquals(headerStyle.getWrapText(), dataStyle.getWrapText());
            assertEquals(HorizontalAlignment.CENTER, dataStyle.getAlignment());
        }
    }

    @Test
    void shouldInsertTableTokenInOds() throws Exception {
        byte[] template = createOdsTableTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "North", "amount", 1200.25),
            row("name", "South", "amount", 900.00)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.ods", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.loadDocument(new ByteArrayInputStream(result.bytes()))) {
            OdfTable table = document.getTableList(false).get(0);
            assertEquals("name", table.getCellByPosition(0, 0).getStringValue());
            assertEquals("amount", table.getCellByPosition(1, 0).getStringValue());
            assertEquals("North", table.getCellByPosition(0, 1).getStringValue());
            assertEquals("South", table.getCellByPosition(0, 2).getStringValue());
            assertEquals("after", table.getCellByPosition(0, 3).getStringValue());
            assertTrue(table.getColumnByIndex(0).getWidth() > 1200);
            assertEquals("center", table.getCellByPosition(0, 1).getHorizontalAlignment());
        }
    }

    @Test
    void shouldReplaceInlineYearTokenInRealOdsTemplateAtA6() throws Exception {
        byte[] template = loadResourceBytes("/fixtures/compensation-template.ods");

        GeneratedReport result = service.generate(
            new TemplateInput("report.ods", null, template),
            new ReportData(Map.of("year", "2026")),
            null
        );

        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.loadDocument(new ByteArrayInputStream(result.bytes()))) {
            OdfTable table = document.getTableList(false).get(0);
            String a6 = table.getCellByPosition(0, 5).getStringValue();
            assertEquals("Январь, 2026", a6);
            assertFalse(a6.contains("{{year}}"));
        }
    }

    @Test
    void shouldInsertTableTokenInDocx() throws Exception {
        byte[] template = createDocxTableTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "North", "amount", 1200.25),
            row("name", "South", "amount", 900.00)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.docx", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(result.bytes()))) {
            assertEquals(1, document.getTables().size());
            assertEquals("name", document.getTables().get(0).getRow(0).getCell(0).getText());
            assertEquals("North", document.getTables().get(0).getRow(1).getCell(0).getText());
            assertTrue(document.getParagraphs().stream().noneMatch(p -> "{{rows}}".equals(p.getText())));
        }
    }

    @Test
    void shouldInsertTableTokenInOdt() throws Exception {
        byte[] template = createOdtTableTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "North", "amount", 1200.25),
            row("name", "South", "amount", 900.00)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.odt", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (OdfTextDocument document = OdfTextDocument.loadDocument(new ByteArrayInputStream(result.bytes()))) {
            assertEquals(1, document.getTableList(false).size());
            OdfTable table = document.getTableList(false).get(0);
            assertEquals("name", table.getCellByPosition(0, 0).getStringValue());
            assertEquals("South", table.getCellByPosition(0, 2).getStringValue());
        }
    }

    @Test
    void shouldInsertTableTokenInPdfAsTextGrid() throws Exception {
        byte[] template = createPdfTableTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "North", "amount", 1200.25),
            row("name", "South", "amount", 900.00)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.pdf", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (PDDocument document = Loader.loadPDF(result.bytes())) {
            String text = new PDFTextStripper().getText(document);
            assertTrue(text.contains("name"));
            assertTrue(text.contains("amount"));
            assertTrue(text.contains("North"));
            assertTrue(text.contains("South"));
        }
    }

    @Test
    void shouldInsertTableTokenInDocAsBasicTextTable() throws Exception {
        byte[] template = loadResourceBytes("/fixtures/doc-table-template.doc");
        List<Map<String, Object>> rows = List.of(
            row("name", "North", "amount", 1200.25),
            row("name", "South", "amount", 900.00)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.doc", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (HWPFDocument document = new HWPFDocument(new ByteArrayInputStream(result.bytes()));
             WordExtractor extractor = new WordExtractor(document)) {
            String text = extractor.getText();
            assertTrue(text.contains("name\tamount"));
            assertTrue(text.contains("North\t1200.25"));
            assertTrue(text.contains("South\t900.0"));
        }
    }

    @Test
    void shouldReplaceMissingTokenWithEmptyAndWarningByDefault() throws Exception {
        byte[] template = createSimpleXlsx("{{missing}}", true);

        GeneratedReport result = service.generate(
            new TemplateInput("missing.xlsx", null, template),
            new ReportData(Map.of()),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            assertEquals("", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
        assertTrue(result.warnings().stream().anyMatch(w -> "MISSING_TOKEN".equals(w.code())));
    }

    private byte[] createXlsScalarTemplate() throws Exception {
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

    private byte[] createOdsScalarTemplate() throws Exception {
        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.newSpreadsheetDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {

            OdfTable table = document.getTableList(false).get(0);
            table.getCellByPosition(0, 0).setStringValue("{{name}}");
            table.getCellByPosition(0, 0).setHorizontalAlignment("center");
            table.getColumnByIndex(0).setWidth(3000);

            document.save(output);
            return output.toByteArray();
        }
    }

    private byte[] createXlsxTableTemplate() throws Exception {
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

    private byte[] createOdsTableTemplate() throws Exception {
        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.newSpreadsheetDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            OdfTable table = document.getTableList(false).get(0);
            table.getCellByPosition(0, 0).setStringValue("{{rows}}");
            table.getCellByPosition(0, 0).setHorizontalAlignment("center");
            table.getCellByPosition(0, 0).setTextWrapped(true);
            table.getColumnByIndex(0).setWidth(1200);
            table.getColumnByIndex(1).setWidth(1200);
            table.getRowByIndex(0).setHeight(1000, false);
            table.getCellByPosition(0, 1).setStringValue("after");
            document.save(output);
            return output.toByteArray();
        }
    }

    private byte[] createDocxTableTemplate() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText("{{rows}}");
            document.createParagraph().createRun().setText("tail");
            document.write(output);
            return output.toByteArray();
        }
    }

    private byte[] createOdtTableTemplate() throws Exception {
        try (OdfTextDocument document = OdfTextDocument.newTextDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.newParagraph("{{rows}}");
            document.newParagraph("tail");
            document.save(output);
            return output.toByteArray();
        }
    }

    private byte[] createPdfTableTemplate() throws Exception {
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

    private byte[] createSimpleXlsx(String value, boolean withStyleAndWidth) throws Exception {
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

    private byte[] loadResourceBytes(String path) throws Exception {
        try (InputStream stream = getClass().getResourceAsStream(path)) {
            assertNotNull(stream, "Missing test resource: " + path);
            return stream.readAllBytes();
        }
    }

    private Map<String, Object> row(Object... values) {
        LinkedHashMap<String, Object> row = new LinkedHashMap<>();
        for (int i = 0; i < values.length; i += 2) {
            row.put(String.valueOf(values[i]), values[i + 1]);
        }
        return row;
    }
}
