package io.github.ogbozoyan;

import io.github.ogbozoyan.contract.GeneratedReport;
import io.github.ogbozoyan.contract.ReportData;
import io.github.ogbozoyan.contract.TemplateFormat;
import io.github.ogbozoyan.contract.TemplateInput;
import io.github.ogbozoyan.exception.UnsupportedTemplateFormatException;
import io.github.ogbozoyan.service.ReportGeneratorService;
import io.github.ogbozoyan.service.ReportGeneratorServiceImpl;
import io.github.ogbozoyan.util.DocumentFormatConverter;
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
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
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
    void shouldInsertTableTokenFromInlineCellAndDropStaticText() throws Exception {
        byte[] template = createXlsxInlineTableTemplate();
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
            assertEquals("after", sheet.getRow(3).getCell(0).getStringCellValue());
        }
        assertTrue(result.warnings().stream().anyMatch(w -> "TABLE_TOKEN_INLINE_TEXT_DROPPED".equals(w.code())));
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
    void shouldApplyConfiguredTableColumnOrderForXlsx() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Map<String, Object>> rows = List.of(
            Map.of("name", "North", "amount", 1200.25, "жопа", "A", "слона", 1),
            Map.of("name", "South", "amount", 900.00, "жопа", "B", "слона", 2)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of(
                "rows", rows,
                "rows__columns", List.of("name", "amount", "жопа", "слона")
            )),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("amount", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("жопа", sheet.getRow(0).getCell(2).getStringCellValue());
            assertEquals("слона", sheet.getRow(0).getCell(3).getStringCellValue());
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
    void shouldInsertTableTokenInsideExistingDocxTableCell() throws Exception {
        byte[] template = createDocxNestedTableTokenTemplate();
        List<Map<String, Object>> innerRows = List.of(
            row("name", "North", "amount", 1200.25),
            row("name", "South", "amount", 900.00)
        );
        List<Map<String, Object>> megaRows = List.of(
            row("code", "M-1", "value", "OK"),
            row("code", "M-2", "value", "DONE")
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.docx", null, template),
            new ReportData(Map.of(
                "inner_table", innerRows,
                "mega_test", megaRows
            )),
            null
        );

        try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(result.bytes()))) {
            assertEquals(1, document.getTables().size());
            XWPFTable outerTable = document.getTables().get(0);

            XWPFTableCell innerTableCell = outerTable.getRow(0).getCell(0);
            assertEquals(1, innerTableCell.getTables().size());
            XWPFTable innerTable = innerTableCell.getTables().get(0);
            assertEquals("name", innerTable.getRow(0).getCell(0).getText());
            assertEquals("North", innerTable.getRow(1).getCell(0).getText());

            XWPFTableCell megaTableCell = outerTable.getRow(0).getCell(1);
            assertEquals(1, megaTableCell.getTables().size());
            XWPFTable megaTable = megaTableCell.getTables().get(0);
            assertEquals("code", megaTable.getRow(0).getCell(0).getText());
            assertEquals("M-1", megaTable.getRow(1).getCell(0).getText());
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
    void shouldConvertSpreadsheetOutputToOdsWhenRequested() throws Exception {
        byte[] template = createSimpleXlsx("{{name}}", false);
        RecordingConverter converter = new RecordingConverter("converted-ods".getBytes(StandardCharsets.UTF_8));
        ReportGeneratorService convertingService = new ReportGeneratorServiceImpl(converter);

        GeneratedReport result = convertingService.generate(
            new TemplateInput("report.ods", null, template),
            new ReportData(Map.of("name", "Alice")),
            null
        );

        assertEquals(1, converter.calls);
        assertEquals(TemplateFormat.XLSX, converter.sourceFormat);
        assertEquals(TemplateFormat.ODS, converter.targetFormat);
        assertEquals("report.ods", result.fileName());
        assertEquals(TemplateFormat.ODS.contentType(), result.contentType());
        assertArrayEquals("converted-ods".getBytes(StandardCharsets.UTF_8), result.bytes());

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(converter.sourceBytes))) {
            assertEquals("Alice", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldConvertWordOutputToOdtWhenRequestedByContentType() throws Exception {
        byte[] template = createDocxScalarTemplate("{{name}}");
        RecordingConverter converter = new RecordingConverter("converted-odt".getBytes(StandardCharsets.UTF_8));
        ReportGeneratorService convertingService = new ReportGeneratorServiceImpl(converter);

        GeneratedReport result = convertingService.generate(
            new TemplateInput("report.docx", TemplateFormat.ODT.contentType(), template),
            new ReportData(Map.of("name", "Nina")),
            null
        );

        assertEquals(1, converter.calls);
        assertEquals(TemplateFormat.DOCX, converter.sourceFormat);
        assertEquals(TemplateFormat.ODT, converter.targetFormat);
        assertEquals("report.odt", result.fileName());
        assertEquals(TemplateFormat.ODT.contentType(), result.contentType());
        assertArrayEquals("converted-odt".getBytes(StandardCharsets.UTF_8), result.bytes());

        try (XWPFDocument generatedDoc = new XWPFDocument(new ByteArrayInputStream(converter.sourceBytes))) {
            assertTrue(generatedDoc.getParagraphs().stream().anyMatch(p -> "Nina".equals(p.getText())));
        }
    }

    @Test
    void shouldNotInvokeConverterWhenOutputMatchesSource() throws Exception {
        byte[] template = createSimpleXlsx("{{name}}", false);
        RecordingConverter converter = new RecordingConverter("unused".getBytes(StandardCharsets.UTF_8));
        ReportGeneratorService convertingService = new ReportGeneratorServiceImpl(converter);

        GeneratedReport result = convertingService.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("name", "Alice")),
            null
        );

        assertEquals(0, converter.calls);
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            assertEquals("Alice", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldRejectOdsSourceTemplates() {
        UnsupportedTemplateFormatException exception = assertThrows(
            UnsupportedTemplateFormatException.class,
            () -> service.generate(
                new TemplateInput("report.ods", null, new byte[] {1, 2, 3, 4}),
                new ReportData(Map.of("name", "Alice")),
                null
            )
        );
        assertTrue(exception.getMessage().contains("ODS/ODT templates are not supported as input"));
    }

    @Test
    void shouldRejectOdtSourceTemplates() {
        UnsupportedTemplateFormatException exception = assertThrows(
            UnsupportedTemplateFormatException.class,
            () -> service.generate(
                new TemplateInput("report.odt", null, new byte[] {1, 2, 3, 4}),
                new ReportData(Map.of("name", "Alice")),
                null
            )
        );
        assertTrue(exception.getMessage().contains("ODS/ODT templates are not supported as input"));
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

    private byte[] createXlsxInlineTableTemplate() throws Exception {
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

    private byte[] createDocxTableTemplate() throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText("{{rows}}");
            document.createParagraph().createRun().setText("tail");
            document.write(output);
            return output.toByteArray();
        }
    }

    private byte[] createDocxNestedTableTokenTemplate() throws Exception {
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

    private byte[] createDocxScalarTemplate(String value) throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            document.createParagraph().createRun().setText(value);
            document.write(output);
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

    private static final class RecordingConverter implements DocumentFormatConverter {
        private final byte[] convertedBytes;
        private int calls;
        private TemplateFormat sourceFormat;
        private TemplateFormat targetFormat;
        private byte[] sourceBytes;

        private RecordingConverter(byte[] convertedBytes) {
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
