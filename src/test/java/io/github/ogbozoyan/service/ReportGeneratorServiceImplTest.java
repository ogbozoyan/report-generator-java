package io.github.ogbozoyan.service;

import io.github.ogbozoyan.BaseTest;
import io.github.ogbozoyan.data.GeneratedReport;
import io.github.ogbozoyan.data.ReportData;
import io.github.ogbozoyan.data.TagConstants;
import io.github.ogbozoyan.data.TemplateFormat;
import io.github.ogbozoyan.data.TemplateInput;
import io.github.ogbozoyan.exception.UnsupportedTemplateFormatException;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReportGeneratorServiceImplTest extends BaseTest {

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
    void shouldInsertRowsOnlyTableTokenInXlsxWithoutHeader() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Object[]> rows = List.of(
            new Object[] {"North", 1200.25},
            new Object[] {"South", 900.00}
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            rowsOnlyOptions()
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("North", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(1200.25, sheet.getRow(0).getCell(1).getNumericCellValue(), 0.0001);
            assertEquals("South", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(900.00, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.0001);
            assertEquals("after", sheet.getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldApplyConfiguredColumnOrderForRowsOnlyTableTokenInXlsx() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Object[]> rows = List.of(
            new Object[] {"RU", "North", 1200.25},
            new Object[] {"EU", "South", 900.00}
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of(
                "rows", rows,
                TagConstants.ROWS_COLUMNS.getValue(), List.of("region", "name")
            )),
            rowsOnlyOptions()
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("RU", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("North", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(1200.25, sheet.getRow(0).getCell(2).getNumericCellValue(), 0.0001);
        }
    }

    @Test
    void shouldUseArrayOrderForRowsOnlyTableTokenWithoutConfiguredColumns() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Object[]> rows = List.of(
            new Object[] {2, 1},
            new Object[] {3, 4, 5}
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            rowsOnlyOptions()
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(2.0, sheet.getRow(0).getCell(0).getNumericCellValue(), 0.0001);
            assertEquals(1.0, sheet.getRow(0).getCell(1).getNumericCellValue(), 0.0001);
            assertEquals(5.0, sheet.getRow(1).getCell(2).getNumericCellValue(), 0.0001);
        }
    }

    @Test
    void shouldShiftTailRowsForRowsOnlyTableTokenInXlsx() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Object[]> rows = List.of(
            new Object[] {"North", 1200.25},
            new Object[] {"South", 900.00},
            new Object[] {"West", 700.00}
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            rowsOnlyOptions()
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("after", sheet.getRow(3).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldResolveTableTokenInsertedByPreviousTablePassInRowsOnlyMode() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Object[]> outerRows = List.<Object[]>of(
            new Object[] {"{{rows2}}"}
        );
        List<Object[]> innerRows = List.of(
            new Object[] {"A"},
            new Object[] {"B"}
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of(
                "rows", outerRows,
                "rows2", innerRows
            )),
            rowsOnlyOptions()
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("A", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("B", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("after", sheet.getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldKeepHeaderModeWhenRowsOnlyFlagIsFalse() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "North", "amount", 1200.25),
            row("name", "South", "amount", 900.00)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            headerModeOptions()
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("North", sheet.getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldReturnEmptyTableWarningForRowsOnlyTableTokenInXlsx() throws Exception {
        byte[] template = createXlsxTableTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", List.of())),
            rowsOnlyOptions()
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(CellType.BLANK, sheet.getRow(0).getCell(0).getCellType());
            assertEquals("after", sheet.getRow(1).getCell(0).getStringCellValue());
        }
        assertTrue(result.warnings().stream().anyMatch(w -> "TABLE_TOKEN_EMPTY".equals(w.code())));
    }

    @Test
    void shouldReturnInvalidTableWarningWhenRowsOnlyPayloadUsesMapRows() throws Exception {
        byte[] template = createXlsxTableTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "North", "amount", 1200.25)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            rowsOnlyOptions()
        );

        assertTrue(result.warnings().stream().anyMatch(w -> "TABLE_TOKEN_INVALID".equals(w.code())));
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
                TagConstants.ROWS_COLUMNS.getValue(), List.of("name", "amount", "жопа", "слона")
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

        assertEquals(1, converter.getCalls());
        assertEquals(TemplateFormat.XLSX, converter.getSourceFormat());
        assertEquals(TemplateFormat.ODS, converter.getTargetFormat());
        assertEquals("report.ods", result.fileName());
        assertEquals(TemplateFormat.ODS.contentType(), result.contentType());
        assertArrayEquals("converted-ods".getBytes(StandardCharsets.UTF_8), result.bytes());

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(converter.getSourceBytes()))) {
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

        assertEquals(1, converter.getCalls());
        assertEquals(TemplateFormat.DOCX, converter.getSourceFormat());
        assertEquals(TemplateFormat.ODT, converter.getTargetFormat());
        assertEquals("report.odt", result.fileName());
        assertEquals(TemplateFormat.ODT.contentType(), result.contentType());
        assertArrayEquals("converted-odt".getBytes(StandardCharsets.UTF_8), result.bytes());

        try (XWPFDocument generatedDoc = new XWPFDocument(new ByteArrayInputStream(converter.getSourceBytes()))) {
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

        assertEquals(0, converter.getCalls());
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

}
