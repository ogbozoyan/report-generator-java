package com.template.reportgenerator;

import com.template.reportgenerator.dto.GeneratedReport;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.TemplateInput;
import com.template.reportgenerator.exception.TemplateSyntaxException;
import com.template.reportgenerator.service.ReportGeneratorService;
import com.template.reportgenerator.service.ReportGeneratorServiceImpl;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReportGeneratorServiceImplTest {

    private final ReportGeneratorService service = new ReportGeneratorServiceImpl();

    @Test
    void shouldGenerateXlsxAndKeepCellStyleAndWidth() throws Exception {
        byte[] template = createXlsxScalarTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("report.xlsx", null, template),
            new ReportData(Map.of("name", "Alice"), null, null),
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
            new ReportData(Map.of("name", "Bob"), null, null),
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
            new ReportData(Map.of("name", "Carol"), null, null),
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
    void shouldExpandTableBlockInXlsx() throws Exception {
        byte[] template = createXlsxTableTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("table.xlsx", null, template),
            new ReportData(
                Map.of(),
                Map.of("rows", List.of(
                    Map.of("name", "A"),
                    Map.of("name", "B"),
                    Map.of("name", "C")
                )),
                Map.of()
            ),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("A", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("B", sheet.getRow(2).getCell(1).getStringCellValue());
            assertEquals("C", sheet.getRow(3).getCell(1).getStringCellValue());

            Font f1 = workbook.getFontAt(sheet.getRow(1).getCell(1).getCellStyle().getFontIndexAsInt());
            Font f2 = workbook.getFontAt(sheet.getRow(2).getCell(1).getCellStyle().getFontIndexAsInt());
            assertEquals(f1.getBold(), f2.getBold());
        }
    }

    @Test
    void shouldExpandColumnBlockInXlsx() throws Exception {
        byte[] template = createXlsxColumnTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("col.xlsx", null, template),
            new ReportData(
                Map.of(),
                Map.of(),
                Map.of("cols", List.of(
                    Map.of("name", "X"),
                    Map.of("name", "Y"),
                    Map.of("name", "Z")
                ))
            ),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("X", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("Y", sheet.getRow(1).getCell(2).getStringCellValue());
            assertEquals("Z", sheet.getRow(1).getCell(3).getStringCellValue());
        }
    }

    @Test
    void shouldReplaceMissingTokenWithEmptyAndWarningByDefault() throws Exception {
        byte[] template = createSimpleXlsx("{{missing}}", true);

        GeneratedReport result = service.generate(
            new TemplateInput("missing.xlsx", null, template),
            new ReportData(Map.of(), null, null),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            assertEquals("", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
        assertTrue(result.warnings().stream().anyMatch(w -> "MISSING_TOKEN".equals(w.code())));
    }

    @Test
    void shouldFailOnUnpairedMarkers() throws Exception {
        byte[] template = createSimpleXlsx("[[TABLE_START:rows]]", false);

        assertThrows(
            TemplateSyntaxException.class,
            () -> service.generate(new TemplateInput("invalid.xlsx", null, template), new ReportData(null, null, null), null)
        );
    }

    private byte[] createXlsxScalarTemplate() throws Exception {
        return createSimpleXlsx("{{name}}", true);
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

            sheet.createRow(0).createCell(0).setCellValue("[[TABLE_START:rows]]");
            Row templateRow = sheet.createRow(1);
            Cell templCell = templateRow.createCell(1);
            templCell.setCellValue("{{item.name}}");
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            templCell.setCellStyle(style);

            sheet.createRow(2).createCell(2).setCellValue("[[TABLE_END:rows]]");

            workbook.write(output);
            return output.toByteArray();
        }
    }

    private byte[] createXlsxColumnTemplate() throws Exception {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("S");

            sheet.createRow(0).createCell(0).setCellValue("[[COL_START:cols]]");
            Row templateRow = sheet.createRow(1);
            templateRow.createCell(1).setCellValue("{{item.name}}");
            sheet.createRow(2).createCell(2).setCellValue("[[COL_END:cols]]");

            workbook.write(output);
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
}
