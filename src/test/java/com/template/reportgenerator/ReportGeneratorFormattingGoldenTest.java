package com.template.reportgenerator;

import com.template.reportgenerator.dto.GeneratedReport;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.TemplateInput;
import com.template.reportgenerator.service.ReportGeneratorService;
import com.template.reportgenerator.service.ReportGeneratorServiceImpl;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReportGeneratorFormattingGoldenTest {

    private final ReportGeneratorService service = new ReportGeneratorServiceImpl();

    @Test
    void shouldPreserveFormattingAndMergedRegionsForExpandedTableInXlsx() throws Exception {
        byte[] template = createXlsxTableFormattingTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("table-format.xlsx", null, template),
            new ReportData(
                Map.of(),
                Map.of("rows", List.of(
                    Map.of("name", "Alice"),
                    Map.of("name", "Bob")
                )),
                Map.of()
            ),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("Alice", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("Bob", sheet.getRow(2).getCell(1).getStringCellValue());

            assertEquals(5555, sheet.getColumnWidth(1));
            assertEquals(sheet.getRow(1).getHeight(), sheet.getRow(2).getHeight());

            Font templateFont = workbook.getFontAt(sheet.getRow(1).getCell(1).getCellStyle().getFontIndexAsInt());
            Font clonedFont = workbook.getFontAt(sheet.getRow(2).getCell(1).getCellStyle().getFontIndexAsInt());

            assertTrue(templateFont.getBold());
            assertEquals(templateFont.getBold(), clonedFont.getBold());
            assertEquals(templateFont.getFontHeightInPoints(), clonedFont.getFontHeightInPoints());
            assertEquals(sheet.getRow(1).getCell(1).getCellStyle().getWrapText(), sheet.getRow(2).getCell(1).getCellStyle().getWrapText());
            assertEquals(HorizontalAlignment.CENTER, sheet.getRow(2).getCell(1).getCellStyle().getAlignment());

            assertTrue(hasMergedRegion(sheet, 1, 1, 1, 2));
            assertTrue(hasMergedRegion(sheet, 2, 2, 1, 2));
        }
    }

    @Test
    void shouldPreserveFormattingAndMergedRegionsForExpandedColumnsInXlsx() throws Exception {
        byte[] template = createXlsxColumnFormattingTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("col-format.xlsx", null, template),
            new ReportData(
                Map.of(),
                Map.of(),
                Map.of("cols", List.of(
                    Map.of("name", "Q1"),
                    Map.of("name", "Q2")
                ))
            ),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("Q1", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("Q2", sheet.getRow(1).getCell(2).getStringCellValue());

            assertEquals(4400, sheet.getColumnWidth(1));
            assertEquals(4400, sheet.getColumnWidth(2));

            Font sourceFont = workbook.getFontAt(sheet.getRow(1).getCell(1).getCellStyle().getFontIndexAsInt());
            Font clonedFont = workbook.getFontAt(sheet.getRow(1).getCell(2).getCellStyle().getFontIndexAsInt());
            assertEquals(sourceFont.getItalic(), clonedFont.getItalic());
            assertEquals(sourceFont.getFontHeightInPoints(), clonedFont.getFontHeightInPoints());

            assertTrue(hasMergedRegion(sheet, 1, 2, 1, 1));
            assertTrue(hasMergedRegion(sheet, 1, 2, 2, 2));
        }
    }

    @Test
    void shouldPreserveStylesAndWidthsForExpandedTableInOds() throws Exception {
        byte[] template = createOdsTableFormattingTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("table-format.ods", null, template),
            new ReportData(
                Map.of(),
                Map.of("rows", List.of(
                    Map.of("name", "North"),
                    Map.of("name", "South")
                )),
                Map.of()
            ),
            null
        );

        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.loadDocument(new ByteArrayInputStream(result.bytes()))) {
            OdfTable table = document.getTableList(false).get(0);

            assertEquals("North", table.getCellByPosition(1, 1).getStringValue());
            assertEquals("South", table.getCellByPosition(1, 2).getStringValue());

            assertTrue(Math.abs(3100 - table.getColumnByIndex(1).getWidth()) <= 1);
            assertEquals(table.getRowByIndex(1).getHeight(), table.getRowByIndex(2).getHeight());

            assertEquals("center", table.getCellByPosition(1, 2).getHorizontalAlignment());
            assertTrue(table.getCellByPosition(1, 2).isTextWrapped());
        }
    }

    @Test
    void shouldPreserveStylesAndWidthsForExpandedColumnsInOds() throws Exception {
        byte[] template = createOdsColumnFormattingTemplate();

        GeneratedReport result = service.generate(
            new TemplateInput("col-format.ods", null, template),
            new ReportData(
                Map.of(),
                Map.of(),
                Map.of("cols", List.of(
                    Map.of("name", "Jan"),
                    Map.of("name", "Feb"),
                    Map.of("name", "Mar")
                ))
            ),
            null
        );

        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.loadDocument(new ByteArrayInputStream(result.bytes()))) {
            OdfTable table = document.getTableList(false).get(0);

            assertEquals("Jan", table.getCellByPosition(1, 1).getStringValue());
            assertEquals("Feb", table.getCellByPosition(2, 1).getStringValue());
            assertEquals("Mar", table.getCellByPosition(3, 1).getStringValue());

            assertTrue(Math.abs(2800 - table.getColumnByIndex(1).getWidth()) <= 1);
            assertTrue(Math.abs(2800 - table.getColumnByIndex(2).getWidth()) <= 1);
            assertTrue(Math.abs(2800 - table.getColumnByIndex(3).getWidth()) <= 1);

            assertTrue(
                "right".equals(table.getCellByPosition(2, 1).getHorizontalAlignment())
                    || "end".equals(table.getCellByPosition(2, 1).getHorizontalAlignment())
            );
            assertTrue(table.getCellByPosition(2, 1).isTextWrapped());
        }
    }

    private byte[] createXlsxTableFormattingTemplate() throws Exception {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("S");

            sheet.createRow(0).createCell(0).setCellValue("[[TABLE_START:rows]]");

            Row templateRow = sheet.createRow(1);
            templateRow.setHeight((short) 680);
            templateRow.createCell(1).setCellValue("{{item.name}}");

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontHeightInPoints((short) 13);
            style.setFont(font);
            style.setWrapText(true);
            style.setAlignment(HorizontalAlignment.CENTER);
            templateRow.getCell(1).setCellStyle(style);

            templateRow.createCell(2).setCellValue("");
            templateRow.getCell(2).setCellStyle(style);

            sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 2));
            sheet.setColumnWidth(1, 5555);

            sheet.createRow(2).createCell(3).setCellValue("[[TABLE_END:rows]]");

            workbook.write(output);
            return output.toByteArray();
        }
    }

    private byte[] createXlsxColumnFormattingTemplate() throws Exception {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("S");

            sheet.createRow(0).createCell(0).setCellValue("[[COL_START:cols]]");

            Row row1 = sheet.createRow(1);
            row1.createCell(1).setCellValue("{{item.name}}");
            Row row2 = sheet.createRow(2);
            row2.createCell(1).setCellValue("");

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setItalic(true);
            font.setFontHeightInPoints((short) 12);
            style.setFont(font);
            style.setWrapText(true);
            row1.getCell(1).setCellStyle(style);
            row2.getCell(1).setCellStyle(style);

            sheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
            sheet.setColumnWidth(1, 4400);

            sheet.createRow(3).createCell(2).setCellValue("[[COL_END:cols]]");

            workbook.write(output);
            return output.toByteArray();
        }
    }

    private byte[] createOdsTableFormattingTemplate() throws Exception {
        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.newSpreadsheetDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {

            OdfTable table = document.getTableList(false).get(0);
            table.getCellByPosition(0, 0).setStringValue("[[TABLE_START:rows]]");
            table.getCellByPosition(1, 1).setStringValue("{{item.name}}");
            table.getCellByPosition(1, 1).setHorizontalAlignment("center");
            table.getCellByPosition(1, 1).setTextWrapped(true);
            table.getRowByIndex(1).setHeight(1600, false);
            table.getColumnByIndex(1).setWidth(3100);
            table.getCellByPosition(3, 2).setStringValue("[[TABLE_END:rows]]");

            document.save(output);
            return output.toByteArray();
        }
    }

    private byte[] createOdsColumnFormattingTemplate() throws Exception {
        try (OdfSpreadsheetDocument document = OdfSpreadsheetDocument.newSpreadsheetDocument();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {

            OdfTable table = document.getTableList(false).get(0);
            table.getCellByPosition(0, 0).setStringValue("[[COL_START:cols]]");
            table.getCellByPosition(1, 1).setStringValue("{{item.name}}");
            table.getCellByPosition(1, 1).setHorizontalAlignment("right");
            table.getCellByPosition(1, 1).setTextWrapped(true);
            table.getColumnByIndex(1).setWidth(2800);
            table.getCellByPosition(2, 3).setStringValue("[[COL_END:cols]]");

            document.save(output);
            return output.toByteArray();
        }
    }

    private boolean hasMergedRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
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
}
