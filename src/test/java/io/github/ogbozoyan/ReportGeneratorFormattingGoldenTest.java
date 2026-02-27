package io.github.ogbozoyan;

import io.github.ogbozoyan.contract.GeneratedReport;
import io.github.ogbozoyan.contract.ReportData;
import io.github.ogbozoyan.contract.TemplateInput;
import io.github.ogbozoyan.service.ReportGeneratorService;
import io.github.ogbozoyan.service.ReportGeneratorServiceImpl;
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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReportGeneratorFormattingGoldenTest {

    private final ReportGeneratorService service = new ReportGeneratorServiceImpl();

    @Test
    void shouldPreserveStylesWidthsAndMergedRegionsForInsertedTableInXlsx() throws Exception {
        byte[] template = createXlsxFormattingTemplate();
        List<Map<String, Object>> rows = List.of(
            row("name", "Alice with long content", "amount", 100),
            row("name", "Bob", "amount", 200)
        );

        GeneratedReport result = service.generate(
            new TemplateInput("table-format.xlsx", null, template),
            new ReportData(Map.of("rows", rows)),
            null
        );

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(result.bytes()))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals("name", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("Alice with long content", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("Bob", sheet.getRow(2).getCell(1).getStringCellValue());

            CellStyle headerStyle = sheet.getRow(0).getCell(1).getCellStyle();
            CellStyle dataStyle = sheet.getRow(1).getCell(1).getCellStyle();
            Font headerFont = workbook.getFontAt(headerStyle.getFontIndexAsInt());
            Font dataFont = workbook.getFontAt(dataStyle.getFontIndexAsInt());

            assertTrue(headerFont.getBold());
            assertEquals(headerFont.getBold(), dataFont.getBold());
            assertEquals(headerStyle.getWrapText(), dataStyle.getWrapText());
            assertEquals(HorizontalAlignment.CENTER, dataStyle.getAlignment());
            assertEquals(sheet.getRow(0).getHeight(), sheet.getRow(1).getHeight());

            assertTrue(sheet.getColumnWidth(1) > 1700);
            assertEquals("static", sheet.getRow(3).getCell(1).getStringCellValue());
            assertTrue(hasMergedRegion(sheet, 3, 3, 1, 2));
        }
    }

    private byte[] createXlsxFormattingTemplate() throws Exception {
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

    private Map<String, Object> row(Object... values) {
        LinkedHashMap<String, Object> row = new LinkedHashMap<>();
        for (int i = 0; i < values.length; i += 2) {
            row.put(String.valueOf(values[i]), values[i + 1]);
        }
        return row;
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
