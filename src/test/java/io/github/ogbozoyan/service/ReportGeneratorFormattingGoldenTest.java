package io.github.ogbozoyan.service;

import io.github.ogbozoyan.BaseTest;
import io.github.ogbozoyan.contract.GeneratedReport;
import io.github.ogbozoyan.contract.ReportData;
import io.github.ogbozoyan.contract.TemplateInput;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReportGeneratorFormattingGoldenTest extends BaseTest {

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

}
