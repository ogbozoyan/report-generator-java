package io.github.ogbozoyan.integration;

import io.github.ogbozoyan.contract.TableXlsxBuilder;
import io.github.ogbozoyan.data.GenerateOptions;
import io.github.ogbozoyan.data.GeneratedReport;
import io.github.ogbozoyan.data.GenerationWarning;
import io.github.ogbozoyan.data.MissingValuePolicy;
import io.github.ogbozoyan.data.ReportData;
import io.github.ogbozoyan.data.TemplateInput;
import io.github.ogbozoyan.service.ReportGeneratorService;
import io.github.ogbozoyan.service.ReportGeneratorServiceImpl;
import lombok.NonNull;
import lombok.SneakyThrows;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.ZoneId;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

@Disabled("Manual integration scenarios for local templates")
class ReportGeneratorManualIntegrationTest {

    /**
     * Generates DOCX sample report used for local smoke testing.
     *
     * @param service report generator io.github.ogbozoyan.service
     * @return generated report
     * @throws IOException when template/result file I/O fails
     */
    private static GeneratedReport docx(ReportGeneratorService service) throws IOException {
        String fileName = "DOC1.docx";
        String resultFileName = "DOC1_result.docx";
        byte[] templateBytes = Files.readAllBytes(Path.of(fileName));
        TemplateInput input = new TemplateInput(
            fileName,
            null,
            templateBytes
        );
        Map<String, Object> tagsMap = new HashMap<>();
        tagsMap.put("SOME_PLACEHOLDER", "Some value");
        tagsMap.put("mega_test", 234);
        tagsMap.put(
            "TABLE_HERE", List.of(
                Map.of("name", "North", "amount", 1200.25),
                Map.of("name", "South", "amount", 900.00),
                Map.of("жопа", "South", "слона", 900.00)
            )
        );

        tagsMap.put(
            "inner_table", List.of(
                Map.of("name", "North", "amount", 1200.25),
                Map.of("name", "South", "amount", 900.00),
                Map.of("жопа", "South", "слона", 900.00)
            )
        );
        tagsMap.put(
            "ANOTHER_TABLE", List.of(
                Map.of("name", "North", "amount", 1200.25),
                Map.of("name", "South", "amount", 900.00),
                Map.of("жопа", "South", "слона", 900.00)
            )
        );

        ReportData data = new ReportData(tagsMap);
        GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.EMPTY_AND_LOG,
            true,
            Locale.forLanguageTag("ru-RU"), ZoneId.of("Europe/Moscow"),
            false
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("out/" + resultFileName), report.bytes());
        return report;
    }

    /**
     * Generates XLSX sample report used for local smoke testing.
     *
     * @param service report generator io.github.ogbozoyan.service
     * @return generated report
     * @throws IOException when template/result file I/O fails
     */
    private static @NonNull GeneratedReport xlsx(ReportGeneratorService service) throws IOException {
        byte[] templateBytes = Files.readAllBytes(Path.of("TABLE_BOOK.xlsx"));

        TemplateInput input = new TemplateInput(
            "TABLE_BOOK.xlsx",
            null,
            templateBytes
        );

        Map<String, Object> tagsMap = new HashMap<>();
        tagsMap.put("Table_2", List.of(
            Map.of("name", "North", "amount", 1200.25),
            Map.of("name", "South", "amount", 900.00),
            Map.of("жопа", "South", "слона", 900.00)
        ));
        tagsMap.put("Table_2__columns", List.of("name", "amount", "жопа", "слона"));
        tagsMap.put("TABLE_1", List.of(
                Map.of("name", "North", "amount", 1200.25),
                Map.of("name", "South", "amount", 900.00),
                Map.of("жопа", "South", "слона", 900.00)
            )
        )
        ;
        tagsMap.put("TABLE_1__columns", List.of("name", "amount", "жопа", "слона"));
        tagsMap.put(
            "year", "1999"
        );
        tagsMap.put("TAG_1", "TAG_1_VALUE");
        tagsMap.put("TAG_2", "TAG_2_VALUE");
        tagsMap.put("TAG_3", "TAG_3_VALUE");
        tagsMap.put("TAG_4", "TAG_4_VALUE");
        tagsMap.put("TAG_5", "TAG_5_VALUE");
        tagsMap.put("TAG_6", "TAG_6_VALUE");
        tagsMap.put("TAG_7", "TAG_7_VALUE");
        tagsMap.put("TAG_8", "TAG_8_VALUE");
        tagsMap.put("TAG_9", "TAG_9_VALUE");
        tagsMap.put("TAG_11", "TAG_11_VALUE");
        tagsMap.put("TAG_12", "TAG_12_VALUE");

        ReportData data = new ReportData(tagsMap);

        GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.EMPTY_AND_LOG,
            true, // recalculate formulas (для XLS/XLSX)
            Locale.forLanguageTag("ru-RU"), ZoneId.of("Europe/Moscow"),
            false
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("out/book_table.xlsx"), report.bytes());
        return report;
    }

    /**
     * Generates XLSX sample report used for local smoke testing.
     *
     * @param service report generator io.github.ogbozoyan.service
     * @return generated report
     * @throws IOException when template/result file I/O fails
     */
    private static @NonNull GeneratedReport xlsxRows(ReportGeneratorService service) throws IOException {
        byte[] templateBytes = Files.readAllBytes(Path.of("TABLE_BOOK_ROWS.xlsx"));

        TemplateInput input = new TemplateInput(
            "TABLE_BOOK_ROWS.xlsx",
            null,
            templateBytes
        );

        Map<String, Object> tagsMap = new HashMap<>();

        tagsMap.put("TABLE_WITH_ROWS_ONLY", List.of(
            new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            },
            new Object[] {
                "Петров Петр Петрович", "67890", "Аналитик", "1985-05-15",
                "39", "2019-03-15", "75000.00", "15000.00", "2019-04-15",
                "15000.00", "15000.00", "15000.00"
            },
            new Object[] {
                "Сидоров Сидор Сидорович", "54321", "Разработчик", "IT отдел", "1992-12-20",
                "32", "2021-06-01", "60000.00", "12000.00", "2021-07-01",
                "12000.00", "12000.00", "12000.00"
            }
        ));


        ReportData data = new ReportData(tagsMap);

        GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.EMPTY_AND_LOG,
            true, // recalculate formulas (для XLS/XLSX)
            Locale.forLanguageTag("ru-RU"), ZoneId.of("Europe/Moscow"),
            true
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("out/book_table_rows.xlsx"), report.bytes());
        return report;
    }

    /**
     * Generates XLSX sample report used for local smoke testing.
     *
     * @param service report generator io.github.ogbozoyan.service
     * @return generated report
     * @throws IOException when template/result file I/O fails
     */
    private static @NonNull GeneratedReport xlsxRowsAndTemplateBellow(ReportGeneratorService service) throws IOException {
        byte[] templateBytes = Files.readAllBytes(Path.of("table_with_row_and_template.xlsx"));

        TemplateInput input = new TemplateInput(
            "table_with_row_and_template.xlsx",
            null,
            templateBytes
        );

        Map<String, Object> tagsMap = new HashMap<>();

        tagsMap.put(
            "TABLE_ROW",
            TableXlsxBuilder.create()
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Петров Петр Петрович"),
                    TableXlsxBuilder.cell("67890"),
                    TableXlsxBuilder.cell("Аналитик"),
                    TableXlsxBuilder.cell(""),
                    TableXlsxBuilder.cell("1985-05-15"),
                    TableXlsxBuilder.cell("39"),
                    TableXlsxBuilder.cell("2019-03-15"),
                    TableXlsxBuilder.cell("75000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("2019-04-15"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Сидоров Сидор Сидорович"),
                    TableXlsxBuilder.cell("54321"),
                    TableXlsxBuilder.cell("Разработчик"),
                    TableXlsxBuilder.cell("IT отдел"),
                    TableXlsxBuilder.cell("1992-12-20"),
                    TableXlsxBuilder.cell("32"),
                    TableXlsxBuilder.cell("2021-06-01"),
                    TableXlsxBuilder.cell("60000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("2021-07-01"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00")
                )
        );
        tagsMap.put("TABLE_ROW_2", List.of(
            new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            }, new Object[] {
                "Иванов Иван Иванович", "12345", "Менеджер", "Главный офис", "1990-01-01",
                "34", "2020-01-01", "50000.00", "10000.00", "2020-02-01",
                "10000.00", "10000.00", "10000.00"
            },
            new Object[] {
                "Петров Петр Петрович", "67890", "Аналитик", "1985-05-15",
                "39", "2019-03-15", "75000.00", "15000.00", "2019-04-15",
                "15000.00", "15000.00", "15000.00"
            },
            new Object[] {
                "Сидоров Сидор Сидорович", "54321", "Разработчик", "IT отдел", "1992-12-20",
                "32", "2021-06-01", "60000.00", "12000.00", "2021-07-01",
                "12000.00", "12000.00", "12000.00"
            }
        ));
        tagsMap.put("RESULT", "ВЫВОДЫ ЫЫЫ");
        tagsMap.put("TOKEN_HEAD", "ЗЕАД ЫЫЫ");
        tagsMap.put("RESULT_VALUE", "ЖОПА ЫЫЫ");


        ReportData data = new ReportData(tagsMap);

        GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.EMPTY_AND_LOG,
            true, // recalculate formulas (для XLS/XLSX)
            Locale.forLanguageTag("ru-RU"), ZoneId.of("Europe/Moscow"),
            true
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("out/table_with_row_and_template_result.xlsx"), report.bytes());
        return report;
    }

    @SneakyThrows
    private static GeneratedReport xlsxRowsAndTemplateBellowDifficult(ReportGeneratorService service) {
        byte[] templateBytes = Files.readAllBytes(Path.of("table_with_row_and_template_difficult.xlsx"));

        TemplateInput input = new TemplateInput(
            "table_with_row_and_template_difficult.xlsx",
            null,
            templateBytes
        );

        Map<String, Object> tagsMap = new HashMap<>();

        tagsMap.put(
            "TABLE_PART_1",
            TableXlsxBuilder.create()
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Петров Петр Петрович"),
                    TableXlsxBuilder.cell("67890"),
                    TableXlsxBuilder.cell("Аналитик"),
                    TableXlsxBuilder.cell(""),
                    TableXlsxBuilder.cell("1985-05-15"),
                    TableXlsxBuilder.cell("39"),
                    TableXlsxBuilder.cell("2019-03-15"),
                    TableXlsxBuilder.cell("75000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("2019-04-15"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Сидоров Сидор Сидорович"),
                    TableXlsxBuilder.cell("54321"),
                    TableXlsxBuilder.cell("Разработчик"),
                    TableXlsxBuilder.cell("IT отдел"),
                    TableXlsxBuilder.cell("1992-12-20"),
                    TableXlsxBuilder.cell("32"),
                    TableXlsxBuilder.cell("2021-06-01"),
                    TableXlsxBuilder.cell("60000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("2021-07-01"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00")
                )
        );

        tagsMap.put(
            "TABLE_PART_2",
            TableXlsxBuilder.create()
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович", 2),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Петров Петр Петрович", 2),
                    TableXlsxBuilder.cell("67890"),
                    TableXlsxBuilder.cell("Аналитик"),
                    TableXlsxBuilder.cell(""),
                    TableXlsxBuilder.cell("1985-05-15"),
                    TableXlsxBuilder.cell("39"),
                    TableXlsxBuilder.cell("2019-03-15"),
                    TableXlsxBuilder.cell("75000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("2019-04-15"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Сидоров Сидор Сидорович", 2),
                    TableXlsxBuilder.cell("54321"),
                    TableXlsxBuilder.cell("Разработчик"),
                    TableXlsxBuilder.cell("IT отдел"),
                    TableXlsxBuilder.cell("1992-12-20"),
                    TableXlsxBuilder.cell("32"),
                    TableXlsxBuilder.cell("2021-06-01"),
                    TableXlsxBuilder.cell("60000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("2021-07-01"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00")
                )
        );

        tagsMap.put(
            "TABLE_PART_3",
            TableXlsxBuilder.create()
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Петров Петр Петрович"),
                    TableXlsxBuilder.cell("67890"),
                    TableXlsxBuilder.cell("Аналитик"),
                    TableXlsxBuilder.cell(""),
                    TableXlsxBuilder.cell("1985-05-15"),
                    TableXlsxBuilder.cell("39"),
                    TableXlsxBuilder.cell("2019-03-15"),
                    TableXlsxBuilder.cell("75000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("2019-04-15"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Сидоров Сидор Сидорович"),
                    TableXlsxBuilder.cell("54321"),
                    TableXlsxBuilder.cell("Разработчик"),
                    TableXlsxBuilder.cell("IT отдел"),
                    TableXlsxBuilder.cell("1992-12-20"),
                    TableXlsxBuilder.cell("32"),
                    TableXlsxBuilder.cell("2021-06-01"),
                    TableXlsxBuilder.cell("60000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("2021-07-01"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00")
                )
        );

        tagsMap.put(
            "TABLE_PART_4",
            TableXlsxBuilder.create()
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Петров Петр Петрович"),
                    TableXlsxBuilder.cell("67890"),
                    TableXlsxBuilder.cell("Аналитик"),
                    TableXlsxBuilder.cell(""),
                    TableXlsxBuilder.cell("1985-05-15"),
                    TableXlsxBuilder.cell("39"),
                    TableXlsxBuilder.cell("2019-03-15"),
                    TableXlsxBuilder.cell("75000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("2019-04-15"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Сидоров Сидор Сидорович"),
                    TableXlsxBuilder.cell("54321"),
                    TableXlsxBuilder.cell("Разработчик"),
                    TableXlsxBuilder.cell("IT отдел"),
                    TableXlsxBuilder.cell("1992-12-20"),
                    TableXlsxBuilder.cell("32"),
                    TableXlsxBuilder.cell("2021-06-01"),
                    TableXlsxBuilder.cell("60000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("2021-07-01"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00")
                )
        );

        tagsMap.put(
            "TABLE_PART_5",
            TableXlsxBuilder.create()
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Петров Петр Петрович"),
                    TableXlsxBuilder.cell("67890"),
                    TableXlsxBuilder.cell("Аналитик"),
                    TableXlsxBuilder.cell(""),
                    TableXlsxBuilder.cell("1985-05-15"),
                    TableXlsxBuilder.cell("39"),
                    TableXlsxBuilder.cell("2019-03-15"),
                    TableXlsxBuilder.cell("75000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("2019-04-15"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Сидоров Сидор Сидорович"),
                    TableXlsxBuilder.cell("54321"),
                    TableXlsxBuilder.cell("Разработчик"),
                    TableXlsxBuilder.cell("IT отдел"),
                    TableXlsxBuilder.cell("1992-12-20"),
                    TableXlsxBuilder.cell("32"),
                    TableXlsxBuilder.cell("2021-06-01"),
                    TableXlsxBuilder.cell("60000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("2021-07-01"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00")
                )
        );

        tagsMap.put(
            "TABLE_PART_6",
            TableXlsxBuilder.create()
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Иванов Иван Иванович"),
                    TableXlsxBuilder.cell("12345"),
                    TableXlsxBuilder.cell("Менеджер"),
                    TableXlsxBuilder.cell("Главный офис"),
                    TableXlsxBuilder.cell("1990-01-01"),
                    TableXlsxBuilder.cell("34"),
                    TableXlsxBuilder.cell("2020-01-01"),
                    TableXlsxBuilder.cell("50000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("2020-02-01"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00"),
                    TableXlsxBuilder.cell("10000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Петров Петр Петрович"),
                    TableXlsxBuilder.cell("67890"),
                    TableXlsxBuilder.cell("Аналитик"),
                    TableXlsxBuilder.cell(""),
                    TableXlsxBuilder.cell("1985-05-15"),
                    TableXlsxBuilder.cell("39"),
                    TableXlsxBuilder.cell("2019-03-15"),
                    TableXlsxBuilder.cell("75000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("2019-04-15"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00"),
                    TableXlsxBuilder.cell("15000.00")
                )
                .row(
                    TableXlsxBuilder.cell("Сидоров Сидор Сидорович"),
                    TableXlsxBuilder.cell("54321"),
                    TableXlsxBuilder.cell("Разработчик"),
                    TableXlsxBuilder.cell("IT отдел"),
                    TableXlsxBuilder.cell("1992-12-20"),
                    TableXlsxBuilder.cell("32"),
                    TableXlsxBuilder.cell("2021-06-01"),
                    TableXlsxBuilder.cell("60000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("2021-07-01"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00"),
                    TableXlsxBuilder.cell("12000.00")
                )
        );

        ReportData data = new ReportData(tagsMap);

        GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.EMPTY_AND_LOG,
            true,
            Locale.forLanguageTag("ru-RU"), ZoneId.of("Europe/Moscow"),
            true
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("out/table_with_row_and_template_difficult_result.xlsx"), report.bytes());
        return report;
    }

    /**
     * Runs local XLSX rows-only scenarios and prints collected warnings.
     */
    @Test
    @SneakyThrows
    void shouldRunRowsOnlyManualScenarios() {
        ReportGeneratorService service = new ReportGeneratorServiceImpl();

        GeneratedReport reportXlsxRows = xlsxRows(service);
        GeneratedReport rowsWithTemplateBelow = xlsxRowsAndTemplateBellow(service);
        GeneratedReport rowsWithTemplateBelowDifficult = xlsxRowsAndTemplateBellowDifficult(service);

        for (List<GenerationWarning> warningList : List.of(
            reportXlsxRows.warnings(),
            rowsWithTemplateBelow.warnings(),
            rowsWithTemplateBelowDifficult.warnings()
        )) {
            for (GenerationWarning warning : warningList) {
                System.out.printf("[%s] %s @ %s%n", warning.code(), warning.message(), warning.location());
            }
        }
    }
}
