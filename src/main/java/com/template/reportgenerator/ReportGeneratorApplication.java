package com.template.reportgenerator;

import com.template.reportgenerator.contract.GenerateOptions;
import com.template.reportgenerator.contract.GeneratedReport;
import com.template.reportgenerator.contract.GenerationWarning;
import com.template.reportgenerator.contract.MissingValuePolicy;
import com.template.reportgenerator.contract.ReportData;
import com.template.reportgenerator.contract.TemplateInput;
import com.template.reportgenerator.service.ReportGeneratorService;
import com.template.reportgenerator.service.ReportGeneratorServiceImpl;
import lombok.SneakyThrows;
import org.jspecify.annotations.NonNull;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.ZoneId;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * Spring Boot entrypoint and local manual runner for quick generation checks.
 */
@SpringBootApplication
public class ReportGeneratorApplication {

    /**
     * Local manual runner for quick integration checks.
     *
     * @param args command-line arguments
     */
    @SneakyThrows
    static void main(String[] args) {
//        SpringApplication.run(ReportGeneratorApplication.class, args);
        ReportGeneratorService service = new ReportGeneratorServiceImpl();

        GeneratedReport reportXlsx = xlsx(service);
        GeneratedReport reportDocx = docx(service);

        for (List<GenerationWarning> warningList : List.of(reportXlsx.warnings(), reportDocx.warnings())) {
            for (GenerationWarning warning : warningList) {
                System.out.printf("[%s] %s @ %s%n", warning.code(), warning.message(), warning.location());
            }
        }
    }

    /**
     * Generates DOCX sample report used for local smoke testing.
     *
     * @param service report generator service
     * @return generated report
     * @throws IOException when template/result file I/O fails
     */
    private static GeneratedReport docx(ReportGeneratorService service) throws IOException {
        String fileName = "DOC1.docx";
        String resultFileName = "DOC1_result.docx";
        byte[] templateBytes = Files.readAllBytes(Path.of("/Users/onbozoyan/Downloads/report-generator/" + fileName));
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
            Locale.forLanguageTag("ru-RU"),
            ZoneId.of("Europe/Moscow")
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("/Users/onbozoyan/Downloads/report-generator/" + resultFileName), report.bytes());
        return report;
    }

    /**
     * Generates XLSX sample report used for local smoke testing.
     *
     * @param service report generator service
     * @return generated report
     * @throws IOException when template/result file I/O fails
     */
    private static @NonNull GeneratedReport xlsx(ReportGeneratorService service) throws IOException {
        byte[] templateBytes = Files.readAllBytes(Path.of("/Users/onbozoyan/Downloads/report-generator/TABLE_BOOK.xlsx"));

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
            Locale.forLanguageTag("ru-RU"),
            ZoneId.of("Europe/Moscow")
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("/Users/onbozoyan/Downloads/report-generator/book_table.xlsx"), report.bytes());
        return report;
    }

}
