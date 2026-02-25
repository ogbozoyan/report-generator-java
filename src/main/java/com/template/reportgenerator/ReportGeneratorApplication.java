package com.template.reportgenerator;

import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.GeneratedReport;
import com.template.reportgenerator.dto.MissingValuePolicy;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.TemplateInput;
import com.template.reportgenerator.service.ReportGeneratorService;
import com.template.reportgenerator.service.ReportGeneratorServiceImpl;
import lombok.SneakyThrows;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.ZoneId;
import java.util.Locale;
import java.util.Map;

/**
 * Spring Boot entrypoint and local manual runner for quick generation checks.
 */
@SpringBootApplication
public class ReportGeneratorApplication {

    @SneakyThrows
    static void main(String[] args) {
//        SpringApplication.run(ReportGeneratorApplication.class, args);
        ReportGeneratorService service = new ReportGeneratorServiceImpl();

        byte[] templateBytes = Files.readAllBytes(Path.of("/Users/onbozoyan/Downloads/report-generator/отчет.xlsx"));

        TemplateInput input = new TemplateInput(
            "отчет.xlsx",
            null,
            templateBytes
        );

        ReportData data = new ReportData(
            Map.of(
//                "period", "2026-Q1",
//                "SOME_PLACEHOLDER", "ПУПУПУ",
//                "total", "4550.75",
//                "token", input.toString(),
//                "TABLE_HERE", List.of(
//                    //БАГ, колонки в неправильном порядке
//                    Map.of("name", "North", "amount", 1200.25),
//                    Map.of("name", "South", "amount", 900.00),
//                    Map.of("жопа", "South", "слона", 900.00)
//                ).reversed(),
//                "ANOTHER_TABLE", List.of(
//                    Map.of(
//                        "a", "b", "c", "d"
//                    )
//                ),
//                "mega_test", "mega_value",
                "year", "1999"
            )
        );

        GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.FAIL_FAST,
            true, // recalculate formulas (для XLS/XLSX)
            Locale.forLanguageTag("ru-RU"),
            ZoneId.of("Europe/Moscow")
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("/Users/onbozoyan/Downloads/report-generator/result_отчет.xlsx"), report.bytes());

        for (var warning : report.warnings()) {
            System.out.printf("[%s] %s @ %s%n", warning.code(), warning.message(), warning.location());
        }
    }

}
