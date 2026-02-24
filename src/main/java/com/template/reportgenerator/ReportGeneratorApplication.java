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
import java.util.List;
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

        byte[] templateBytes = Files.readAllBytes(Path.of("/Users/onbozoyan/Downloads/report-generator/Book1.xlsx"));

        TemplateInput input = new TemplateInput(
            "Book1.xlsx",
            null,
            templateBytes
        );

        ReportData data = new ReportData(
            Map.of(
                "period", "2026-Q1",
                "total", "4550.75",
                "token", input.toString(),
                "TABLE_HERE", List.of(
                    Map.of("name", "North", "amount", 1200.25),
                    Map.of("name", "South", "amount", 900.00)
                )
            ),
            Map.of(),
            Map.of()
        );

        GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.FAIL_FAST,
            true, // recalculate formulas (для XLS/XLSX)
            Locale.forLanguageTag("ru-RU"),
            ZoneId.of("Europe/Moscow")
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("/Users/onbozoyan/Downloads/report-generator/sales-report.xlsx"), report.bytes());

        for (var warning : report.warnings()) {
            System.out.printf("[%s] %s @ %s%n", warning.code(), warning.message(), warning.location());
        }
    }

}
