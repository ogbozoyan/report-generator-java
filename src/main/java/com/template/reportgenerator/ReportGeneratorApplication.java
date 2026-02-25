package com.template.reportgenerator;

import com.template.reportgenerator.contract.GenerateOptions;
import com.template.reportgenerator.contract.GeneratedReport;
import com.template.reportgenerator.contract.MissingValuePolicy;
import com.template.reportgenerator.contract.ReportData;
import com.template.reportgenerator.contract.TemplateInput;
import com.template.reportgenerator.service.ReportGeneratorService;
import com.template.reportgenerator.service.ReportGeneratorServiceImpl;
import lombok.SneakyThrows;
import org.springframework.boot.autoconfigure.SpringBootApplication;

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

    @SneakyThrows
    static void main(String[] args) {
//        SpringApplication.run(ReportGeneratorApplication.class, args);
        ReportGeneratorService service = new ReportGeneratorServiceImpl();

        byte[] templateBytes = Files.readAllBytes(Path.of("/Users/onbozoyan/Downloads/report-generator/Отчет_о_компенсации_процентовl_—_ТЭГ.ods"));

        TemplateInput input = new TemplateInput(
            "Отчет_о_компенсации_процентовl_—_ТЭГ.ods",
            null,
            templateBytes
        );

        Map<String, Object> tagsMap = new HashMap<>();
        tagsMap.put("TABLE_2", List.of(
            //БАГ, колонки в неправильном порядке
            Map.of("name", "North", "amount", 1200.25),
            Map.of("name", "South", "amount", 900.00),
            Map.of("жопа", "South", "слона", 900.00)
        ));
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
        tagsMap.put("TABLE_1", List.of(
            //БАГ, колонки в неправильном порядке
            Map.of("name", "North", "amount", 1200.25),
            Map.of("name", "South", "amount", 900.00),
            Map.of("жопа", "South", "слона", 900.00)
            )
        );
        tagsMap.put(
            "TABLE_3", List.of(
                //БАГ, колонки в неправильном порядке
                Map.of("name", "North", "amount", 1200.25),
                Map.of("name", "South", "amount", 900.00),
                Map.of("жопа", "South", "слона", 900.00)
            )
        );

        ReportData data = new ReportData(tagsMap);

        GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.FAIL_FAST,
            true, // recalculate formulas (для XLS/XLSX)
            Locale.forLanguageTag("ru-RU"),
            ZoneId.of("Europe/Moscow")
        );

        GeneratedReport report = service.generate(input, data, options);

        Files.createDirectories(Path.of("out"));
        Files.write(Path.of("/Users/onbozoyan/Downloads/report-generator/result_отчет.ods"), report.bytes());

        for (var warning : report.warnings()) {
            System.out.printf("[%s] %s @ %s%n", warning.code(), warning.message(), warning.location());
        }
    }

}
