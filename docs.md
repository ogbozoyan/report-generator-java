# Report Generator: документация по модулям, классам и тестам

## 1. Назначение
`report-generator` реализует библиотечный service-layer для генерации отчетов по шаблонам электронных таблиц.

Поддерживаемые форматы:
- `XLS`
- `XLSX`
- `ODS`

Поддерживаемый DSL в шаблонах:
- скалярные токены: `{{token}}`
- блоки строк: `[[TABLE_START:key]] ... [[TABLE_END:key]]`
- блоки колонок: `[[COL_START:key]] ... [[COL_END:key]]`

Ключевые функции:
- замена скалярных значений и значений из `item`/`index` в блоках
- динамическая экспансия строк и колонок
- политика отсутствующих значений (`MissingValuePolicy`)
- сбор warnings по проблемным местам
- пересчет формул для POI (`recalculateFormulas`)
- сохранение форматирования (стили, шрифты, размеры колонок/строк, merged-регионы в XLS/XLSX)

## 2. Архитектурные модули

### 2.1 Пакет `com.template.reportgenerator`
- Оркестрация генерации.
- Точка входа Spring Boot.

### 2.2 Пакет `com.template.reportgenerator.api`
- Публичный контракт библиотеки: вход, выход, опции, политика.

### 2.3 Пакет `com.template.reportgenerator.util`
- Детектор формата шаблона.
- Сканирование DSL-маркеров/токенов.
- Валидация структуры блоков.
- Процессинг workbook для POI и ODS.
- Резолв токенов и запись типизированных значений.
- Сериализация результата и сбор предупреждений.

### 2.4 Пакет `com.template.reportgenerator.exception`
- Доменные исключения для формата, синтаксиса, структуры, привязки данных и I/O.

## 3. Поток выполнения `generate(...)`
Реализован в `ReportGeneratorServiceImpl`:
1. Проверка входа (`template != null`).
2. Нормализация `ReportData` и `GenerateOptions`.
3. Определение формата (`TemplateFormatDetector`).
4. Создание процессора (`PoiWorkbookProcessor` или `OdsWorkbookProcessor`).
5. Скан шаблона (`TemplateScanner`).
6. Валидация блоков (`TemplateValidator`) и построение `BlockRegion`.
7. Применение scalar-токенов.
8. Экспансия TABLE-блоков.
9. Экспансия COL-блоков.
10. Очистка marker-ячеек.
11. Пересчет формул (для POI по флагу).
12. Сериализация результата (`ReportSerializer`) + warnings.

## 4. Описание классов по пакетам

### 4.1 `com.template.reportgenerator`

| Класс | Ответственность | Тесты |
|---|---|---|
| `ReportGeneratorServiceImpl` | Главный оркестратор pipeline генерации. Делегирует в `internal`-компоненты. | Прямо: `ReportGeneratorServiceImplTest`, `ReportGeneratorFormattingGoldenTest` |
| `ReportGeneratorApplication` | Bootstrap Spring Boot приложения. | Прямо: `ReportGeneratorApplicationTests.contextLoads` |

### 4.2 `com.template.reportgenerator.api`

| Класс | Ответственность | Тесты |
|---|---|---|
| `ReportGeneratorService` | Публичный интерфейс сервиса генерации. | Косвенно: все интеграционные сценарии через `service.generate(...)` |
| `TemplateInput` | Вход шаблона: `fileName`, `contentType`, `bytes` (`bytes` обязателен). | Косвенно: все тесты генерации, `TemplateFormatDetectorTest` |
| `ReportData` | Данные для биндинга: `scalars`, `tables`, `columns` (с null-safe инициализацией). | Косвенно: `ReportGeneratorServiceImplTest`, `ReportGeneratorFormattingGoldenTest` |
| `GenerateOptions` | Опции генерации: политика пропусков, пересчет формул, locale/zoneId с дефолтами. | Косвенно: `ReportGeneratorServiceImplTest` (дефолтный путь) |
| `GeneratedReport` | Результат: имя файла, content type, bytes, immutable warnings. | Косвенно: `ReportGeneratorServiceImplTest`, `ReportGeneratorFormattingGoldenTest` |
| `GenerationWarning` | DTO предупреждения (`code/message/location`). | Косвенно: проверка `MISSING_TOKEN` в `ReportGeneratorServiceImplTest.shouldReplaceMissingTokenWithEmptyAndWarningByDefault` |
| `MissingValuePolicy` | Политика обработки отсутствующих токенов: `EMPTY_AND_LOG`, `LEAVE_TOKEN`, `FAIL_FAST`. | Частично косвенно: `EMPTY_AND_LOG` покрыт тестами; `LEAVE_TOKEN`/`FAIL_FAST` отдельными тестами пока не покрыты |

### 4.3 `com.template.reportgenerator.util`

| Класс | Ответственность | Тесты |
|---|---|---|
| `TemplateFormat` | Enum формата + `contentType` и расширение. | Прямо/косвенно: `TemplateFormatDetectorTest`, `ReportGeneratorServiceImplTest` |
| `TemplateFormatDetector` | Определение формата по extension/content-type/magic-bytes. | Прямо: `TemplateFormatDetectorTest` |
| `TemplateScanner` | Поиск block-маркеров и scalar-токенов в POI/ODS. | Косвенно: `ReportGeneratorServiceImplTest`, `ReportGeneratorFormattingGoldenTest`, ошибки маркеров в `shouldFailOnUnpairedMarkers` |
| `TemplateScanResult` | Результат сканирования (`markers`, `scalarTokens`). | Прямо: используется в `TemplateValidatorTest` |
| `BlockMarker` | Представление START/END marker c типом блока и позицией. | Прямо: `TemplateValidatorTest` |
| `TokenOccurrence` | Найденный scalar token и позиция в шаблоне. | Косвенно: через scanner/validator/service flow |
| `CellPosition` | Координаты ячейки + текстовый `asLocation()`. | Прямо: `TemplateValidatorTest` (создание сценариев) |
| `BlockType` | Тип блока (`TABLE`, `COL`). | Прямо: `TemplateValidatorTest`; косвенно: все сценарии block expansion |
| `BlockRegion` | Валидированный регион блока + расчет внутренней области. | Прямо: `TemplateValidatorTest`; косвенно: экспансия в сервисных тестах |
| `TemplateValidator` | Проверка парности START/END, прямоугольника, пустой внутренней зоны, overlap/nesting. | Прямо: `TemplateValidatorTest`; косвенно: `ReportGeneratorServiceImplTest.shouldFailOnUnpairedMarkers` |
| `WorkbookProcessor` | Контракт процессора форматов (scan/apply/expand/clear/recalc/serialize). | Косвенно: через `ReportGeneratorServiceImpl` |
| `PoiWorkbookProcessor` | Реализация для XLS/XLSX (POI): токены, TABLE/COL экспансия, стили, merged, формулы, сериализация. | Косвенно: `ReportGeneratorServiceImplTest`, `ReportGeneratorFormattingGoldenTest` |
| `OdsWorkbookProcessor` | Реализация для ODS (ODFDOM): токены, TABLE/COL экспансия, стили/размеры, сериализация. | Косвенно: `ReportGeneratorServiceImplTest`, `ReportGeneratorFormattingGoldenTest` |
| `TokenResolver` | Поиск/резолв токенов в тексте, поддержка dot-path и политики пропусков. | Косвенно: все генерационные тесты |
| `ResolvedText` | Результат резолва текста (`value`, `changed`). | Косвенно: через `TokenResolver` в генерационных тестах |
| `ValueWriter` | Типизированная запись значений в POI/ODS (числа, bool, даты, java.time, строки). | Косвенно: `ReportGeneratorServiceImplTest` |
| `WarningCollector` | Коллектор warnings и безопасная выдача immutable списка. | Косвенно: `ReportGeneratorServiceImplTest.shouldReplaceMissingTokenWithEmptyAndWarningByDefault` |
| `ReportSerializer` | Формирование `GeneratedReport`, нормализация имени и content type. | Косвенно: все генерационные тесты |

### 4.4 `com.template.reportgenerator.exception`

| Класс | Когда используется | Тесты |
|---|---|---|
| `UnsupportedTemplateFormatException` | Не удалось определить формат шаблона. | Прямо: `TemplateFormatDetectorTest.shouldThrowOnUnknown` |
| `TemplateSyntaxException` | Ошибки DSL/парности/геометрии блоков. | Прямо: `TemplateValidatorTest.shouldFailOnUnpairedMarkers`; косвенно: `ReportGeneratorServiceImplTest.shouldFailOnUnpairedMarkers` |
| `TemplateStructureException` | Пересекающиеся/вложенные блоки. | Прямо: `TemplateValidatorTest.shouldFailOnOverlappingBlocks` |
| `TemplateDataBindingException` | Не найден токен при `FAIL_FAST`. | Прямых тестов сейчас нет |
| `TemplateReadWriteException` | Ошибка чтения/записи workbook. | Прямых тестов сейчас нет |

## 5. Тестовый модуль: что проверяется

### 5.1 `ReportGeneratorServiceImplTest`
Проверяет end-to-end сценарии:
- `shouldGenerateXlsxAndKeepCellStyleAndWidth`:
  - scalar replacement в XLSX
  - сохранение column width и стиля ячейки
- `shouldGenerateXls`:
  - scalar replacement в XLS
- `shouldGenerateOds`:
  - scalar replacement в ODS
  - сохранение ширины колонки и alignment
- `shouldExpandTableBlockInXlsx`:
  - экспансия TABLE-блока и копирование стиля
- `shouldExpandColumnBlockInXlsx`:
  - экспансия COL-блока
- `shouldReplaceMissingTokenWithEmptyAndWarningByDefault`:
  - поведение `MissingValuePolicy.EMPTY_AND_LOG`
  - генерация warning `MISSING_TOKEN`
- `shouldFailOnUnpairedMarkers`:
  - ошибка синтаксиса при непарных маркерах

### 5.2 `ReportGeneratorFormattingGoldenTest`
Golden regression-тесты на форматирование:
- `shouldPreserveFormattingAndMergedRegionsForExpandedTableInXlsx`:
  - значения, width, row height, font/style, wrap/alignment, merged-копирование
- `shouldPreserveFormattingAndMergedRegionsForExpandedColumnsInXlsx`:
  - значения, width, font/style, merged-копирование
- `shouldPreserveStylesAndWidthsForExpandedTableInOds`:
  - значения, width, row height, alignment, wrap
- `shouldPreserveStylesAndWidthsForExpandedColumnsInOds`:
  - значения, width, alignment, wrap

### 5.3 `TemplateFormatDetectorTest`
- `shouldDetectByExtension`
- `shouldDetectByMagicBytes`
- `shouldThrowOnUnknown`

Покрывает логику определения формата и ошибку unsupported формата.

### 5.4 `TemplateValidatorTest`
- `shouldBuildRegionsForValidMarkers`
- `shouldFailOnUnpairedMarkers`
- `shouldFailOnOverlappingBlocks`

Покрывает структурную валидацию DSL-блоков.

### 5.5 `ReportGeneratorApplicationTests`
- `contextLoads` — smoke-test поднятия Spring-контекста.

## 6. Матрица покрытия «класс -> тест»

### 6.1 Прямо покрытые отдельными юнит-тестами
- `TemplateFormatDetector` -> `TemplateFormatDetectorTest`
- `TemplateValidator` -> `TemplateValidatorTest`
- `ReportGeneratorServiceImpl` -> `ReportGeneratorServiceImplTest`, `ReportGeneratorFormattingGoldenTest`
- `ReportGeneratorApplication` -> `ReportGeneratorApplicationTests`

### 6.2 Косвенно покрытые интеграционными сценариями
- `PoiWorkbookProcessor`, `OdsWorkbookProcessor`
- `TemplateScanner`, `TokenResolver`, `ValueWriter`
- `ReportSerializer`, `WarningCollector`
- `BlockRegion`, `BlockMarker`, `CellPosition`, `TemplateScanResult`, `TokenOccurrence`
- DTO/API records (`TemplateInput`, `ReportData`, `GenerateOptions`, `GeneratedReport`, `GenerationWarning`)

## 7. Текущие пробелы тестового покрытия
- Нет прямых тестов для:
  - `TemplateDataBindingException` (`MissingValuePolicy.FAIL_FAST`)
  - `TemplateReadWriteException` (ошибки чтения/сохранения поврежденных шаблонов)
  - ветки `MissingValuePolicy.LEAVE_TOKEN`
- ODS merged-регионы не зафиксированы отдельным golden-тестом (в отличие от XLSX).

## 8. Быстрый чек-лист при изменениях
- Любое изменение DSL -> обновить `TemplateScanner`/`TemplateValidator` + тесты validator/service.
- Любое изменение копирования формата -> обновить golden-тесты в `ReportGeneratorFormattingGoldenTest`.
- Любое изменение детектора формата -> обновить `TemplateFormatDetectorTest`.
- Любое изменение `MissingValuePolicy` -> добавить/обновить сценарии в `ReportGeneratorServiceImplTest`.

## 9. Пример файлов

### 9.1 Пример шаблона (логическая раскладка `sales-template.xlsx`)
Ниже показан пример содержимого ячеек (не бинарный `.xlsx`, а схема заполнения):

| Ячейка | Значение |
|---|---|
| `A1` | `Sales report: {{period}}` |
| `A3` | `[[TABLE_START:rows]]` |
| `A4` | `Item` |
| `B4` | `Amount` |
| `A5` | `{{item.name}}` |
| `B5` | `{{item.amount}}` |
| `C6` | `[[TABLE_END:rows]]` |
| `A8` | `Total: {{total}}` |

Что это значит:
- `period` и `total` берутся из `scalars`
- блок `rows` берется из `tables.rows`
- строка `A5:B5` копируется для каждого `item` из `rows`

### 9.2 Пример данных (`report-data.json`)
```json
{
  "scalars": {
    "period": "2026-Q1",
    "total": 4550.75
  },
  "tables": {
    "rows": [
      { "name": "North", "amount": 1200.25 },
      { "name": "South", "amount": 900.00 },
      { "name": "West", "amount": 2450.50 }
    ]
  },
  "columns": {}
}
```

### 9.3 Пример шаблона с колонками (логическая раскладка)

| Ячейка | Значение |
|---|---|
| `A1` | `[[COL_START:months]]` |
| `B2` | `{{item.name}}` |
| `C4` | `[[COL_END:months]]` |

Что это значит:
- блок `months` берется из `columns.months`
- колонка с `{{item.name}}` будет размножена по числу элементов

## 10. Usage (Java)

```java
import com.template.reportgenerator.service.ReportGeneratorServiceImpl;
import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.GeneratedReport;
import com.template.reportgenerator.dto.MissingValuePolicy;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.service.ReportGeneratorService;
import com.template.reportgenerator.dto.TemplateInput;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.ZoneId;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class ExampleUsage {
  public static void main(String[] args) throws Exception {
    ReportGeneratorService service = new ReportGeneratorServiceImpl();

    byte[] templateBytes = Files.readAllBytes(Path.of("templates/sales-template.xlsx"));

    TemplateInput input = new TemplateInput(
            "sales-template.xlsx",
            null,
            templateBytes
    );

    ReportData data = new ReportData(
            Map.of(
                    "period", "2026-Q1",
                    "total", 4550.75
            ),
            Map.of(
                    "rows", List.of(
                            Map.of("name", "North", "amount", 1200.25),
                            Map.of("name", "South", "amount", 900.00),
                            Map.of("name", "West", "amount", 2450.50)
                    )
            ),
            Map.of()
    );

    GenerateOptions options = new GenerateOptions(
            MissingValuePolicy.EMPTY_AND_LOG,
            true, // recalculate formulas (для XLS/XLSX)
            Locale.forLanguageTag("ru-RU"),
            ZoneId.of("Europe/Moscow")
    );

    GeneratedReport report = service.generate(input, data, options);

    Files.createDirectories(Path.of("out"));
    Files.write(Path.of("out/sales-report.xlsx"), report.bytes());

    for (var warning : report.warnings()) {
      System.out.printf("[%s] %s @ %s%n", warning.code(), warning.message(), warning.location());
    }
  }
}
```

Минимальный usage без явных опций:
```java
GeneratedReport report = service.generate(input, data, null);
```
