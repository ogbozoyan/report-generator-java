# Report Generator: внутренняя документация

## 1. Цель документа

`docs.md` описывает внутреннее устройство библиотеки: pipeline обработки, алгоритмы, компромиссы и причины архитектурных
решений.

Этот документ ориентирован на разработчиков, которые поддерживают или расширяют кодовую базу.

## 2. Pipeline генерации

Класс: `com.template.reportgenerator.io.github.ogbozoyan.service.ReportGeneratorServiceImpl`

Поток `generate(...)`:

1. Валидация входных аргументов.
2. Разрешение `GenerateOptions` (подстановка defaults).
3. Детект фактического входного формата (`TemplateFormatDetector.detectFormat`).
4. Детект желаемого выходного формата (`TemplateFormatDetector.detectRequestedOutputFormat`).
5. Проверка допустимости маршрута конвертации.
6. Создание форматного процессора (`WorkbookProcessor` implementation).
7. Применение токенов (`applyTemplateTokens`).
8. Пересчёт формул (только где поддерживается).
9. Сериализация обработанного документа.
10. Опциональный post-convert (`DocumentFormatConverter`) для `ODS/ODT` выгрузки.
11. Нормализация output metadata и сбор warnings (`ReportSerializer`).

Почему так:

- pipeline разделён на чистые этапы, чтобы локализовать ответственность и упрощать диагностику;
- форматная логика вынесена в процессоры, сервис остаётся оркестратором;
- post-convert отделён от этапа подстановки, чтобы не усложнять алгоритм токенизации.

## 3. Контракты данных

### 3.1 `TemplateInput`

- `fileName`: имя входного шаблона или ожидаемого output.
- `contentType`: optional MIME hint.
- `bytes`: байты шаблона.

### 3.2 `ReportData`

- `templateTokens`: единая карта данных `Map<String, Object>`.
- scalar token: любое строковое/числовое/дата-значение.
- table token (default mode): `List<Map<String, Object>>`.
- table token (rows-only mode): `List<Object[]>`.
- optional порядок колонок: `TOKEN__columns` (или `TOKEN_columns`, `TOKEN.columns`).

### 3.3 `GenerateOptions`

- `missingValuePolicy`: поведение при отсутствии токена.
- `recalculateFormulas`: пересчёт формул для spreadsheet.
- `rowsOnlyTableTokens`: глобальный rows-only режим вставки table token для `XLS/XLSX`
  (без header-строки, ожидается `List<Object[]>`).
- `locale`, `zoneId`: локализация и time-zone для записи дат.

## 4. Карта модулей и ответственности

### 4.1 `io.github.ogbozoyan.service/*`

- `ReportGeneratorService`: публичный API генерации.
- `ReportGeneratorServiceImpl`: orchestration и маршрутизация по форматам.

### 4.2 `io.github.ogbozoyan.processor/*`

- `WorkbookProcessor`: единый lifecycle-контракт форматных обработчиков.
- `PoiWorkbookProcessor`: `XLS/XLSX` таблицы, типизированная запись значений, auto-width, формулы.
- `DocxDocumentProcessor`: работа по дереву body/table/cell, вставка таблиц в корректный контейнер.
- `DocDocumentProcessor`: basic text-table в `.doc`.
- `PdfDocumentProcessor`: text reconstruction и ASCII-grid таблицы.

### 4.3 `io.github.ogbozoyan.util/*`

- `TemplateFormatDetector`: format detection по magic bytes/extension/MIME.
- `TokenResolver`: поиск/резолв токенов и table-typing.
- `WarningCollector`: накопление non-fatal предупреждений.
- `ReportSerializer`: fileName/contentType/warnings финального результата.
- `LibreOfficeDocumentFormatConverter`: post-convert в `ODS/ODT`.
- `TemplateScanner`, `TemplateValidator`: scan/validation helpers для legacy-DSL сценариев.

### 4.4 `contract/*` и `exception/*`

- типы контрактов для передачи данных между слоями;
- явные типы исключений для чтения/формата/синтаксиса/структуры/биндинга.

## 5. Алгоритмы и почему выбран такой подход

## 5.1 `WorkbookProcessor` (единый контракт и lifecycle)

Контракт:

- `scan()`;
- `applyTemplateTokens(...)`;
- `recalculateFormulas(...)` (default no-op);
- `serialize()`;
- `close()`.

Почему:

- единый интерфейс позволяет сервису оставаться независимым от деталей формата;
- `default` для `recalculateFormulas` не заставляет non-spreadsheet процессоры реализовывать неактуальную логику;
- `AutoCloseable` делает ресурсную дисциплину одинаковой для всех реализаций.

## 5.2 `PoiWorkbookProcessor`

Ключевые алгоритмы:

- sparse traversal: обход только физически существующих строк/ячеек.
- anchor-first strategy: сначала сбор якорей таблиц, потом вставка.
- reverse apply: вставка таблиц снизу вверх по листу.
- dual table modes:
  - default: header + data;
  - rows-only (`GenerateOptions.rowsOnlyTableTokens=true`): только data-строки из `List<Object[]>`.
- multi-pass table expansion:
  - сначала выполняются только table-pass проходы;
  - каждый проход: scan anchors -> reverse insert;
  - проходы повторяются, пока есть новые anchors;
  - после стабилизации выполняется scalar-pass.
- style baseline: reuse стиля маркерной ячейки для header/data.
- auto-width: ширины меняются только у вставленных колонок.
- formula policy: формулы с токенами не переписываются, только warning.

Почему:

- sparse traversal устраняет зависания на больших разреженных листах;
- reverse apply делает множественные вставки детерминированными при `shiftRows`;
- rows-only режим позволяет использовать отдельную строку descriptor/маппинга без дублирования заголовка;
- multi-pass устраняет потерю table token, которые появляются только после предыдущей вставки;
- baseline style минимизирует визуальные регрессии шаблонов;
- локальный auto-width не ломает внешний layout листа;
- пропуск formula-токенов безопаснее, чем риск повреждения формульного синтаксиса.

## 5.3 `DocxDocumentProcessor`

Ключевые алгоритмы:

- recursive traversal по `IBody`: document body -> table -> cell -> nested body;
- сбор `ParagraphTarget` с порядком обхода;
- table anchors применяются в обратном порядке;
- вставка таблицы строго в контейнер абзаца (`XWPFDocument` или `XWPFTableCell`);
- placeholder paragraph удаляется из исходного контейнера после вставки.

Почему:

- DOCX часто содержит токены внутри существующих таблиц, а не только в body-параграфах;
- корректный контейнер вставки устраняет кейс, когда таблица создавалась не там, где стоял placeholder;
- удаление placeholder-абзаца предотвращает дублирование контента.

## 5.4 `DocDocumentProcessor`

Ключевые алгоритмы:

- exact paragraph placeholder распознаётся в `HWPF Range`;
- table token рендерится как text-grid: header/rows с разделителями `\t` и `\r`;
- scalar токены заменяются массовым `range.replaceText(...)`.

Почему:

- `.doc` (HWPF) ограничен по возможностям безопасного структурного редактирования;
- text-grid даёт устойчивую "basic" поддержку, достаточную для простых отчётов;
- подход минимизирует риск повреждения бинарной структуры `.doc`.

Ограничение:

- это не полноценная Word table model, а текстовая имитация таблицы.

## 5.5 `PdfDocumentProcessor`

Ключевые алгоритмы:

- PDF читается как текст (`PDFTextStripper`), затем выполняется токен-замена;
- table token рендерится как ASCII-grid;
- сериализация строит новый PDF построчно с word-wrap и pagination.

Почему:

- PDF не поддерживает надёжное in-place редактирование текстовых объектов без сложной геометрической реконструкции;
- text reconstruction обеспечивает предсказуемый результат и устойчивость;
- ASCII-grid даёт переносимую репрезентацию таблиц для текстового потока.

Ограничение:

- layout выходного PDF не является 1:1 копией исходного шаблона.

## 6. Failure modes и troubleshooting

- `MISSING_TOKEN`:
  - токен отсутствует в `templateTokens`; проверьте ключи и `missingValuePolicy`.
- `TABLE_TOKEN_INVALID`:
  - в default mode значение токена не является `List<Map<String,Object>>`;
  - в rows-only mode значение токена не является `List<Object[]>`.
- `TABLE_TOKEN_EMPTY`:
  - таблица передана как пустой список.
- `TABLE_TOKEN_RECURSIVE`:
  - вставки table token не стабилизировались за лимит `MAX_TABLE_PASSES`;
  - проверьте, что table-токены не ссылаются циклически друг на друга.
- `GenerateOptions.rowsOnlyTableTokens=true`:
  - включён rows-only режим для `XLS/XLSX`: marker row становится первой data-строкой;
  - payload токена должен быть `List<Object[]>`.
- `TABLE_TOKEN_INLINE_IGNORED`:
  - table token найден inline и не вставлен как таблица.
- `TABLE_TOKEN_INLINE_TEXT_DROPPED`:
  - для single-token режима рядом был статический текст, он отброшен при вставке таблицы.
- `FORMULA_TOKEN_SKIPPED`:
  - токен находится в formula cell; формула оставлена без изменений.
- `Unsupported output conversion`:
  - разрешены только `XLS/XLSX -> ODS` и `DOC/DOCX -> ODT`.
- `UnsupportedTemplateFormatException` для входного `ODS/ODT`:
  - используйте входной `XLS/XLSX` либо `DOC/DOCX`, затем запрашивайте `ODS/ODT` на выходе.
- Ошибки конвертации LibreOffice:
  - проверьте наличие `soffice`/`libreoffice` в `PATH`.

## 7. Примеры использования

### 7.1 XLSX с table token

```java
ReportGeneratorService serviceI = new io.github.ogbozoyan.service.ReportGeneratorServiceImpl();

TemplateInput input = new TemplateInput("TABLE_BOOK.xlsx", null, xlsxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "report_year", 2026,
        "Table_2", List.of(
                Map.of("amount", 1200.25, "name", "North"),
                Map.of("amount", 900.00, "name", "South")
        ),
        "Table_2__columns", List.of("name", "amount")
));

GenerateOptions options = new GenerateOptions(
        MissingValuePolicy.EMPTY_AND_LOG,
        true,
        Locale.getDefault(),
        ZoneId.systemDefault(),
        false
);

GeneratedReport report = serviceI.generate(input, data, options);
```

### 7.2 DOCX: table token внутри существующей таблицы

Условие шаблона:

- в ячейке DOCX-таблицы есть отдельный paragraph с `{{inner_table}}`.

Пример:

```java
TemplateInput input = new TemplateInput("DOC1.docx", null, docxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "inner_table", List.of(
                Map.of("kpi", "Revenue", "value", "125000"),
                Map.of("kpi", "Margin", "value", "24%")
        )
));

GeneratedReport report = io.github.ogbozoyan.service.generate(input, data, GenerateOptions.defaults());
```

### 7.3 XLSX -> ODS

```java
TemplateInput input = new TemplateInput("sales-report.ods", null, xlsxTemplateBytes);
GeneratedReport report = io.github.ogbozoyan.service.generate(input, data, GenerateOptions.defaults());
```

### 7.4 DOCX -> ODT

```java
TemplateInput input = new TemplateInput(
        "letter.odt",
        "application/vnd.oasis.opendocument.text",
        docxTemplateBytes
);
GeneratedReport report = io.github.ogbozoyan.service.generate(input, data, GenerateOptions.defaults());
```

## 8. Связь решений с тестами

Ключевые тестовые наборы и что они подтверждают:

- `src/test/java/com/template/reportgenerator/io.github.ogbozoyan.service.ReportGeneratorServiceImplTest.java`
  - сервисный pipeline;
  - вставка таблиц в `XLS/XLSX` и non-spreadsheet форматах;
  - порядок колонок;
  - inline/exact-placeholder поведение;
  - поддерживаемые маршруты post-convert.

- `src/test/java/com/template/reportgenerator/io.github.ogbozoyan.service.ReportGeneratorFormattingGoldenTest.java`
  - регрессионная проверка форматирования spreadsheet при table insertion.

- `src/test/java/io/github/ogbozoyan/integration/ReportGeneratorManualIntegrationTest.java`
  - ручные интеграционные сценарии, вынесенные из `main` (класс помечен `@Disabled`).

- `src/test/java/com/template/reportgenerator/io.github.ogbozoyan.util/TemplateFormatDetectorTest.java`
  - детект формата по magic bytes/content-type/extension;
  - различение OLE2 (`DOC` vs `XLS`);
  - маршрутизация requested output format.

- `src/test/java/com/template/reportgenerator/io.github.ogbozoyan.util/TemplateValidatorTest.java`
  - корректность scan/validation helper-логики для block-маркеров.

## 9. Почему разделены `README.md` и `docs.md`

- `README.md` отвечает на вопросы "что это" и "как быстро запустить".
- `docs.md` отвечает на вопросы "как это реализовано" и "почему именно так".

Это уменьшает дублирование и упрощает сопровождение документации при изменениях алгоритмов.
