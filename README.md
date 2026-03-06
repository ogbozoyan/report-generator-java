# Report Generator

`report-generator` — библиотека генерации документов по шаблонам с маркерами `{{TOKEN}}`.

Библиотека предназначена для использования как io.github.ogbozoyan.service-layer (без REST): вы передаёте байты шаблона,
данные токенов и
опции генерации, на выходе получаете готовый файл и список предупреждений.

## Что умеет

- scalar tokens: подстановка значений в `{{TOKEN}}`;
- table tokens: `{{TABLE_TOKEN}}` для значения типа `List<Map<String, Object>>`;
- declarative table tokens для `DOC/DOCX` через `TableBuilder`
  (создание таблицы по placeholder без заранее вставленной таблицы в шаблон);
- declarative table tokens для `XLS/XLSX` через `TableXlsxBuilder`
  (поддержка `colSpan` и `bold`, вставка по placeholder);
- rows-only table mode для `XLS/XLSX` через `GenerateOptions.rowsOnlyTableTokens=true`
  и значения типа `List<Object[]>` (вставка без header-строки);
- multi-pass обработка table tokens в `XLS/XLSX`: токены, появившиеся после вставки
  (например `{{TABLE_PART_2}}` внутри вставленных строк), обрабатываются на следующем проходе;
- единая модель данных для spreadsheet и non-spreadsheet форматов;
- политика отсутствующих токенов: `EMPTY_AND_LOG`, `LEAVE_TOKEN`, `FAIL_FAST`;
- пересчёт формул для `XLS/XLSX`;
- post-convert выгрузка:
  - `XLS/XLSX -> ODS`
  - `DOC/DOCX -> ODT`

## Поддерживаемые форматы

- Входные шаблоны:
  - `XLS`, `XLSX`
  - `DOC`, `DOCX`
  - `PDF`
- Выход:
  - исходный формат,
  - либо `ODS`/`ODT` через post-convert.

Важно: входные `ODS`/`ODT` шаблоны не поддерживаются.

## Quickstart

```java
ReportGeneratorService service = new ReportGeneratorServiceImpl();

TemplateInput input = new TemplateInput("sales-report.xlsx", null, templateBytes);

ReportData data = new ReportData(Map.of(
        "period", "2026-Q1",
        "rows", List.of(
                Map.of("name", "North", "amount", 1200.25),
                Map.of("name", "South", "amount", 900.00)
        ),
        TagConstants.ROWS_COLUMNS.getValue(), List.of("name", "amount")
));

GenerateOptions options = new GenerateOptions(
        MissingValuePolicy.EMPTY_AND_LOG,
        true,
        Locale.getDefault(),
        ZoneId.systemDefault(),
        false // rowsOnlyTableTokens
);
GeneratedReport report = service.generate(input, data, options);
```

## DOCX TableBuilder (declarative)

```java
TableBuilder schedule = TableBuilder.create()
        .row(TableBuilder.boldCell("Payment schedule", 4))
        .row(
                TableBuilder.boldCell("No"),
                TableBuilder.boldCell("Payment month"),
                TableBuilder.boldCell("Payment amount"),
                TableBuilder.boldCell("Remaining balance")
        )
        .row(
                TableBuilder.cell("1."),
                TableBuilder.cell("{{payment_date}}"),
                TableBuilder.cell("{{amount}}"),
                TableBuilder.cell("{{ost_osn_dolg}}")
        );

TemplateInput input = new TemplateInput("report.docx", null, docxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "TABLE_HERE", schedule,
        "payment_date", "2026-03",
        "amount", "250000",
        "ost_osn_dolg", "750000"
));
GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

## XLSX TableXlsxBuilder (declarative)

```java
TableXlsxBuilder table = TableXlsxBuilder.create()
        .row(TableXlsxBuilder.boldCell("Payment schedule", 4))
        .row(
                TableXlsxBuilder.cell("1."),
                TableXlsxBuilder.cell("{{payment_date}}"),
                TableXlsxBuilder.cell("{{amount}}"),
                TableXlsxBuilder.cell("{{balance}}")
        );

TemplateInput input = new TemplateInput("report.xlsx", null, xlsxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "rows", table,
        "payment_date", "2026-03",
        "amount", 250000,
        "balance", 750000
));
GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

## Частые ошибки

- `TABLE_TOKEN_INLINE_IGNORED` / `TABLE_TOKEN_INLINE_TEXT_DROPPED`:
  - таблица вставляется только когда контейнер содержит placeholder-токен;
  - избегайте смешивания таблицы с произвольным текстом в одной ячейке/абзаце.
- `MISSING_TOKEN`:
  - токен есть в шаблоне, но отсутствует в `ReportData.templateTokens()`.
- `FORMULA_TOKEN_SKIPPED`:
  - токен найден внутри формулы, формула не переписывается намеренно.
- `TABLE_TOKEN_RECURSIVE`:
  - table-вставки не стабилизировались за защитный лимит проходов (`MAX_TABLE_PASSES`).
- `TABLE_TOKEN_INVALID` для `TableBuilder`:
  - декларативная таблица пустая или содержит некорректные строки/colspan.

## Документация

Подробная архитектура, алгоритмы, rationale по всем `WorkbookProcessor`, troubleshooting и расширенные примеры:

- [docs.md](docs.md)

Локальные ручные интеграционные сценарии перенесены в
`src/test/java/io/github/ogbozoyan/integration/ReportGeneratorManualIntegrationTest.java`.
