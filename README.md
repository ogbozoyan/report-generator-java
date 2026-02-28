# Report Generator

`report-generator` — библиотека генерации документов по шаблонам с маркерами `{{TOKEN}}`.

Библиотека предназначена для использования как io.github.ogbozoyan.service-layer (без REST): вы передаёте байты шаблона,
данные токенов и
опции генерации, на выходе получаете готовый файл и список предупреждений.

## Что умеет

- scalar tokens: подстановка значений в `{{TOKEN}}`;
- table tokens: `{{TABLE_TOKEN}}` для значения типа `List<Map<String, Object>>`;
- rows-only table mode для `XLS/XLSX` через `GenerateOptions.rowsOnlyTableTokens=true`
  и значения типа `List<Object[]>` (вставка без header-строки);
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
ReportGeneratorService io.github.ogbozoyan.service =new

ReportGeneratorServiceImpl();

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
GeneratedReport report = io.github.ogbozoyan.service.generate(input, data, options);
```

## Частые ошибки

- `TABLE_TOKEN_INLINE_IGNORED` / `TABLE_TOKEN_INLINE_TEXT_DROPPED`:
  - таблица вставляется только когда контейнер содержит placeholder-токен;
  - избегайте смешивания таблицы с произвольным текстом в одной ячейке/абзаце.
- `MISSING_TOKEN`:
  - токен есть в шаблоне, но отсутствует в `ReportData.templateTokens()`.
- `FORMULA_TOKEN_SKIPPED`:
  - токен найден внутри формулы, формула не переписывается намеренно.

## Документация

Подробная архитектура, алгоритмы, rationale по всем `WorkbookProcessor`, troubleshooting и расширенные примеры:

- [docs.md](docs.md)
