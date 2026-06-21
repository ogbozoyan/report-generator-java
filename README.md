# Report Generator

**🌐 Language / Язык:** [English](docs/en/README.md) · [Русский](docs/ru/README.md)

A Java library that generates documents from templates with `{{TOKEN}}` markers — scalar values, tables and
declarative builders for `XLS/XLSX`, `DOC/DOCX` and `PDF`, with optional `ODS/ODT` post-convert.

Библиотека генерации документов по шаблонам с маркерами `{{TOKEN}}` — скалярные значения, таблицы и декларативные
билдеры для `XLS/XLSX`, `DOC/DOCX` и `PDF`, с опциональной выгрузкой в `ODS/ODT`.

## Documentation / Документация

| | English | Русский |
|---|---|---|
| Overview / Обзор | [docs/en/README.md](docs/en/README.md) | [docs/ru/README.md](docs/ru/README.md) |
| Architecture & internals / Архитектура | [docs/en/architecture.md](docs/en/architecture.md) | [docs/ru/architecture.md](docs/ru/architecture.md) |
| Publishing to Maven Central / Публикация | [docs/en/publishing.md](docs/en/publishing.md) | [docs/ru/publishing.md](docs/ru/publishing.md) |

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

GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

See the full overview for all builders, table modes and options:
[English](docs/en/README.md) · [Русский](docs/ru/README.md).

## License

[MIT](LICENSE)
