> **Language:** **English** · [Русский](../ru/README.md)
>
> [Overview](README.md) · [Architecture & internals](architecture.md) · [Publishing to Maven Central](publishing.md)

# Report Generator

`report-generator` is a library that generates documents from templates with `{{TOKEN}}` markers.

It is designed to be used as a service layer (no REST): you pass the template bytes, the token data and the generation
options, and you get back a ready file and a list of warnings.

## Features

- scalar tokens: substituting values into `{{TOKEN}}`;
- table tokens: `{{TABLE_TOKEN}}` for a value of type `List<Map<String, Object>>`;
- declarative table tokens for `DOC/DOCX` via `TableBuilder`
  (create a table from a placeholder without a pre-inserted table in the template);
- template-row row tokens for `DOCX` via `RowBuilder`
  (clone a row of an existing table while preserving its style);
- declarative table tokens for `XLS/XLSX` via `TableXlsxBuilder`
  (supports `colSpan` and `bold`, inserted at a placeholder);
- rows-only table mode for `XLS/XLSX` via `GenerateOptions.rowsOnlyTableTokens=true`
  and a value of type `List<Object[]>` (insertion without a header row);
- multi-pass processing of table tokens in `XLS/XLSX`: tokens that appear after an insertion
  (for example `{{TABLE_PART_2}}` inside inserted rows) are processed on the next pass;
- a unified data model for spreadsheet and non-spreadsheet formats;
- missing-token policy: `EMPTY_AND_LOG`, `LEAVE_TOKEN`, `FAIL_FAST`;
- formula recalculation for `XLS/XLSX`;
- post-convert export:
  - `XLS/XLSX -> ODS`
  - `DOC/DOCX -> ODT`

## Supported formats

- Input templates:
  - `XLS`, `XLSX`
  - `DOC`, `DOCX`
  - `PDF`
- Output:
  - the source format,
  - or `ODS`/`ODT` via post-convert.

Note: input `ODS`/`ODT` templates are not supported.

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

## DOCX RowBuilder (template row clone)

```java
RowBuilder paymentRows = RowBuilder.create()
        .row(
                RowBuilder.cell("1"),
                RowBuilder.cell("2026-03"),
                RowBuilder.cell("250000"),
                RowBuilder.cell("750000")
        )
        .row(
                RowBuilder.cell("2"),
                RowBuilder.cell("2026-04"),
                RowBuilder.cell("250000"),
                RowBuilder.cell("500000")
        );

// {{PAYMENT_ROWS}} must be placed in a row of a DOCX template table.
ReportData data = new ReportData(Map.of(
        "PAYMENT_ROWS", paymentRows
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

## Common errors

- `TABLE_TOKEN_INLINE_IGNORED` / `TABLE_TOKEN_INLINE_TEXT_DROPPED`:
  - a table is inserted only when the container holds a placeholder token;
  - avoid mixing a table with arbitrary text in the same cell/paragraph.
- `MISSING_TOKEN`:
  - the token is in the template but absent from `ReportData.templateTokens()`.
- `FORMULA_TOKEN_SKIPPED`:
  - the token was found inside a formula; the formula is intentionally not rewritten.
- `TABLE_TOKEN_RECURSIVE`:
  - table insertions did not stabilize within the guard pass limit (`MAX_TABLE_PASSES`).
- `TABLE_TOKEN_INVALID` for `TableBuilder`:
  - the declarative table is empty or contains invalid rows/colspan.

## Documentation

Detailed architecture, algorithms, rationale for every `WorkbookProcessor`, troubleshooting and extended examples:

- [Architecture & internals](architecture.md)

Local manual integration scenarios live in
`src/test/java/io/github/ogbozoyan/integration/ReportGeneratorManualIntegrationTest.java`.
