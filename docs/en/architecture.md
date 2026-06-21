> **Language:** **English** · [Русский](../ru/architecture.md)
>
> [Overview](README.md) · [Architecture & internals](architecture.md) · [Publishing to Maven Central](publishing.md)

# Report Generator: internal documentation

## 1. Purpose of this document

This document describes the internals of the library: the processing pipeline, the algorithms, the trade-offs and the
reasons behind the architectural decisions.

It is aimed at developers who maintain or extend the codebase.

## 2. Generation pipeline

Class: `io.github.ogbozoyan.service.ReportGeneratorServiceImpl`

The `generate(...)` flow:

1. Validate input arguments.
2. Resolve `GenerateOptions` (apply defaults).
3. Detect the actual input format (`TemplateFormatDetector.detectFormat`).
4. Detect the requested output format (`TemplateFormatDetector.detectRequestedOutputFormat`).
5. Check that the conversion route is allowed.
6. Create the format processor (`WorkbookProcessor` implementation).
7. Apply tokens (`applyTemplateTokens`).
8. Recalculate formulas (only where supported).
9. Serialize the processed document.
10. Optional post-convert (`DocumentFormatConverter`) for `ODS/ODT` export.
11. Normalize output metadata and collect warnings (`ReportSerializer`).

Why:

- the pipeline is split into pure stages to localize responsibility and ease diagnostics;
- format logic is moved into processors, the service stays an orchestrator;
- post-convert is decoupled from the substitution stage to keep the tokenization algorithm simple.

## 3. Data contracts

### 3.1 `TemplateInput`

- `fileName`: name of the input template or expected output.
- `contentType`: optional MIME hint.
- `bytes`: template bytes.

### 3.2 `ReportData`

- `templateTokens`: a single data map `Map<String, Object>`.
- scalar token: any string/numeric/date value.
- table token (default mode): `List<Map<String, Object>>`.
- table token (rows-only mode): `List<Object[]>`.
- table token (declarative XLS/XLSX mode): `TableXlsxBuilder`.
- table token (declarative DOC/DOCX mode): `TableBuilder`.
- table token (DOCX template-row mode): `RowBuilder`.
- optional column order: `TOKEN__columns` (or `TOKEN_columns`, `TOKEN.columns`).

### 3.3 `GenerateOptions`

- `missingValuePolicy`: behavior when a token is missing.
- `recalculateFormulas`: recalculate formulas for spreadsheets.
- `rowsOnlyTableTokens`: global rows-only insertion mode for `XLS/XLSX` table tokens
  (no header row, a `List<Object[]>` is expected).
- `locale`, `zoneId`: localization and time zone for writing dates.

## 4. Module and responsibility map

### 4.1 `service/*`

- `ReportGeneratorService`: the public generation API.
- `ReportGeneratorServiceImpl`: orchestration and routing by format.

### 4.2 `processor/*`

- `WorkbookProcessor`: a single lifecycle contract for format handlers.
- `PoiWorkbookProcessor`: `XLS/XLSX` tables, typed value writing, auto-width, formulas.
- `DocxDocumentProcessor`: traversal over the body/table/cell tree, table insertion into the correct container.
- `DocDocumentProcessor`: basic text-table in `.doc` (including declarative `TableBuilder` as a text-grid fallback).
- `PdfDocumentProcessor`: text reconstruction and ASCII-grid tables.

### 4.3 `util/*`

- `TemplateFormatDetector`: format detection by magic bytes/extension/MIME.
- `TokenResolver`: token lookup/resolution and table typing.
- `WarningCollector`: accumulation of non-fatal warnings.
- `ReportSerializer`: fileName/contentType/warnings of the final result.
- `LibreOfficeDocumentFormatConverter`: post-convert to `ODS/ODT`.
- `TemplateScanner`, `TemplateValidator`: scan/validation helpers for legacy-DSL scenarios.

### 4.4 `contract/*` and `exception/*`

- contract types for passing data between layers;
- explicit exception types for reading/format/syntax/structure/binding.

## 5. Algorithms and why this approach was chosen

### 5.1 `WorkbookProcessor` (single contract and lifecycle)

Contract:

- `scan()`;
- `applyTemplateTokens(...)`;
- `recalculateFormulas(...)` (default no-op);
- `serialize()`;
- `close()`.

Why:

- a single interface lets the service stay independent of format details;
- the `default` for `recalculateFormulas` does not force non-spreadsheet processors to implement irrelevant logic;
- `AutoCloseable` makes resource discipline uniform across all implementations.

### 5.2 `PoiWorkbookProcessor`

Key algorithms:

- sparse traversal: iterate only over physically existing rows/cells.
- anchor-first strategy: first collect table anchors, then insert.
- reverse apply: insert tables bottom-up over the sheet.
- dual table modes:
  - default: header + data;
  - rows-only (`GenerateOptions.rowsOnlyTableTokens=true`): data rows only from `List<Object[]>`.
- declarative XLS/XLSX mode (`TableXlsxBuilder`):
  - an explicit row/cell model;
  - `colSpan` support via merged regions;
  - `bold` support while preserving the marker baseline style.
- multi-pass table expansion:
  - first run only table passes;
  - each pass: scan anchors -> reverse insert;
  - passes repeat while there are new anchors;
  - after stabilization a scalar pass runs.
- style baseline: reuse the marker cell style for header/data.
- auto-width: widths change only for inserted columns.
- formula policy: formulas with tokens are not rewritten, only a warning is emitted.

Why:

- sparse traversal avoids hangs on large sparse sheets;
- reverse apply makes multiple insertions deterministic under `shiftRows`;
- rows-only mode allows using a dedicated descriptor/mapping row without duplicating the header;
- multi-pass eliminates the loss of table tokens that only appear after a previous insertion;
- baseline style minimizes visual regressions of templates;
- local auto-width does not break the external sheet layout;
- skipping formula tokens is safer than risking corruption of the formula syntax.

### 5.3 `DocxDocumentProcessor`

Key algorithms:

- recursive traversal over `IBody`: document body -> table -> cell -> nested body;
- collection of `ParagraphTarget` in traversal order;
- table anchors are applied in reverse order;
- a table is inserted strictly into the paragraph container (`XWPFDocument` or `XWPFTableCell`);
- the placeholder paragraph is removed from the source container after insertion.
- the declarative `TableBuilder` payload is supported:
  - rows/cells are defined in code;
  - `colSpan` maps to `w:gridSpan`;
  - `bold` is applied at the run level inside the cell.
- the template-row `RowBuilder` payload is supported:
  - the token is placed at the start of a row of an existing table;
  - the template row is cloned for each provided row while preserving the table format;
  - row/cell formatting is preserved via cloning `CTRow`.

Why:

- DOCX often contains tokens inside existing tables, not only in body paragraphs;
- the correct insertion container eliminates the case where a table was created somewhere other than the placeholder;
- removing the placeholder paragraph prevents content duplication.

### 5.4 `DocDocumentProcessor`

Key algorithms:

- an exact paragraph placeholder is recognized in an `HWPF Range`;
- a table token is rendered as a text grid: header/rows with `\t` and `\r` separators;
- a declarative `TableBuilder` is rendered into the same text-grid format;
- scalar tokens are replaced via a bulk `range.replaceText(...)`.

Why:

- `.doc` (HWPF) is limited in safe structural editing;
- a text grid provides robust "basic" support sufficient for simple reports;
- the approach minimizes the risk of corrupting the binary structure of `.doc`.

Limitation:

- this is not a full Word table model, but a textual imitation of a table.

### 5.5 `PdfDocumentProcessor`

Key algorithms:

- the PDF is read as text (`PDFTextStripper`), then token replacement is performed;
- a table token is rendered as an ASCII grid;
- serialization builds a new PDF line by line with word-wrap and pagination.

Why:

- PDF does not support reliable in-place editing of text objects without complex geometric reconstruction;
- text reconstruction provides a predictable and robust result;
- the ASCII grid gives a portable representation of tables for the text stream.

Limitation:

- the output PDF layout is not a 1:1 copy of the source template.

## 6. Failure modes and troubleshooting

- `MISSING_TOKEN`:
  - the token is absent from `templateTokens`; check the keys and `missingValuePolicy`.
- `TABLE_TOKEN_INVALID`:
  - in default mode the token value is not a `List<Map<String,Object>>`;
  - in rows-only mode the token value is not a `List<Object[]>`;
  - in declarative XLS/XLSX mode an empty/invalid `TableXlsxBuilder` was passed;
  - in declarative mode an empty/invalid `TableBuilder` was passed;
  - in DOCX template-row mode the token is outside a table or the payload is invalid.
- `TABLE_TOKEN_EMPTY`:
  - the table was passed as an empty list.
- `TABLE_TOKEN_RECURSIVE`:
  - table-token insertions did not stabilize within the `MAX_TABLE_PASSES` limit;
  - check that table tokens do not reference each other cyclically.
- `GenerateOptions.rowsOnlyTableTokens=true`:
  - rows-only mode is enabled for `XLS/XLSX`: the marker row becomes the first data row;
  - the token payload must be a `List<Object[]>`.
- `TABLE_TOKEN_INLINE_IGNORED`:
  - a table token was found inline and was not inserted as a table.
- `TABLE_TOKEN_INLINE_TEXT_DROPPED`:
  - in single-token mode there was adjacent static text, which was dropped when the table was inserted.
- `FORMULA_TOKEN_SKIPPED`:
  - the token is in a formula cell; the formula was left unchanged.
- `Unsupported output conversion`:
  - only `XLS/XLSX -> ODS` and `DOC/DOCX -> ODT` are allowed.
- `UnsupportedTemplateFormatException` for an input `ODS/ODT`:
  - use an input `XLS/XLSX` or `DOC/DOCX`, then request `ODS/ODT` as the output.
- LibreOffice conversion errors:
  - check that `soffice`/`libreoffice` is available on the `PATH`.

## 7. Usage examples

### 7.1 XLSX with a table token

```java
ReportGeneratorService service = new ReportGeneratorServiceImpl();

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

GeneratedReport report = service.generate(input, data, options);
```

### 7.2 DOCX: a table token inside an existing table

Template condition:

- a cell of a DOCX table contains a separate paragraph with `{{inner_table}}`.

Example:

```java
TemplateInput input = new TemplateInput("DOC1.docx", null, docxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "inner_table", List.of(
                Map.of("kpi", "Revenue", "value", "125000"),
                Map.of("kpi", "Margin", "value", "24%")
        )
));

GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

### 7.3 XLSX -> ODS

```java
TemplateInput input = new TemplateInput("sales-report.ods", null, xlsxTemplateBytes);
GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

### 7.4 DOCX -> ODT

```java
TemplateInput input = new TemplateInput(
        "letter.odt",
        "application/vnd.oasis.opendocument.text",
        docxTemplateBytes
);
GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

### 7.5 DOCX declarative table (`TableBuilder`)

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
                TableBuilder.cell("{{balance}}")
        );

TemplateInput input = new TemplateInput("contract.docx", null, docxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "TABLE_HERE", schedule,
        "payment_date", "2026-03",
        "amount", "250000",
        "balance", "750000"
));
GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

### 7.6 XLSX declarative table (`TableXlsxBuilder`)

```java
TableXlsxBuilder schedule = TableXlsxBuilder.create()
        .row(TableXlsxBuilder.boldCell("Payment schedule", 4))
        .row(
                TableXlsxBuilder.cell("1."),
                TableXlsxBuilder.cell("{{payment_date}}"),
                TableXlsxBuilder.cell("{{amount}}"),
                TableXlsxBuilder.cell("{{balance}}")
        );

TemplateInput input = new TemplateInput("table.xlsx", null, xlsxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "rows", schedule,
        "payment_date", "2026-03",
        "amount", 250000,
        "balance", 750000
));
GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

### 7.7 DOCX template row clone (`RowBuilder`)

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

TemplateInput input = new TemplateInput("contract.docx", null, docxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "PAYMENT_ROWS", paymentRows
));
GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

Template condition:
- `{{PAYMENT_ROWS}}` is placed inside a row of an existing DOCX table (the row is used as the template row).

## 8. How decisions map to tests

Key test suites and what they confirm:

- `src/test/java/io/github/ogbozoyan/service/ReportGeneratorServiceImplTest.java`
  - the service pipeline;
  - table insertion in `XLS/XLSX` and non-spreadsheet formats;
  - column order;
  - inline/exact-placeholder behavior;
  - supported post-convert routes.

- `src/test/java/io/github/ogbozoyan/service/ReportGeneratorFormattingGoldenTest.java`
  - regression check of spreadsheet formatting during table insertion.

- `src/test/java/io/github/ogbozoyan/integration/ReportGeneratorManualIntegrationTest.java`
  - manual integration scenarios moved out of `main` (the class is marked `@Disabled`).

- `src/test/java/io/github/ogbozoyan/util/TemplateFormatDetectorTest.java`
  - format detection by magic bytes/content-type/extension;
  - distinguishing OLE2 (`DOC` vs `XLS`);
  - routing of the requested output format.

- `src/test/java/io/github/ogbozoyan/util/TemplateValidatorTest.java`
  - correctness of the scan/validation helper logic for block markers.

## 9. Why overview and architecture are separate

- The [Overview](README.md) answers "what is this" and "how to get started quickly".
- The [Architecture](architecture.md) answers "how it is implemented" and "why it is done this way".

This reduces duplication and eases documentation maintenance when algorithms change.
