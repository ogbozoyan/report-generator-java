# Report Generator Docs

## 1. Что делает библиотека

`report-generator` генерирует отчёты из шаблонов по токенам `{{TOKEN}}`.

- scalar token: `{{name}}` -> подстановка значения;
- table token: `{{TABLE_TOKEN}}` -> если значение токена это `List<Map<String,Object>>`, вставляется таблица
  `header + rows`.

Legacy DSL (`[[TABLE_START...]]`, `[[COL_START...]]`) не поддерживается и должен падать fail-fast через
`TemplateSyntaxException`.

---

## 2. Контракт данных

### `TemplateInput`

- `fileName`: имя файла шаблона или желаемого выходного файла;
- `contentType`: optional MIME;
- `bytes`: байты исходного шаблона.

### `ReportData`

- `templateTokens`: единая карта токенов `Map<String,Object>`.
- Таблица передаётся как значение токена типа `List<Map<String,Object>>`.

Пример:

```java
ReportData data = new ReportData(
        Map.of(
                "period", "2026-Q1",
                "rows", List.of(
                        Map.of("name", "North", "amount", 1200.25),
                        Map.of("name", "South", "amount", 900.00)
                )
        )
);
```

---

## 3. Форматы

### Входные шаблоны (processing source)

- `XLS/XLSX` -> `PoiWorkbookProcessor`
- `DOC` -> `DocDocumentProcessor`
- `DOCX` -> `DocxDocumentProcessor`
- `PDF` -> `PdfDocumentProcessor`

### Выходные форматы

- По умолчанию формат выхода = формат обработанного шаблона.
- Поддерживаемые post-convert сценарии:
  - `XLS/XLSX -> ODS`
  - `DOC/DOCX -> ODT`

### Важно

- `ODS/ODT` как входные шаблоны **не поддерживаются**.
- Если вход `ODS/ODT`, сервис бросает `UnsupportedTemplateFormatException` с текстом миграции на `XLS/XLSX` или
  `DOC/DOCX`.

---

## 4. Архитектура и классы

### Service layer

- `com.template.reportgenerator.service.ReportGeneratorService`
  - публичный контракт `generate(template, data, options)`.
- `com.template.reportgenerator.service.ReportGeneratorServiceImpl`
  - оркестрация пайплайна:
    1) определение фактического входного формата (`TemplateFormatDetector.detectFormat`);
    2) определение желаемого выходного формата (`TemplateFormatDetector.detectRequestedOutputFormat`);
    3) валидация допустимой конвертации;
    4) применение токенов в выбранном процессоре;
    5) recalculate formulas (для spreadsheet);
    6) optional post-convert (`DocumentFormatConverter`);
    7) финальная сериализация (`ReportSerializer`).

### Processors

- `com.template.reportgenerator.processor.PoiWorkbookProcessor`
  - scalar replacement;
  - table insertion от маркерной ячейки;
  - сохранение baseline style;
  - auto-width колонок таблицы.
- `com.template.reportgenerator.processor.DocxDocumentProcessor`
  - scalar replacement;
  - table token -> `XWPFTable`.
- `com.template.reportgenerator.processor.DocDocumentProcessor`
  - basic text-grid table (`\t`/`\r`).
- `com.template.reportgenerator.processor.PdfDocumentProcessor`
  - text reconstruction + text-grid table.

### Utilities

- `TemplateFormatDetector`:
  - детект формата по magic bytes + extension/content-type;
  - OLE2 (`.doc`/`.xls`) различается по контейнерным entry (`WordDocument` vs `Workbook/Book`).
- `DocumentFormatConverter`:
  - контракт post-convert.
- `LibreOfficeDocumentFormatConverter`:
  - реализация через `soffice/libreoffice --headless --convert-to ...`.
- `ReportSerializer`:
  - нормализация `fileName` и `contentType` под финальный формат.
- `TokenResolver`, `WarningCollector`, `TemplateScanner`, `TemplateValidator`, `ValueWriter`:
  - резолв токенов, warnings, валидация синтаксиса и запись typed values.

---

## 5. Что проверяют тесты

- `src/test/java/com/template/reportgenerator/ReportGeneratorServiceImplTest.java`
  - scalar/table обработка для `XLS/XLSX/DOC/DOCX/PDF`;
  - авторасширение колонок и сохранение baseline style;
  - post-convert маршрутизация `XLSX->ODS` и `DOCX->ODT`;
  - запрет входных `ODS/ODT`;
  - missing token warnings.

- `src/test/java/com/template/reportgenerator/ReportGeneratorFormattingGoldenTest.java`
  - регрессии форматирования для таблиц в `XLSX` (styles/fonts/width/merged).

- `src/test/java/com/template/reportgenerator/util/TemplateFormatDetectorTest.java`
  - детект extension/content-type/magic;
  - различение OLE2 `DOC` и `XLS` по container entries;
  - детект requested output format.

---

## 6. Примеры шаблонов и usage

### 6.1 XLSX шаблон с таблицей

В шаблоне (например `Book1.xlsx`) в ячейке:

```text
{{rows}}
```

При `rows = List<Map<...>>` будет вставлено:

- строка header;
- строки данных;
- нижний контент смещается вниз.

Если порядок колонок нужно зафиксировать явно (особенно когда строки собираются через `Map.of(...)`),
добавьте токен `rows__columns`:

```java
ReportData data = new ReportData(Map.of(
        "rows", rows,
        "rows__columns", List.of("name", "amount", "жопа", "слона")
));
```

### 6.2 Базовый usage (без конвертации)

```java
ReportGeneratorService service = new ReportGeneratorServiceImpl();

TemplateInput input = new TemplateInput("report.xlsx", null, xlsxTemplateBytes);
ReportData data = new ReportData(Map.of(
        "name", "Alice",
        "rows", List.of(
                Map.of("name", "North", "amount", 1200.25),
                Map.of("name", "South", "amount", 900.00)
        )
));

GeneratedReport result = service.generate(input, data, GenerateOptions.defaults());
```

### 6.3 Выгрузка в ODS (post-convert)

Входной шаблон остаётся `XLSX`, но выход запрашиваем как `.ods`:

```java
TemplateInput input = new TemplateInput("report.ods", null, xlsxTemplateBytes);
GeneratedReport result = service.generate(input, data, GenerateOptions.defaults());
```

### 6.4 Выгрузка в ODT (post-convert)

Входной шаблон `DOC/DOCX`, выход `ODT`:

```java
TemplateInput input = new TemplateInput(
        "report.docx",
        "application/vnd.oasis.opendocument.text",
        docxTemplateBytes
);
GeneratedReport result = service.generate(input, data, GenerateOptions.defaults());
```

---

## 7. Ограничения

- table token вставляется как таблица только когда токен является exact-placeholder контейнера;
- `.doc` поддерживается в basic text-table режиме;
- для `ODS/ODT` нужен установленный `soffice/libreoffice` в `PATH` (только для post-convert);
- входные `ODS/ODT` шаблоны не поддерживаются.
