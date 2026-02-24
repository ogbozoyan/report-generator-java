# Report Generator Docs

## 1. Что делает библиотека

`report-generator` генерирует отчёты из шаблонов по токенам вида `{{TOKEN}}` без REST-слоя.

Ключевая модель:

- scalar token: `{{name}}` -> подставляет скаляр;
- table token: `{{TABLE_TOKEN}}` -> если значение токена это `List<Map<String,Object>>`, вставляется таблица
  `Header + rows`.

Старый DSL блоков (`[[TABLE_START...]]`, `[[COL_START...]]`) удалён из пайплайна и теперь всегда падает fail-fast с
`TemplateSyntaxException`.

---

## 2. Контракт данных

### `ReportData`

- `scalars`: основной token-map (включая таблицы).
- `tables`, `columns`: legacy-поля, оставлены только для совместимости, в генерации больше не используются.

### Как передавать таблицу

Для токена `{{rows}}`:

```java
ReportData data = new ReportData(
    Map.of(
        "period", "2026-Q1",
        "rows", List.of(
            Map.of("name", "North", "amount", 1200.25),
            Map.of("name", "South", "amount", 900.00)
        )
    ),
    Map.of(),
    Map.of()
);
```

### Правила table token

- таблица вставляется только если токен является **единственным содержимым контейнера**:
    - spreadsheet: единственное содержимое ячейки;
    - docx/odt/doc/pdf: единственное содержимое абзаца/строки.
- если table token встроен inline в текст, вставка таблицы не делается, остаётся scalar-режим + warning
  `TABLE_TOKEN_INLINE_IGNORED`.
- порядок колонок: ключи первой строки + новые ключи из следующих строк в конец.

---

## 3. Поддержка форматов

| Формат          | Процессор               | Таблицы        | Примечание                                                |
|-----------------|-------------------------|----------------|-----------------------------------------------------------|
| `.xls`, `.xlsx` | `PoiWorkbookProcessor`  | Да             | Вставка с якорем в маркерной ячейке, auto-width через POI |
| `.ods`          | `OdsWorkbookProcessor`  | Да             | Аналогичная логика, auto-width через ODFDOM               |
| `.docx`         | `DocxDocumentProcessor` | Да             | Вставка `XWPFTable` в позицию placeholder-параграфа       |
| `.odt`          | `OdtDocumentProcessor`  | Да             | Вставка `OdfTable` в позицию placeholder-параграфа        |
| `.doc`          | `DocDocumentProcessor`  | Да (basic)     | Таблица как text-grid с `\t`/`\r` (базовая поддержка)     |
| `.pdf`          | `PdfDocumentProcessor`  | Да (text-grid) | Рендер текстовой таблицы в текущем PDF pipeline           |

---

## 4. Архитектура и зоны ответственности

## Service layer

- `com.template.reportgenerator.service.ReportGeneratorService`
    - публичный контракт `generate(template, data, options)`.
- `com.template.reportgenerator.service.ReportGeneratorServiceImpl`
    - оркестрация:
        1) detect format (`TemplateFormatDetector`);
        2) fail-fast legacy DSL (`LegacyDslDetector`);
        3) token apply в format-процессоре;
        4) formula recalc (где применимо);
        5) serialize (`ReportSerializer`).

## DTO / model

- `TemplateInput`, `GeneratedReport`, `GenerateOptions`, `GenerationWarning`, `TemplateFormat`, `ReportData`.

## Format processors

- `PoiWorkbookProcessor` (XLS/XLSX)
    - scalar token replace;
    - table token insert (`Header + rows`);
    - baseline style/height от маркерной ячейки;
    - сдвиг строк вниз при необходимости;
    - auto-width по контенту таблицы.
- `OdsWorkbookProcessor` (ODS)
    - тот же контракт table token + baseline style/height + auto-width.
- `DocxDocumentProcessor` (DOCX)
    - table token -> `XWPFTable` в позицию placeholder paragraph.
- `OdtDocumentProcessor` (ODT)
    - table token -> `OdfTable` в позицию placeholder paragraph.
- `DocDocumentProcessor` (DOC)
    - basic text-table для exact placeholder токена.
- `PdfDocumentProcessor` (PDF)
    - table token -> ASCII/text grid.

## Utilities

- `TemplateFormatDetector`: определение формата по extension/content-type/magic.
- `LegacyDslDetector`: обнаружение legacy DSL в шаблоне любого поддержанного формата.
- `TokenResolver`: резолв токенов + table helpers (`isTableValue`, `toTableRows`).
- `ValueWriter`: запись typed values в POI/ODF.
- `WarningCollector`: сбор предупреждений.
- `ReportSerializer`: нормализация имени/типа результата.

---

## 5. Какие тесты что проверяют

### Основной coverage

- `src/test/java/com/template/reportgenerator/ReportGeneratorServiceImplTest.java`
    - scalar генерация для xls/xlsx/ods;
    - table token для xlsx/ods/docx/odt/doc/pdf;
    - порядок колонок;
    - auto-width;
    - baseline style;
    - fail-fast для legacy DSL;
    - missing token warnings.

### Форматирование regression

- `src/test/java/com/template/reportgenerator/ReportGeneratorFormattingGoldenTest.java`
    - сохранение style/font/row height/column width;
    - сохранение merged-region поведения в xlsx при вставке таблицы.

### Detection / validator

- `src/test/java/com/template/reportgenerator/util/TemplateFormatDetectorTest.java`
    - детект формата, включая `.doc`.
- `src/test/java/com/template/reportgenerator/util/LegacyDslDetectorTest.java`
    - fail-fast на legacy DSL и миграционный текст ошибки.

---

## 6. Примеры шаблонов и usage

### Spreadsheet шаблон (`Book1.xlsx`)

Вместо старых блоков DSL используйте **один marker token**:

```text
Ячейка A1: {{TABLE_HERE}}
```

Если `TABLE_HERE` -> `List<Map<...>>`, движок вставит:

- `A1..`: header;
- ниже строки данных;
- контент ниже сдвинется вниз.

### Non-spreadsheet шаблон

В абзаце документа:

```text
{{TABLE_HERE}}
```

Тогда:

- DOCX/ODT: вставится настоящая таблица;
- DOC/PDF: вставится text-grid таблица.

### Runtime usage

```java
ReportGeneratorService service = new ReportGeneratorServiceImpl();

TemplateInput input = new TemplateInput("Book1.xlsx", null, templateBytes);
ReportData data = new ReportData(
    Map.of(
        "period", "2026-Q1",
        "TABLE_HERE", List.of(
            Map.of("name", "North", "amount", 1200.25),
            Map.of("name", "South", "amount", 900.00)
        )
    ),
    Map.of(),
    Map.of()
);

GeneratedReport report = service.generate(input, data, GenerateOptions.defaults());
```

### Локальные файлы из проекта

- Шаблон: `/Users/onbozoyan/Downloads/report-generator/Book1.xlsx`
- Сгенерированный пример: `/Users/onbozoyan/Downloads/report-generator/sales-report.xlsx`

---

## 7. Важные ограничения

- table token требует exact-placeholder контейнер;
- `.doc` поддерживается в базовом текстовом виде;
- PDF не является 1:1 редактором исходного layout, используется text reconstruction pipeline.

