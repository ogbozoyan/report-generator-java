package io.github.ogbozoyan.helper;

import lombok.experimental.Helper;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;

@Helper
@Slf4j
public class PoiHelper extends CommonHelper {

    /**
     * Applies baseline style to cell when style is available.
     *
     * @param cell          destination cell
     * @param baselineStyle style copied from marker cell
     */
    public static void applyBaselineStyle(Cell cell, CellStyle baselineStyle) {
        if (baselineStyle != null) {
            cell.setCellStyle(baselineStyle);
        }
    }

    /**
     * Builds one-based cell location for diagnostics.
     *
     * @param sheet    source sheet
     * @param rowIndex zero-based row index
     * @param colIndex zero-based column index
     * @return location string
     */
    public static String cellLocation(Sheet sheet, int rowIndex, int colIndex) {
        return sheet.getSheetName() + "!R" + (rowIndex + 1) + "C" + (colIndex + 1);
    }


    /**
     * Checks whether cell type participates in token replacement phase.
     *
     * @param cell source cell
     * @return {@code true} for string/blank cells
     */
    public static boolean isEligibleCellType(Cell cell) {
        return cell.getCellType() == CellType.STRING || cell.getCellType() == CellType.BLANK;
    }


    /**
     * Calculates desired POI column width from maximum content length.
     *
     * @param maxLength max string length in column
     * @return clamped width in 1/256th character units
     */
    public static int calculateDesiredWidth(int maxLength) {
        int width = (maxLength + 2) * 256;
        int min = 8 * 256;
        int max = 100 * 256;
        return Math.max(min, Math.min(width, max));
    }

    /**
     * Converts value to string length for width estimation.
     *
     * @param value source value
     * @return string length or {@code 0} for {@code null}
     */
    public static int stringifyLength(Object value) {
        return value == null ? 0 : String.valueOf(value).length();
    }

    /**
     * Builds final column order for table insertion.
     *
     * <p>Configured order is used as base when provided, then unseen keys are appended
     * in row encounter order.
     *
     * @param rows                  normalized table rows
     * @param configuredColumnOrder optional explicit column order
     * @return ordered column names
     */
    public static List<String> buildColumnOrder(List<Map<String, Object>> rows, List<String> configuredColumnOrder) {
        LinkedHashSet<String> ordered = new LinkedHashSet<>();
        if (configuredColumnOrder != null && !configuredColumnOrder.isEmpty()) {
            ordered.addAll(configuredColumnOrder);
        } else {
            ordered.addAll(rows.get(0).keySet());
        }
        for (Map<String, Object> row : rows) {
            ordered.addAll(row.keySet());
        }
        return List.copyOf(ordered);
    }

    /**
     * Converts rows-only token payload into internal map-row representation.
     *
     * <p>Expected payload shape is {@code List<Object[]>}. Each array item represents
     * a single row where field order maps to columns by index.
     *
     * <p>Column naming strategy:
     * <ul>
     *     <li>configured column names are used first by index ({@code TOKEN__columns}, etc.),</li>
     *     <li>remaining unnamed fields use synthetic keys {@code __col{index}}.</li>
     * </ul>
     *
     * @param value                 raw token value
     * @param configuredColumnOrder optional configured column names
     * @return normalized rows or {@code null} for unsupported structure
     */
    public static List<Map<String, Object>> toRowsOnlyTableRows(Object value, List<String> configuredColumnOrder) {
        if (!(value instanceof List<?> list)) {
            return null;
        }
        if (list.isEmpty()) {
            return List.of();
        }

        List<Map<String, Object>> rows = new ArrayList<>(list.size());
        for (int rowIndex = 0; rowIndex < list.size(); rowIndex++) {
            Object rawRow = list.get(rowIndex);
            if (!(rawRow instanceof Object[] rowArray)) {
                return null;
            }

            LinkedHashMap<String, Object> normalizedRow = new LinkedHashMap<>();
            for (int columnIndex = 0; columnIndex < rowArray.length; columnIndex++) {
                String columnName = resolveRowsOnlyColumnName(configuredColumnOrder, columnIndex);
                normalizedRow.put(columnName, rowArray[columnIndex]);
            }
            rows.add(normalizedRow);
        }
        return rows;
    }

    /**
     * Resolves internal column key for rows-only row arrays.
     *
     * @param configuredColumnOrder optional configured column names
     * @param columnIndex           zero-based column index
     * @return configured name when present, otherwise synthetic key
     */
    public static String resolveRowsOnlyColumnName(List<String> configuredColumnOrder, int columnIndex) {
        if (columnIndex < configuredColumnOrder.size()) {
            return configuredColumnOrder.get(columnIndex);
        }
        return "__col" + columnIndex;
    }

    /**
     * Resolves optional explicit column order for a table token.
     *
     * <p>Supported keys:
     * <ul>
     *     <li>{@code token__columns}</li>
     *     <li>{@code token_columns}</li>
     *     <li>{@code token.columns}</li>
     * </ul>
     *
     * <p>Example:
     * <pre>{@code
     * Map<String, Object> tokens = Map.of(
     *     "rows", rows,
     *     TagConstants.ROWS_COLUMNS.getValue(), List.of("name", "amount", "region")
     * );
     * }</pre>
     */
    public static List<String> resolveConfiguredColumnOrder(Map<String, Object> context, String tableToken) {
        Object raw = context.get(tableToken + "__columns");
        if (raw == null) {
            raw = context.get(tableToken + "_columns");
        }
        if (raw == null) {
            raw = context.get(tableToken + ".columns");
        }
        if (raw instanceof List<?> list) {
            List<String> columns = list.stream()
                .filter(Objects::nonNull)
                .map(String::valueOf)
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .toList();
            if (!columns.isEmpty()) {
                return columns;
            }
        }
        return List.of();
    }

    /**
     * Writes Java value to cell preserving semantic type where possible.
     *
     * <p>Mapping examples:
     * <ul>
     *     <li>{@link Number} -> numeric cell,</li>
     *     <li>{@link Boolean} -> boolean cell,</li>
     *     <li>{@link LocalDate}/{@link LocalDateTime}/{@link Instant}/{@link Date} -> date cell value,</li>
     *     <li>fallback -> string.</li>
     * </ul>
     */
    public static void writeValueToCell(Cell cell, Object value, ZoneId zoneId) {
        if (value == null) {
            cell.setBlank();
            return;
        }
        if (value instanceof Number number) {
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(number.doubleValue());
        } else if (value instanceof Boolean bool) {
            cell.setCellType(CellType.BOOLEAN);
            cell.setCellValue(bool);
        } else if (value instanceof Date date) {
            cell.setCellValue(date);
        } else if (value instanceof LocalDate localDate) {
            Date date = Date.from(localDate.atStartOfDay(zoneId).toInstant());
            cell.setCellValue(date);
        } else if (value instanceof LocalDateTime localDateTime) {
            Date date = Date.from(localDateTime.atZone(zoneId).toInstant());
            cell.setCellValue(date);
        } else if (value instanceof Instant instant) {
            cell.setCellValue(Date.from(instant));
        } else {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(String.valueOf(value));
        }
    }

    /**
     * Returns row by index or creates missing row.
     *
     * @param sheet    target sheet
     * @param rowIndex zero-based row index
     * @return existing or new row
     */
    public static Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            log.trace("getOrCreateRow() - create: rowIndex={}", rowIndex);
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    /**
     * Returns cell by index or creates missing cell.
     *
     * @param row      target row
     * @param colIndex zero-based column index
     * @return existing or new cell
     */
    public static Cell getOrCreateCell(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            log.trace("getOrCreateCell() - create: columnIndex={}", colIndex);
            cell = row.createCell(colIndex);
        }
        return cell;
    }


    /**
     * Resolves last cell number for logging current sheet geometry.
     *
     * @param sheet source sheet
     * @return last cell number in last row, or {@code 0} when unavailable
     */
    public static short resolveSheetLastCellNum(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return 0;
        }
        Row row = sheet.getRow(lastRowNum);
        return row == null ? 0 : row.getLastCellNum();
    }
}
