package io.github.ogbozoyan.util;


import io.github.ogbozoyan.contract.BlockMarker;
import io.github.ogbozoyan.contract.BlockType;
import io.github.ogbozoyan.contract.CellPosition;
import io.github.ogbozoyan.contract.TemplateScanResult;
import io.github.ogbozoyan.contract.TokenOccurrence;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.odftoolkit.odfdom.doc.OdfDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Scans spreadsheet templates and collects DSL markers and scalar tokens.
 */
@UtilityClass
@Slf4j
public class TemplateScanner {

    private static final Pattern BLOCK_PATTERN = Pattern.compile("^\\[\\[\\s*(TABLE_START|TABLE_END|COL_START|COL_END)\\s*:\\s*([a-zA-Z0-9_.-]+)\\s*]]$");
    private static final Pattern TOKEN_PATTERN = TokenResolver.TOKEN_PATTERN;

    /**
     * Scans POI workbook for legacy block markers and scalar token occurrences.
     *
     * @param workbook workbook to scan
     * @return scan result
     */
    public static TemplateScanResult scanPoi(Workbook workbook) {
        log.trace("scanPoi() - start: sheetCount={}", workbook == null ? null : workbook.getNumberOfSheets());
        List<BlockMarker> markers = new ArrayList<>();
        List<TokenOccurrence> tokens = new ArrayList<>();

        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String text = readCellAsText(cell);
                    if (text == null || text.isBlank()) {
                        continue;
                    }
                    CellPosition position = new CellPosition(s, sheet.getSheetName(), cell.getRowIndex(), cell.getColumnIndex());
                    collectFromText(text, position, markers, tokens);
                }
            }
        }

        TemplateScanResult result = new TemplateScanResult(markers, tokens);
        log.trace("scanPoi() - end: markers={}, tokens={}", result.markers().size(), result.scalarTokens().size());
        return result;
    }

    /**
     * Scans ODS document for markers and scalar tokens.
     *
     * <p>Retained for compatibility with legacy ODS scanner path.
     *
     * @param document ODS document
     * @return scan result
     */
    public static TemplateScanResult scanOds(OdfDocument document) {
        log.trace("scanOds() - start: tableCount={}", document == null ? null : document.getTableList(false).size());
        List<BlockMarker> markers = new ArrayList<>();
        List<TokenOccurrence> tokens = new ArrayList<>();

        List<OdfTable> tables = document.getTableList(false);
        for (int t = 0; t < tables.size(); t++) {
            OdfTable table = tables.get(t);
            String sheetName = table.getTableName() == null ? ("Sheet" + t) : table.getTableName();

            int maxCols = detectUsedColumns(table);
            int maxRows = detectUsedRows(table, maxCols);

            for (int row = 0; row < maxRows; row++) {
                for (int col = 0; col < maxCols; col++) {
                    String text = table.getCellByPosition(col, row).getStringValue();
                    if (text == null || text.isBlank()) {
                        continue;
                    }
                    CellPosition position = new CellPosition(t, sheetName, row, col);
                    collectFromText(text, position, markers, tokens);
                }
            }
        }

        TemplateScanResult result = new TemplateScanResult(markers, tokens);
        log.trace("scanOds() - end: markers={}, tokens={}", result.markers().size(), result.scalarTokens().size());
        return result;
    }

    /**
     * Heuristically detects used ODS column range to avoid full-grid traversal.
     *
     * @param table ODS table
     * @return detected column bound
     */
    private int detectUsedColumns(OdfTable table) {
        int probeLimit = Math.min(table.getColumnCount(), 1024);
        int maxDetected = 0;
        int emptyStreak = 0;

        for (int col = 0; col < probeLimit; col++) {
            String value = table.getCellByPosition(col, 0).getStringValue();
            if (value != null && !value.isBlank()) {
                maxDetected = col + 1;
                emptyStreak = 0;
            } else {
                emptyStreak++;
            }

            if (emptyStreak >= 16 && maxDetected > 0) {
                break;
            }
        }

        return Math.max(maxDetected, Math.min(table.getColumnCount(), 128));
    }

    /**
     * Heuristically detects used ODS row range to avoid full-grid traversal.
     *
     * @param table   ODS table
     * @param maxCols max columns to probe
     * @return detected row bound
     */
    private int detectUsedRows(OdfTable table, int maxCols) {
        int probeLimit = Math.min(table.getRowCount(), 20000);
        int maxDetected = 0;
        int emptyStreak = 0;

        for (int row = 0; row < probeLimit; row++) {
            boolean hasAnyCellData = false;
            for (int col = 0; col < maxCols; col++) {
                String value = table.getCellByPosition(col, row).getStringValue();
                if (value != null && !value.isBlank()) {
                    hasAnyCellData = true;
                    break;
                }
            }

            if (hasAnyCellData) {
                maxDetected = row + 1;
                emptyStreak = 0;
            } else {
                emptyStreak++;
            }

            if (emptyStreak >= 32 && maxDetected > 0) {
                break;
            }
        }

        return Math.max(maxDetected, Math.min(table.getRowCount(), 512));
    }

    /**
     * Collects markers/tokens from single text node.
     *
     * @param text     source text
     * @param position token position
     * @param markers  destination marker list
     * @param tokens   destination token list
     */
    private void collectFromText(
        String text,
        CellPosition position,
        List<BlockMarker> markers,
        List<TokenOccurrence> tokens
    ) {
        String trimmed = text.trim();
        Matcher markerMatcher = BLOCK_PATTERN.matcher(trimmed);
        if (markerMatcher.matches()) {
            String kind = markerMatcher.group(1).toUpperCase(Locale.ROOT);
            String key = markerMatcher.group(2);
            switch (kind) {
                case "TABLE_START" -> markers.add(new BlockMarker(BlockType.TABLE, "START", key, position));
                case "TABLE_END" -> markers.add(new BlockMarker(BlockType.TABLE, "END", key, position));
                case "COL_START" -> markers.add(new BlockMarker(BlockType.COL, "START", key, position));
                case "COL_END" -> markers.add(new BlockMarker(BlockType.COL, "END", key, position));
                default -> {
                    // no-op
                }
            }
            return;
        }

        Matcher tokenMatcher = TOKEN_PATTERN.matcher(text);
        while (tokenMatcher.find()) {
            String token = tokenMatcher.group(1);
            if ("index".equals(token) || token.startsWith("item.")) {
                continue;
            }
            tokens.add(new TokenOccurrence(token, position));
        }
    }

    /**
     * Reads cell content as text for scanner purposes.
     *
     * @param cell source cell
     * @return string content or {@code null}
     */
    private String readCellAsText(Cell cell) {
        if (cell == null) {
            return null;
        }

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }

        if (cell.getCellType() == CellType.FORMULA) {
            return cell.getCellFormula();
        }

        return null;
    }
}
