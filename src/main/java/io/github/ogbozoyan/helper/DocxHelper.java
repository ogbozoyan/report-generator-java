package io.github.ogbozoyan.helper;

import io.github.ogbozoyan.data.ParagraphTarget;
import io.github.ogbozoyan.exception.TemplateReadWriteException;
import lombok.experimental.Helper;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Helper
@Slf4j
public class DocxHelper extends CommonHelper {
    /**
     * Ensures row has at least requested number of cells.
     *
     * @param row   table row
     * @param count minimum cell count
     */
    public static void ensureCells(XWPFTableRow row, int count) {
        while (row.getTableCells().size() < count) {
            row.addNewTableCell();
        }
    }


    /**
     * Replaces cell paragraphs with single paragraph containing provided value.
     *
     * @param cell  destination cell
     * @param value text value
     */
    public static void setCellText(XWPFTableCell cell, String value) {
        int paragraphCount = cell.getParagraphs().size();
        for (int i = paragraphCount - 1; i >= 0; i--) {
            cell.removeParagraph(i);
        }
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(value == null ? "" : value);
    }


    /**
     * Replaces paragraph content with a single run.
     *
     * @param paragraph paragraph to rewrite
     * @param value     replacement text
     */
    public static void replaceParagraphText(XWPFParagraph paragraph, String value) {
        int runs = paragraph.getRuns().size();
        for (int i = runs - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
        XWPFRun run = paragraph.createRun();
        run.setText(value == null ? "" : value);
    }

    /**
     * Writes header and data rows into newly created DOCX table.
     *
     * @param table   destination table
     * @param columns ordered columns
     * @param rows    table payload rows
     */
    public static void writeTable(XWPFTable table, List<String> columns, List<Map<String, Object>> rows) {
        XWPFTableRow headerRow = getOrCreateFirstRow(table);
        ensureCells(headerRow, columns.size());
        for (int c = 0; c < columns.size(); c++) {
            setCellText(headerRow.getCell(c), columns.get(c));
        }

        for (Map<String, Object> row : rows) {
            XWPFTableRow dataRow = table.createRow();
            ensureCells(dataRow, columns.size());
            for (int c = 0; c < columns.size(); c++) {
                Object value = row.get(columns.get(c));
                setCellText(dataRow.getCell(c), value == null ? "" : String.valueOf(value));
            }
        }
    }

    /**
     * Returns existing first row or creates one for table header.
     *
     * @param table destination table
     * @return first row
     */
    public static XWPFTableRow getOrCreateFirstRow(XWPFTable table) {
        XWPFTableRow headerRow = table.getRow(0);
        if (headerRow != null) {
            return headerRow;
        }
        headerRow = table.insertNewTableRow(0);
        if (headerRow != null) {
            if (headerRow.getCell(0) == null) {
                headerRow.createCell();
            }
            return headerRow;
        }
        throw new TemplateReadWriteException("Failed to create header row for DOCX table insertion");
    }


    /**
     * Inserts table at paragraph cursor according to paragraph container type.
     *
     * @param paragraph anchor paragraph
     * @param cursor    insertion cursor
     * @return newly inserted table
     */
    public static XWPFTable insertTableAtParagraphCursor(XWPFParagraph paragraph, XmlCursor cursor) {
        IBody body = paragraph.getBody();
        if (body instanceof XWPFDocument doc) {
            return doc.insertNewTbl(cursor);
        }
        if (body instanceof XWPFTableCell cell) {
            return cell.insertNewTbl(cursor);
        }
        throw new TemplateReadWriteException(
            "Unsupported DOCX paragraph container for table insertion: " + body.getClass().getName()
        );
    }

    /**
     * Removes placeholder paragraph from its container after table insertion.
     *
     * @param paragraph placeholder paragraph
     */
    public static void removeParagraphFromContainer(XWPFParagraph paragraph) {
        IBody body = paragraph.getBody();
        if (body instanceof XWPFDocument doc) {
            int paragraphPos = doc.getPosOfParagraph(paragraph);
            if (paragraphPos >= 0) {
                doc.removeBodyElement(paragraphPos);
                return;
            }
        } else if (body instanceof XWPFTableCell cell) {
            int paragraphPos = cell.getParagraphs().indexOf(paragraph);
            if (paragraphPos >= 0) {
                cell.removeParagraph(paragraphPos);
                return;
            }
        }
        replaceParagraphText(paragraph, "");
    }


    /**
     * Collects every paragraph from document body and nested table cells.
     *
     * @return ordered paragraph targets
     */
    public static List<ParagraphTarget> collectParagraphTargets(XWPFDocument document) {
        List<ParagraphTarget> targets = new ArrayList<>();
        int[] order = {0};
        collectParagraphTargetsFromBody(document, "docx:body", order, targets);
        return targets;
    }


    /**
     * Recursively traverses body elements and records paragraph targets.
     *
     * @param body           body container (document or table cell)
     * @param locationPrefix diagnostic location prefix
     * @param order          mutable traversal counter
     * @param targets        output list
     */
    public static void collectParagraphTargetsFromBody(
        IBody body,
        String locationPrefix,
        int[] order,
        List<ParagraphTarget> targets
    ) {
        List<IBodyElement> bodyElements = body.getBodyElements();
        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement bodyElement = bodyElements.get(i);
            if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                targets.add(new ParagraphTarget(
                    paragraph,
                    paragraph.getText(),
                    locationPrefix + "/p#" + i,
                    order[0]++
                ));
                continue;
            }
            if (bodyElement.getElementType() == BodyElementType.TABLE) {
                XWPFTable table = (XWPFTable) bodyElement;
                collectParagraphTargetsFromTable(table, locationPrefix + "/tbl#" + i, order, targets);
            }
        }
    }

    /**
     * Traverses paragraphs inside table cells recursively.
     *
     * @param table          source table
     * @param locationPrefix diagnostic location prefix
     * @param order          mutable traversal counter
     * @param targets        output list
     */
    public static void collectParagraphTargetsFromTable(
        XWPFTable table,
        String locationPrefix,
        int[] order,
        List<ParagraphTarget> targets
    ) {
        List<XWPFTableRow> rows = table.getRows();
        for (int r = 0; r < rows.size(); r++) {
            XWPFTableRow row = rows.get(r);
            List<XWPFTableCell> cells = row.getTableCells();
            for (int c = 0; c < cells.size(); c++) {
                XWPFTableCell cell = cells.get(c);
                collectParagraphTargetsFromBody(
                    cell,
                    locationPrefix + "/r#" + r + "/c#" + c,
                    order,
                    targets
                );
            }
        }
    }

}
