package com.template.reportgenerator.util;

import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;

/**
 * Writes Java values to format-specific cells while preserving type where possible.
 */
@UtilityClass
public class ValueWriter {

    public static void writePoiValue(Cell cell, Object value, ZoneId zoneId) {
        switch (value) {
            case null -> {
                cell.setBlank();
                return;
            }
            case Number number -> {
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(number.doubleValue());
                return;
            }
            case Boolean bool -> {
                cell.setCellType(CellType.BOOLEAN);
                cell.setCellValue(bool);
                return;
            }
            case Date date -> {
                cell.setCellValue(date);
                return;
            }
            case LocalDate localDate -> {
                Date date = Date.from(localDate.atStartOfDay(zoneId).toInstant());
                cell.setCellValue(date);
                return;
            }
            case LocalDateTime localDateTime -> {
                Date date = Date.from(localDateTime.atZone(zoneId).toInstant());
                cell.setCellValue(date);
                return;
            }
            case Instant instant -> {
                cell.setCellValue(Date.from(instant));
                return;
            }
            default -> {
            }
        }

        cell.setCellType(CellType.STRING);
        cell.setCellValue(String.valueOf(value));
    }

    public void writeOdsValue(OdfTableCell cell, Object value, ZoneId zoneId) {
        Calendar calendar = Calendar.getInstance();

        switch (value) {
            case null -> {
                cell.setStringValue("");
                return;
            }
            case Number number -> {
                cell.setDoubleValue(number.doubleValue());
                return;
            }
            case Boolean bool -> {
                cell.setBooleanValue(bool);
                return;
            }
            case Date date -> {
                calendar.setTime(date);
                cell.setDateValue(calendar);
                return;
            }
            case LocalDate localDate -> {
                calendar.setTime(Date.from(localDate.atStartOfDay(zoneId).toInstant()));
                cell.setDateValue(calendar);
                return;
            }
            case LocalDateTime localDateTime -> {
                calendar.setTime(Date.from(localDateTime.atZone(zoneId).toInstant()));
                cell.setDateValue(calendar);
                return;
            }
            case Instant instant -> {
                calendar.setTime(Date.from(instant));
                cell.setDateValue(calendar);
                return;
            }
            default -> {

            }
        }

        cell.setStringValue(String.valueOf(value));
    }
}
