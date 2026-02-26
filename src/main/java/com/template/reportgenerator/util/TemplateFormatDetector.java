package com.template.reportgenerator.util;

import com.template.reportgenerator.contract.TemplateFormat;
import com.template.reportgenerator.contract.TemplateInput;
import com.template.reportgenerator.exception.UnsupportedTemplateFormatException;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.Locale;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * Detects the template format from file extension, content type, and binary signature.
 */
@UtilityClass
@Slf4j
public class TemplateFormatDetector {

    private static final int ZIP_SIGNATURE_0 = 0x50;
    private static final int ZIP_SIGNATURE_1 = 0x4B;
    private static final int ZIP_SIGNATURE_2 = 0x03;
    private static final int ZIP_SIGNATURE_3 = 0x04;

    private static final int OLE2_SIGNATURE_0 = 0xD0;
    private static final int OLE2_SIGNATURE_1 = 0xCF;
    private static final int OLE2_SIGNATURE_2 = 0x11;
    private static final int OLE2_SIGNATURE_3 = 0xE0;

    /**
     * Detects requested output format from user-facing hints only.
     *
     * <p>This method intentionally does not inspect bytes and is used to resolve
     * output format preference from file name/content type.
     *
     * @param input template input with fileName/contentType hints
     * @return requested output format or {@code null} when no explicit hint is present
     */
    public static TemplateFormat detectRequestedOutputFormat(TemplateInput input) {
        log.info("detectRequestedOutputFormat() - start: fileName={}, contentType={}",
            input == null ? null : input.fileName(),
            input == null ? null : input.contentType());
        TemplateFormat result = detectByContentType(input == null ? null : input.contentType());
        if (result == null) {
            result = detectByFileName(input == null ? null : input.fileName());
        }
        log.info("detectRequestedOutputFormat() - end: format={}", result);
        return result;
    }

    /**
     * Detects actual source format using magic bytes, container signatures and fallback hints.
     *
     * @param input template input with bytes and optional hints
     * @return detected source format
     * @throws UnsupportedTemplateFormatException when format cannot be detected
     */
    public static TemplateFormat detectFormat(TemplateInput input) {
        log.info("detectFormat() - start: fileName={}, contentType={}, bytesLength={}",
            input == null ? null : input.fileName(),
            input == null ? null : input.contentType(),
            input == null || input.bytes() == null ? null : input.bytes().length);
        byte[] bytes = input.bytes();
        if (bytes.length >= 4) {
            int b0 = bytes[0] & 0xFF;
            int b1 = bytes[1] & 0xFF;
            int b2 = bytes[2] & 0xFF;
            int b3 = bytes[3] & 0xFF;

            // ZIP signature (xlsx/ods/docx/odt)
            if (b0 == ZIP_SIGNATURE_0 && b1 == ZIP_SIGNATURE_1 && b2 == ZIP_SIGNATURE_2 && b3 == ZIP_SIGNATURE_3) {
                TemplateFormat zipFormat = detectZipContainer(bytes);
                if (zipFormat != null) {
                    log.info("detectFormat() - end: format={}", zipFormat);
                    return zipFormat;
                }
            }

            // OLE2 signature (xls)
            if (b0 == OLE2_SIGNATURE_0 && b1 == OLE2_SIGNATURE_1 && b2 == OLE2_SIGNATURE_2 && b3 == OLE2_SIGNATURE_3) {
                TemplateFormat ole2Format = detectOle2Container(bytes, input);
                if (ole2Format != null) {
                    log.info("detectFormat() - end: format={}", ole2Format);
                    return ole2Format;
                }
            }

            // PDF signature (%PDF)
            if (b0 == 0x25 && b1 == 0x50 && b2 == 0x44 && b3 == 0x46) {
                log.info("detectFormat() - end: format={}", TemplateFormat.PDF);
                return TemplateFormat.PDF;
            }
        }

        TemplateFormat contentTypeFormat = detectByContentType(input.contentType());
        if (contentTypeFormat != null) {
            log.info("detectFormat() - end: format={}", contentTypeFormat);
            return contentTypeFormat;
        }
        TemplateFormat fileNameFormat = detectByFileName(input.fileName());
        if (fileNameFormat != null) {
            log.info("detectFormat() - end: format={}", fileNameFormat);
            return fileNameFormat;
        }

        log.error("detectFormat() - end with error: unsupportedFormat, fileName={}", input.fileName());
        throw new UnsupportedTemplateFormatException("Unsupported template format for file: " + input.fileName());
    }

    /**
     * Detects ZIP-based formats by inspecting container entries.
     *
     * @param bytes source bytes with ZIP signature
     * @return detected format or {@code null} when container cannot be recognized
     */
    private static TemplateFormat detectZipContainer(byte[] bytes) {
        log.info("detectZipContainer() - start: bytesLength={}", bytes == null ? null : bytes.length);
        boolean hasWordFolder = false;
        boolean hasExcelFolder = false;
        boolean hasOdfSpreadsheetMime = false;
        boolean hasOdfTextMime = false;

        try (ZipInputStream zipInputStream = new ZipInputStream(new ByteArrayInputStream(bytes), StandardCharsets.UTF_8)) {
            ZipEntry entry;
            while ((entry = zipInputStream.getNextEntry()) != null) {
                String name = entry.getName();
                if (name == null) {
                    continue;
                }

                if ("mimetype".equals(name)) {
                    String mime = new String(zipInputStream.readAllBytes(), StandardCharsets.UTF_8);
                    if (mime.contains("application/vnd.oasis.opendocument.spreadsheet")) {
                        hasOdfSpreadsheetMime = true;
                    } else if (mime.contains("application/vnd.oasis.opendocument.text")) {
                        hasOdfTextMime = true;
                    }
                    continue;
                }

                if (name.startsWith("word/")) {
                    hasWordFolder = true;
                } else if (name.startsWith("xl/")) {
                    hasExcelFolder = true;
                }
            }
        } catch (Exception ignored) {
            log.warn("detectZipContainer() - end with warning: parseFailed=true");
            return null;
        }

        if (hasOdfSpreadsheetMime) {
            log.info("detectZipContainer() - end: format={}", TemplateFormat.ODS);
            return TemplateFormat.ODS;
        }
        if (hasOdfTextMime) {
            log.info("detectZipContainer() - end: format={}", TemplateFormat.ODT);
            return TemplateFormat.ODT;
        }
        if (hasWordFolder) {
            log.info("detectZipContainer() - end: format={}", TemplateFormat.DOCX);
            return TemplateFormat.DOCX;
        }
        if (hasExcelFolder) {
            log.info("detectZipContainer() - end: format={}", TemplateFormat.XLSX);
            return TemplateFormat.XLSX;
        }
        log.info("detectZipContainer() - end: format=unknown");
        return null;
    }

    /**
     * Detects OLE2 formats by root entries with fallback to metadata hints.
     *
     * @param bytes source OLE2 bytes
     * @param input original input for filename/content-type hints
     * @return detected format (defaults to {@link TemplateFormat#XLS} for backward compatibility)
     */
    private static TemplateFormat detectOle2Container(byte[] bytes, TemplateInput input) {
        log.info("detectOle2Container() - start: bytesLength={}, fileName={}, contentType={}",
            bytes == null ? null : bytes.length,
            input == null ? null : input.fileName(),
            input == null ? null : input.contentType());
        try (POIFSFileSystem fileSystem = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
            DirectoryNode root = fileSystem.getRoot();
            if (root.hasEntry("WordDocument")) {
                log.info("detectOle2Container() - end: format={}", TemplateFormat.DOC);
                return TemplateFormat.DOC;
            }
            if (root.hasEntry("Workbook") || root.hasEntry("Book")) {
                log.info("detectOle2Container() - end: format={}", TemplateFormat.XLS);
                return TemplateFormat.XLS;
            }
        } catch (Exception ignored) {
            log.warn("detectOle2Container() - warning: failedToInspectOle2=true");
        }

        if (hasDocHint(input)) {
            log.info("detectOle2Container() - end: format={}", TemplateFormat.DOC);
            return TemplateFormat.DOC;
        }
        if (hasXlsHint(input)) {
            log.info("detectOle2Container() - end: format={}", TemplateFormat.XLS);
            return TemplateFormat.XLS;
        }

        // keep backward-compatible default for unknown OLE2 containers
        log.info("detectOle2Container() - end: format={}", TemplateFormat.XLS);
        return TemplateFormat.XLS;
    }

    /**
     * Detects format by MIME content type hint.
     *
     * @param contentType MIME content type
     * @return detected format or {@code null}
     */
    private static TemplateFormat detectByContentType(String contentType) {
        if (contentType == null) {
            return null;
        }
        String lower = contentType.toLowerCase(Locale.ROOT);
        if (lower.contains("spreadsheetml")) {
            return TemplateFormat.XLSX;
        }
        if (lower.contains("ms-excel")) {
            return TemplateFormat.XLS;
        }
        if (lower.contains("oasis.opendocument.spreadsheet")) {
            return TemplateFormat.ODS;
        }
        if (lower.contains("msword")) {
            return TemplateFormat.DOC;
        }
        if (lower.contains("wordprocessingml.document")) {
            return TemplateFormat.DOCX;
        }
        if (lower.contains("oasis.opendocument.text")) {
            return TemplateFormat.ODT;
        }
        if (lower.contains("application/pdf")) {
            return TemplateFormat.PDF;
        }
        return null;
    }

    /**
     * Detects format by file extension.
     *
     * @param fileName file name
     * @return detected format or {@code null}
     */
    private static TemplateFormat detectByFileName(String fileName) {
        if (fileName == null) {
            return null;
        }
        String lower = fileName.toLowerCase(Locale.ROOT);
        if (lower.endsWith(".xlsx")) {
            return TemplateFormat.XLSX;
        }
        if (lower.endsWith(".xls")) {
            return TemplateFormat.XLS;
        }
        if (lower.endsWith(".ods")) {
            return TemplateFormat.ODS;
        }
        if (lower.endsWith(".doc")) {
            return TemplateFormat.DOC;
        }
        if (lower.endsWith(".docx")) {
            return TemplateFormat.DOCX;
        }
        if (lower.endsWith(".odt")) {
            return TemplateFormat.ODT;
        }
        if (lower.endsWith(".pdf")) {
            return TemplateFormat.PDF;
        }
        return null;
    }

    /**
     * Checks whether input metadata hints DOC format.
     *
     * @param input template input
     * @return {@code true} when metadata points to DOC
     */
    private static boolean hasDocHint(TemplateInput input) {
        if (input.contentType() != null && input.contentType().toLowerCase(Locale.ROOT).contains("msword")) {
            return true;
        }
        return input.fileName() != null && input.fileName().toLowerCase(Locale.ROOT).endsWith(".doc");
    }

    /**
     * Checks whether input metadata hints XLS format.
     *
     * @param input template input
     * @return {@code true} when metadata points to XLS
     */
    private static boolean hasXlsHint(TemplateInput input) {
        if (input.contentType() != null && input.contentType().toLowerCase(Locale.ROOT).contains("ms-excel")) {
            return true;
        }
        return input.fileName() != null && input.fileName().toLowerCase(Locale.ROOT).endsWith(".xls");
    }
}
