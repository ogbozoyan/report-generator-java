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

    public static TemplateFormat detectRequestedOutputFormat(TemplateInput input) {
        if (input.contentType() != null) {
            String contentType = input.contentType().toLowerCase(Locale.ROOT);
            if (contentType.contains("spreadsheetml")) {
                return TemplateFormat.XLSX;
            }
            if (contentType.contains("ms-excel")) {
                return TemplateFormat.XLS;
            }
            if (contentType.contains("oasis.opendocument.spreadsheet")) {
                return TemplateFormat.ODS;
            }
            if (contentType.contains("msword")) {
                return TemplateFormat.DOC;
            }
            if (contentType.contains("wordprocessingml.document")) {
                return TemplateFormat.DOCX;
            }
            if (contentType.contains("oasis.opendocument.text")) {
                return TemplateFormat.ODT;
            }
            if (contentType.contains("application/pdf")) {
                return TemplateFormat.PDF;
            }
        }

        if (input.fileName() != null) {
            String name = input.fileName().toLowerCase(Locale.ROOT);
            if (name.endsWith(".xlsx")) {
                return TemplateFormat.XLSX;
            }
            if (name.endsWith(".xls")) {
                return TemplateFormat.XLS;
            }
            if (name.endsWith(".ods")) {
                return TemplateFormat.ODS;
            }
            if (name.endsWith(".doc")) {
                return TemplateFormat.DOC;
            }
            if (name.endsWith(".docx")) {
                return TemplateFormat.DOCX;
            }
            if (name.endsWith(".odt")) {
                return TemplateFormat.ODT;
            }
            if (name.endsWith(".pdf")) {
                return TemplateFormat.PDF;
            }
        }


        return null;
    }

    public static TemplateFormat detectFormat(TemplateInput input) {
        log.debug("detectFormat() - start: input {}", input);
        byte[] bytes = input.bytes();
        if (bytes.length >= 4) {
            log.debug("detectFormat() - bytes length: {}", bytes.length);
            int b0 = bytes[0] & 0xFF;
            int b1 = bytes[1] & 0xFF;
            int b2 = bytes[2] & 0xFF;
            int b3 = bytes[3] & 0xFF;

            // ZIP signature (xlsx/ods/docx/odt)
            if (b0 == ZIP_SIGNATURE_0 && b1 == ZIP_SIGNATURE_1 && b2 == ZIP_SIGNATURE_2 && b3 == ZIP_SIGNATURE_3) {
                TemplateFormat zipFormat = detectZipContainer(bytes);
                if (zipFormat != null) {
                    return zipFormat;
                }
            }

            // OLE2 signature (xls)
            if (b0 == OLE2_SIGNATURE_0 && b1 == OLE2_SIGNATURE_1 && b2 == OLE2_SIGNATURE_2 && b3 == OLE2_SIGNATURE_3) {
                TemplateFormat ole2Format = detectOle2Container(bytes, input);
                if (ole2Format != null) {
                    return ole2Format;
                }
            }

            // PDF signature (%PDF)
            if (b0 == 0x25 && b1 == 0x50 && b2 == 0x44 && b3 == 0x46) {
                return TemplateFormat.PDF;
            }
        }

        if (input.contentType() != null) {
            log.debug("detectFormat() - contentType: {}", input.contentType());
            String contentType = input.contentType().toLowerCase(Locale.ROOT);
            if (contentType.contains("spreadsheetml")) {
                return TemplateFormat.XLSX;
            }
            if (contentType.contains("ms-excel")) {
                return TemplateFormat.XLS;
            }
            if (contentType.contains("oasis.opendocument.spreadsheet")) {
                return TemplateFormat.ODS;
            }
            if (contentType.contains("msword")) {
                return TemplateFormat.DOC;
            }
            if (contentType.contains("wordprocessingml.document")) {
                return TemplateFormat.DOCX;
            }
            if (contentType.contains("oasis.opendocument.text")) {
                return TemplateFormat.ODT;
            }
            if (contentType.contains("application/pdf")) {
                return TemplateFormat.PDF;
            }
        }
        if (input.fileName() != null) {
            log.debug("detectFormat() - fileName: {}", input.fileName());
            String name = input.fileName().toLowerCase(Locale.ROOT);
            if (name.endsWith(".xlsx")) {
                return TemplateFormat.XLSX;
            }
            if (name.endsWith(".xls")) {
                return TemplateFormat.XLS;
            }
            if (name.endsWith(".ods")) {
                return TemplateFormat.ODS;
            }
            if (name.endsWith(".doc")) {
                return TemplateFormat.DOC;
            }
            if (name.endsWith(".docx")) {
                return TemplateFormat.DOCX;
            }
            if (name.endsWith(".odt")) {
                return TemplateFormat.ODT;
            }
            if (name.endsWith(".pdf")) {
                return TemplateFormat.PDF;
            }
        }

        throw new UnsupportedTemplateFormatException("Unsupported template format for file: " + input.fileName());
    }

    private static TemplateFormat detectZipContainer(byte[] bytes) {
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
            return null;
        }

        if (hasOdfSpreadsheetMime) {
            return TemplateFormat.ODS;
        }
        if (hasOdfTextMime) {
            return TemplateFormat.ODT;
        }
        if (hasWordFolder) {
            return TemplateFormat.DOCX;
        }
        if (hasExcelFolder) {
            return TemplateFormat.XLSX;
        }
        return null;
    }

    private static TemplateFormat detectOle2Container(byte[] bytes, TemplateInput input) {
        try (POIFSFileSystem fileSystem = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
            DirectoryNode root = fileSystem.getRoot();
            if (root.hasEntry("WordDocument")) {
                return TemplateFormat.DOC;
            }
            if (root.hasEntry("Workbook") || root.hasEntry("Book")) {
                return TemplateFormat.XLS;
            }
        } catch (Exception ignored) {
            // fallback to metadata hints below
        }

        if (hasDocHint(input)) {
            return TemplateFormat.DOC;
        }
        if (hasXlsHint(input)) {
            return TemplateFormat.XLS;
        }

        // keep backward-compatible default for unknown OLE2 containers
        return TemplateFormat.XLS;
    }

    private static boolean hasDocHint(TemplateInput input) {
        if (input.contentType() != null && input.contentType().toLowerCase(Locale.ROOT).contains("msword")) {
            return true;
        }
        return input.fileName() != null && input.fileName().toLowerCase(Locale.ROOT).endsWith(".doc");
    }

    private static boolean hasXlsHint(TemplateInput input) {
        if (input.contentType() != null && input.contentType().toLowerCase(Locale.ROOT).contains("ms-excel")) {
            return true;
        }
        return input.fileName() != null && input.fileName().toLowerCase(Locale.ROOT).endsWith(".xls");
    }
}
