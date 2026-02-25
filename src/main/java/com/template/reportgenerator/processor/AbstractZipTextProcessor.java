package com.template.reportgenerator.processor;

import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.TemplateScanResult;
import com.template.reportgenerator.dto.TokenOccurrence;
import com.template.reportgenerator.exception.TemplateReadWriteException;
import com.template.reportgenerator.util.TextTemplateEngine;
import com.template.reportgenerator.util.WarningCollector;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * Base processor for ZIP-based text documents (for example DOCX and ODT).
 * <p>
 * It replaces scalar tokens inside selected XML entries and keeps all other entries untouched.
 * TABLE/COL expansion is not supported in text document formats and therefore treated as no-op.
 */
abstract class AbstractZipTextProcessor implements WorkbookProcessor {

    private final List<ArchiveEntry> entries;

    protected AbstractZipTextProcessor(byte[] bytes, String formatName) {
        this.entries = readEntries(bytes, formatName);
    }

    @Override
    public TemplateScanResult scan() {
        return new TemplateScanResult(List.of(), List.<TokenOccurrence>of());
    }

    @Override
    public void applyScalarTokens(Map<String, Object> scalars, GenerateOptions options, WarningCollector warningCollector) {
        for (int i = 0; i < entries.size(); i++) {
            ArchiveEntry entry = entries.get(i);
            if (entry.directory() || !shouldProcessEntry(entry.name())) {
                continue;
            }

            String source = new String(entry.bytes(), StandardCharsets.UTF_8);
            String replaced = TextTemplateEngine.replaceText(source, scalars, options, warningCollector, entry.name());
            if (!source.equals(replaced)) {
                entries.set(i, entry.withBytes(replaced.getBytes(StandardCharsets.UTF_8)));
            }
        }
    }

    @Override
    public byte[] serialize() {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
             ZipOutputStream zipOutputStream = new ZipOutputStream(outputStream, StandardCharsets.UTF_8)) {

            for (ArchiveEntry entry : entries) {
                ZipEntry zipEntry = new ZipEntry(entry.name());
                if (entry.time() > 0) {
                    zipEntry.setTime(entry.time());
                }
                if (entry.comment() != null) {
                    zipEntry.setComment(entry.comment());
                }
                if (entry.extra() != null) {
                    zipEntry.setExtra(entry.extra());
                }

                zipOutputStream.putNextEntry(zipEntry);
                if (!entry.directory()) {
                    zipOutputStream.write(entry.bytes());
                }
                zipOutputStream.closeEntry();
            }
            zipOutputStream.finish();
            return outputStream.toByteArray();
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to serialize ZIP text document", e);
        }
    }

    /**
     * Returns whether the specific ZIP entry should be treated as editable text source.
     */
    protected abstract boolean shouldProcessEntry(String entryName);

    private static List<ArchiveEntry> readEntries(byte[] bytes, String formatName) {
        try (ZipInputStream zipInputStream = new ZipInputStream(new ByteArrayInputStream(bytes), StandardCharsets.UTF_8)) {
            List<ArchiveEntry> result = new ArrayList<>();
            ZipEntry zipEntry;
            while ((zipEntry = zipInputStream.getNextEntry()) != null) {
                byte[] entryBytes = zipEntry.isDirectory() ? new byte[0] : zipInputStream.readAllBytes();
                result.add(new ArchiveEntry(
                    zipEntry.getName(),
                    zipEntry.isDirectory(),
                    entryBytes,
                    zipEntry.getTime(),
                    zipEntry.getExtra(),
                    zipEntry.getComment()
                ));
            }
            return result;
        } catch (Exception e) {
            throw new TemplateReadWriteException("Failed to read " + formatName + " template", e);
        }
    }

    private record ArchiveEntry(
        String name,
        boolean directory,
        byte[] bytes,
        long time,
        byte[] extra,
        String comment
    ) {
        private ArchiveEntry withBytes(byte[] updatedBytes) {
            return new ArchiveEntry(name, directory, updatedBytes, time, extra, comment);
        }
    }
}
