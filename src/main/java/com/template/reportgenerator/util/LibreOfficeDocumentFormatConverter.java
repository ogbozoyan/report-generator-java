package com.template.reportgenerator.util;

import com.template.reportgenerator.contract.TemplateFormat;
import com.template.reportgenerator.exception.TemplateReadWriteException;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.TimeUnit;
import java.util.stream.Stream;

/**
 * Converts documents via LibreOffice/soffice in headless mode.
 */
@Slf4j
public class LibreOfficeDocumentFormatConverter implements DocumentFormatConverter {

    private static final List<String> OFFICE_BINARIES = List.of("soffice", "libreoffice");
    private static final long CONVERT_TIMEOUT_SECONDS = 60L;

    @Override
    public byte[] convert(byte[] sourceBytes, TemplateFormat sourceFormat, TemplateFormat targetFormat) {
        if (sourceFormat == targetFormat) {
            return sourceBytes;
        }

        String binary = resolveOfficeBinary();
        String targetExtension = targetFormat.extension().substring(1).toLowerCase(Locale.ROOT);

        Path workDir = null;
        try {
            workDir = Files.createTempDirectory("reportgen-convert-");
            Path inputFile = workDir.resolve("generated" + sourceFormat.extension());
            Files.write(inputFile, sourceBytes);

            List<String> command = List.of(
                binary,
                "--headless",
                "--convert-to",
                targetExtension,
                "--outdir",
                workDir.toString(),
                inputFile.toString()
            );

            log.info("Converting {} -> {} using {}", sourceFormat, targetFormat, binary);
            Process process = new ProcessBuilder(command)
                .redirectErrorStream(true)
                .start();

            boolean finished = process.waitFor(CONVERT_TIMEOUT_SECONDS, TimeUnit.SECONDS);
            String processOutput = new String(process.getInputStream().readAllBytes(), StandardCharsets.UTF_8);

            if (!finished) {
                process.destroyForcibly();
                throw new TemplateReadWriteException("LibreOffice conversion timed out: " + sourceFormat + " -> " + targetFormat);
            }

            if (process.exitValue() != 0) {
                throw new TemplateReadWriteException(
                    "LibreOffice conversion failed (" + process.exitValue() + "): " + processOutput
                );
            }

            Path outputFile = workDir.resolve("generated." + targetExtension);
            if (!Files.exists(outputFile)) {
                throw new TemplateReadWriteException(
                    "LibreOffice did not produce output file for " + sourceFormat + " -> " + targetFormat
                );
            }

            return Files.readAllBytes(outputFile);
        } catch (IOException e) {
            throw new TemplateReadWriteException("Failed to execute LibreOffice conversion", e);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new TemplateReadWriteException("LibreOffice conversion interrupted", e);
        } finally {
            if (workDir != null) {
                deleteRecursively(workDir);
            }
        }
    }

    private String resolveOfficeBinary() {
        for (String candidate : OFFICE_BINARIES) {
            try {
                Process process = new ProcessBuilder(candidate, "--version")
                    .redirectErrorStream(true)
                    .start();
                if (process.waitFor(5, TimeUnit.SECONDS) && process.exitValue() == 0) {
                    return candidate;
                }
            } catch (Exception ignored) {
                // try next candidate
            }
        }
        throw new TemplateReadWriteException(
            "LibreOffice is required for ODS/ODT export. Install 'soffice' or 'libreoffice' and add it to PATH."
        );
    }

    private void deleteRecursively(Path root) {
        try (Stream<Path> stream = Files.walk(root)) {
            stream.sorted(Comparator.reverseOrder())
                .forEach(path -> {
                    try {
                        Files.deleteIfExists(path);
                    } catch (IOException ignored) {
                        log.debug("");
                    }
                });
        } catch (IOException ignored) {
            log.debug("");
        }
    }
}
