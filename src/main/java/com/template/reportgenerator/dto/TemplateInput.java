package com.template.reportgenerator.dto;

import java.util.Objects;

public record TemplateInput(
    String fileName,
    String contentType,
    byte[] bytes
) {
    public TemplateInput {
        Objects.requireNonNull(bytes, "bytes must not be null");
    }
}
