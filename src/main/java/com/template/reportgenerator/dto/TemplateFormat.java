package com.template.reportgenerator.dto;

public enum TemplateFormat {
    XLS("application/vnd.ms-excel", ".xls"),
    XLSX("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    ODS("application/vnd.oasis.opendocument.spreadsheet", ".ods");

    private final String contentType;
    private final String extension;

    TemplateFormat(String contentType, String extension) {
        this.contentType = contentType;
        this.extension = extension;
    }

    public String contentType() {
        return contentType;
    }

    public String extension() {
        return extension;
    }
}
