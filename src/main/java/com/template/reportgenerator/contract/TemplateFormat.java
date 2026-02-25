package com.template.reportgenerator.contract;

/**
 * Supported template formats and their output metadata.
 */
public enum TemplateFormat {
    XLS("application/vnd.ms-excel", ".xls"),
    XLSX("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    ODS("application/vnd.oasis.opendocument.spreadsheet", ".ods"),
    DOC("application/msword", ".doc"),
    DOCX("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
    ODT("application/vnd.oasis.opendocument.text", ".odt"),
    PDF("application/pdf", ".pdf");

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
