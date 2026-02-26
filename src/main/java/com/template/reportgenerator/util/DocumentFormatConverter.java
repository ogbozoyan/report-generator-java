package com.template.reportgenerator.util;

import com.template.reportgenerator.contract.TemplateFormat;

/**
 * Converts generated document bytes between output formats.
 */
public interface DocumentFormatConverter {

    byte[] convert(byte[] sourceBytes, TemplateFormat sourceFormat, TemplateFormat targetFormat);
}
