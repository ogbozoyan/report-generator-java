package com.template.reportgenerator.service;

import com.template.reportgenerator.dto.GenerateOptions;
import com.template.reportgenerator.dto.GeneratedReport;
import com.template.reportgenerator.dto.ReportData;
import com.template.reportgenerator.dto.TemplateInput;

public interface ReportGeneratorService {
    GeneratedReport generate(TemplateInput template, ReportData data, GenerateOptions options);
}
