package com.template.reportgenerator.contract;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.List;
import java.util.Map;

public record DocxTableAnchor(XWPFParagraph paragraph, String token, List<Map<String, Object>> rows) {
}
