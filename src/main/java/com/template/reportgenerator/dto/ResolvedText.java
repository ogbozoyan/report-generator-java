package com.template.reportgenerator.dto;

/**
 * Text replacement result flagging whether source changed.
 */
public record ResolvedText(String value, boolean changed) {
}
