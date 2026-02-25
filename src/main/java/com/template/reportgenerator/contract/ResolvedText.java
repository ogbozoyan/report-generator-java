package com.template.reportgenerator.contract;

/**
 * Text replacement result flagging whether source changed.
 */
public record ResolvedText(String value, boolean changed) {
}
