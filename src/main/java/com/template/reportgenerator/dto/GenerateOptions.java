package com.template.reportgenerator.dto;

import java.time.ZoneId;
import java.util.Locale;

public record GenerateOptions(
    MissingValuePolicy missingValuePolicy,
    boolean recalculateFormulas,
    Locale locale,
    ZoneId zoneId
) {
    public GenerateOptions {
        missingValuePolicy = missingValuePolicy == null ? MissingValuePolicy.EMPTY_AND_LOG : missingValuePolicy;
        locale = locale == null ? Locale.getDefault() : locale;
        zoneId = zoneId == null ? ZoneId.systemDefault() : zoneId;
    }

    public static GenerateOptions defaults() {
        return new GenerateOptions(MissingValuePolicy.EMPTY_AND_LOG, true, Locale.getDefault(), ZoneId.systemDefault());
    }
}
