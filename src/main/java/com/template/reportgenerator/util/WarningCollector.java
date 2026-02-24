package com.template.reportgenerator.util;


import com.template.reportgenerator.dto.GenerationWarning;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Mutable collector for generation warnings.
 */
public class WarningCollector {

    private final List<GenerationWarning> warnings = new ArrayList<>();

    public void add(String code, String message, String location) {
        warnings.add(new GenerationWarning(code, message, location));
    }

    public List<GenerationWarning> asList() {
        return Collections.unmodifiableList(warnings);
    }
}
