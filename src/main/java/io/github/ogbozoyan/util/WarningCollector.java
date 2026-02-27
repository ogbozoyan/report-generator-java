package io.github.ogbozoyan.util;


import io.github.ogbozoyan.contract.GenerationWarning;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Mutable collector for generation warnings.
 */
public class WarningCollector {

    private final List<GenerationWarning> warnings = new ArrayList<>();

    /**
     * Adds warning entry.
     *
     * @param code     warning code
     * @param message  warning message
     * @param location template location
     */
    public void add(String code, String message, String location) {
        warnings.add(new GenerationWarning(code, message, location));
    }

    /**
     * Returns read-only warning list snapshot.
     *
     * @return immutable warning list
     */
    public List<GenerationWarning> asList() {
        return Collections.unmodifiableList(warnings);
    }
}
