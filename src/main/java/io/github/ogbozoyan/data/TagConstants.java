package io.github.ogbozoyan.data;

import lombok.Getter;

@Getter
public enum TagConstants {
    ROWS_COLUMNS("rows__columns");

    private final String value;

    TagConstants(String value) {
        this.value = value;
    }
}
