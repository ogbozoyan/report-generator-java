package com.template.reportgenerator.util;

import com.template.reportgenerator.contract.MissingValuePolicy;
import com.template.reportgenerator.contract.ResolvedText;
import com.template.reportgenerator.exception.TemplateDataBindingException;
import lombok.experimental.UtilityClass;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Resolves {@code {{token}}} expressions against runtime context.
 */
@UtilityClass
public class TokenResolver {

    public static final Pattern TOKEN_PATTERN = Pattern.compile("\\{\\{\\s*([a-zA-Z0-9_.-]+)\\s*}}");
    public static final Pattern EXACT_TOKEN_PATTERN = Pattern.compile("^\\{\\{\\s*([a-zA-Z0-9_.-]+)\\s*}}$");

    public static boolean hasTokens(String value) {
        return value != null && TOKEN_PATTERN.matcher(value).find();
    }

    public static String getExactToken(String value) {
        if (value == null) {
            return null;
        }
        Matcher matcher = EXACT_TOKEN_PATTERN.matcher(value.trim());
        if (!matcher.matches()) {
            return null;
        }
        return matcher.group(1);
    }

    public static ResolvedText resolve(
        String text,
        Map<String, Object> context,
        MissingValuePolicy policy,
        WarningCollector warningCollector,
        String location,
        boolean allowItemTokens
    ) {
        if (text == null || text.isEmpty()) {
            return new ResolvedText(text, false);
        }

        Matcher matcher = TOKEN_PATTERN.matcher(text);
        StringBuilder sb = new StringBuilder();
        boolean changed = false;

        while (matcher.find()) {
            String token = matcher.group(1);

            if (!allowItemTokens && isItemOrIndexToken(token)) {
                matcher.appendReplacement(sb, Matcher.quoteReplacement(matcher.group(0)));
                continue;
            }

            Object resolved = resolvePath(context, token);
            if (resolved == null) {
                String replacement = switch (policy) {
                    case EMPTY_AND_LOG -> {
                        warningCollector.add(
                            "MISSING_TOKEN",
                            "Token not found: " + token,
                            location
                        );
                        yield "";
                    }
                    case LEAVE_TOKEN -> matcher.group(0);
                    case FAIL_FAST -> throw new TemplateDataBindingException(
                        "Token not found: " + token + " at " + location
                    );
                };
                matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
                changed = changed || !Objects.equals(matcher.group(0), replacement);
                continue;
            }

            if (isTableValue(resolved)) {
                warningCollector.add(
                    "TABLE_TOKEN_INLINE_IGNORED",
                    "Table token can be inserted only as exact placeholder",
                    location
                );
                matcher.appendReplacement(sb, Matcher.quoteReplacement(matcher.group(0)));
                continue;
            }

            String replacement = stringify(resolved);
            matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
            changed = true;
        }

        matcher.appendTail(sb);
        return new ResolvedText(sb.toString(), changed);
    }

    public static Object resolvePath(Map<String, Object> context, String path) {
        if (context == null || path == null || path.isBlank()) {
            return null;
        }

        if (context.containsKey(path)) {
            return context.get(path);
        }

        String[] parts = path.split("\\.");
        Object current = context.get(parts[0]);

        for (int i = 1; i < parts.length && current != null; i++) {
            current = getChild(current, parts[i]);
        }

        return current;
    }

    /**
     * Checks whether token value can be treated as table payload.
     */
    public static boolean isTableValue(Object value) {
        return toTableRows(value) != null;
    }

    /**
     * Safely converts a value to table rows.
     *
     * @return ordered rows or {@code null} when structure is not {@code List<Map<...>>}
     */
    public static List<Map<String, Object>> toTableRows(Object value) {
        if (!(value instanceof List<?> list)) {
            return null;
        }

        List<Map<String, Object>> rows = new ArrayList<>(list.size());
        for (Object item : list) {
            if (!(item instanceof Map<?, ?> map)) {
                return null;
            }
            LinkedHashMap<String, Object> row = new LinkedHashMap<>();
            for (Map.Entry<?, ?> entry : map.entrySet()) {
                String key = entry.getKey() == null ? "" : String.valueOf(entry.getKey());
                row.put(key, entry.getValue());
            }
            rows.add(row);
        }
        return rows;
    }

    public static boolean isItemOrIndexToken(String token) {
        return "index".equals(token) || token.startsWith("item.");
    }

    private Object getChild(Object current, String key) {
        if (current instanceof Map<?, ?> map) {
            return map.get(key);
        }

        String getterName = "get" + key.substring(0, 1).toUpperCase(Locale.ROOT) + key.substring(1);
        try {
            Method getter = current.getClass().getMethod(getterName);
            return getter.invoke(current);
        } catch (Exception ignored) {
            return null;
        }
    }

    private String stringify(Object value) {
        if (value == null) {
            return "";
        }
        return String.valueOf(value);
    }
}
