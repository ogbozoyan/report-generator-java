package io.github.ogbozoyan.util;


import io.github.ogbozoyan.data.MissingValuePolicy;
import io.github.ogbozoyan.data.ResolvedText;
import io.github.ogbozoyan.exception.TemplateDataBindingException;
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

    /**
     * Pattern for token occurrence detection.
     */
    public static final Pattern TOKEN_PATTERN = Pattern.compile("\\{\\{\\s*([a-zA-Z0-9_.-]+)\\s*}}");
    /**
     * Pattern for exact placeholder detection (whole value is one token).
     */
    public static final Pattern EXACT_TOKEN_PATTERN = Pattern.compile("^\\{\\{\\s*([a-zA-Z0-9_.-]+)\\s*}}$");

    /**
     * Checks whether string contains at least one token expression.
     *
     * @param value source text
     * @return {@code true} when token is present
     */
    public static boolean hasTokens(String value) {
        return value != null && TOKEN_PATTERN.matcher(value).find();
    }

    /**
     * Returns token name if value is exact placeholder.
     *
     * @param value source value
     * @return token name or {@code null}
     */
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

    /**
     * Returns token name when text contains exactly one {@code {{token}}} occurrence.
     * If there are multiple tokens, returns {@code null}.
     *
     * @param value source value
     * @return single token name or {@code null}
     */
    public static String getSingleToken(String value) {
        if (value == null) {
            return null;
        }
        Matcher matcher = TOKEN_PATTERN.matcher(value);
        String token = null;
        while (matcher.find()) {
            if (token != null) {
                return null;
            }
            token = matcher.group(1);
        }
        return token;
    }

    /**
     * Resolves token expressions in free-form text.
     *
     * <p>Missing token behavior is controlled by {@link MissingValuePolicy}. Table values are not
     * expanded inline and produce warning instead.
     *
     * @param text             source text
     * @param context          token context
     * @param policy           missing token policy
     * @param warningCollector warning collector
     * @param location         diagnostic location
     * @param allowItemTokens  whether {@code index}/{@code item.*} placeholders are allowed
     * @return resolved text and changed flag
     */
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
                            "Token not found in file but was present in template: " + token,
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

    /**
     * Resolves dotted path against token context.
     *
     * <p>Supports direct key lookup and dotted traversal through map/object getters.
     *
     * @param context root context
     * @param path    token path
     * @return resolved value or {@code null}
     */
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
     *
     * @param value token value
     * @return {@code true} when value is {@code List<Map<...>>}
     */
    public static boolean isTableValue(Object value) {
        return toTableRows(value) != null;
    }

    /**
     * Safely converts a value to table rows.
     *
     * @param value source token value
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

    /**
     * Checks whether token is special context token used in legacy item expansion flow.
     *
     * @param token token name
     * @return {@code true} for {@code index} or {@code item.*}
     */
    public static boolean isItemOrIndexToken(String token) {
        return "index".equals(token) || token.startsWith("item.");
    }

    /**
     * Resolves child property from map key or JavaBean getter.
     *
     * @param current current object
     * @param key     property key
     * @return child value or {@code null}
     */
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

    /**
     * Converts resolved value to string replacement.
     *
     * @param value resolved value
     * @return replacement text
     */
    private String stringify(Object value) {
        if (value == null) {
            return "";
        }
        return String.valueOf(value);
    }
}
