let cachedLocaleNumberSeparators;
export function localeNumberSeparators() {
    if (cachedLocaleNumberSeparators)
        return cachedLocaleNumberSeparators;
    const locale = Intl.NumberFormat().resolvedOptions().locale, parts = new Intl.NumberFormat(locale, { useGrouping: true }).formatToParts(123456789.5), decimal = parts.find(part => part.type === 'decimal')?.value ?? '.', group = parts.find(part => part.type === 'group')?.value ?? ',', digits = new Map();
    for (let digit = 0; digit <= 9; digit++) {
        const shown = new Intl.NumberFormat(locale, { useGrouping: false, maximumFractionDigits: 0 }).formatToParts(digit).find(part => part.type === 'integer')?.value;
        if (shown)
            digits.set(shown, String(digit));
    }
    return cachedLocaleNumberSeparators = { decimal, group, normalizeDigits: (value) => Array.from(value, char => digits.get(char) ?? char).join('') };
}
/** Return null for invalid numeric text, or a finite number using Excel percent rules. */
export function parseLocaleNumberText(raw, decimalArg, groupArg) {
    const locale = localeNumberSeparators(), decimal = (decimalArg === undefined ? locale.decimal : decimalArg[0]) ?? '', group = (groupArg === undefined ? locale.group : groupArg[0]) ?? '';
    if (!decimal || (group && decimal === group))
        return null;
    let input = locale.normalizeDigits(raw).replace(/[ \t\r\n\u00a0\u2000-\u200a\u202f\u205f\u3000]/gu, space => space === group ? space : '');
    if (input === '')
        return 0;
    let percentCount = 0;
    while (input.endsWith('%')) {
        percentCount++;
        input = input.slice(0, -1);
    }
    if (input.includes('%'))
        return null;
    const exponentParts = input.split(/[eE]/);
    if (exponentParts.length > 2)
        return null;
    let mantissa = exponentParts[0];
    const exponent = exponentParts.length === 2 ? `e${exponentParts[1]}` : '';
    const decimalIndex = mantissa.indexOf(decimal);
    if (decimalIndex >= 0 && mantissa.indexOf(decimal, decimalIndex + decimal.length) >= 0)
        return null;
    const integerPart = decimalIndex < 0 ? mantissa : mantissa.slice(0, decimalIndex), fractionPart = decimalIndex < 0 ? '' : mantissa.slice(decimalIndex + decimal.length);
    if (group && fractionPart.includes(group))
        return null;
    mantissa = integerPart.split(group).join('') + (decimalIndex < 0 ? '' : '.' + fractionPart);
    const cleaned = mantissa + exponent;
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(cleaned))
        return null;
    let result = Number(cleaned);
    for (let index = 0; index < percentCount; index++)
        result /= 100;
    return Number.isFinite(result) ? result : null;
}
/** Parse VALUE-style currency text while leaving NUMBERVALUE's grammar unchanged. */
export function parseLocaleCurrencyText(raw) {
    const input = raw.trim();
    const prefix = /^([+-]?)\s*(\p{Sc})\s*(.*?)$/u.exec(input);
    if (prefix && prefix[3]?.trim())
        return parseLocaleNumberText(`${prefix[1] ?? ''}${prefix[3]}`);
    const suffix = /^(.*?)\s*(\p{Sc})$/u.exec(input);
    return suffix && suffix[1]?.trim() ? parseLocaleNumberText(suffix[1]) : null;
}
