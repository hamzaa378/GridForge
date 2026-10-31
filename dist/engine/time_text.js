import { localeNumberSeparators } from './locale_number.js';
const escapeRegex = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
let cachedClockSettings;
function createLocaleClock() {
    const locale = Intl.DateTimeFormat().resolvedOptions().locale, periods = new Map();
    const dayPeriod = (hour) => new Intl.DateTimeFormat(locale, { hour: 'numeric', minute: '2-digit', hourCycle: 'h12', timeZone: 'UTC' }).formatToParts(new Date(Date.UTC(2020, 0, 1, hour, 5))).find(part => part.type === 'dayPeriod')?.value;
    for (const [hour, period] of [[9, 'AM'], [21, 'PM']]) {
        const value = dayPeriod(hour);
        if (value) {
            const raw = value.toLocaleLowerCase(locale).trim();
            periods.set(raw, period);
            periods.set(raw.replace(/\./g, ''), period);
        }
    }
    periods.set('am', 'AM');
    periods.set('pm', 'PM');
    const parts = new Intl.DateTimeFormat(locale, { hour: '2-digit', minute: '2-digit', hourCycle: 'h23', timeZone: 'UTC' }).formatToParts(new Date(Date.UTC(2020, 0, 1, 13, 5))), hourIndex = parts.findIndex(part => part.type === 'hour'), minuteIndex = parts.findIndex(part => part.type === 'minute'), separator = parts.slice(hourIndex + 1, minuteIndex).filter(part => part.type === 'literal').map(part => part.value.replace(/[\u200e\u200f\u061c]/g, '').trim()).join('') || ':';
    const separators = [...new Set([':', separator])].map(escapeRegex).join('|'), markers = [...periods.keys()].sort((a, b) => b.length - a.length).map(escapeRegex).join('|'), body = `(\\d{1,2})(?:${separators})(\\d{2})(?:(?:${separators})(\\d{2}))?\\s*(${markers})?`;
    return { locale, periods, body };
}
const localeClock = () => cachedClockSettings ?? (cachedClockSettings = createLocaleClock());
/** Parse a time-only string to an Excel fractional-day value, or return null. */
export function parseTimeOfDayText(raw) {
    const settings = localeClock(), normalized = localeNumberSeparators().normalizeDigits(raw.normalize('NFKC')).trim(), match = new RegExp(`^${settings.body}$`, 'i').exec(normalized);
    if (!match)
        return null;
    let hour = Number(match[1]), minute = Number(match[2]), second = Number(match[3] ?? 0);
    const period = match[4]?.toLocaleLowerCase(settings.locale).replace(/\./g, '').trim(), meridiem = period ? settings.periods.get(period) : undefined;
    if (period && !meridiem)
        return null;
    if (minute > 59 || second > 59 || (meridiem ? hour < 1 || hour > 12 : hour > 23))
        return null;
    if (meridiem)
        hour = hour % 12 + (meridiem === 'PM' ? 12 : 0);
    return (hour * 3600 + minute * 60 + second) / 86400;
}
/** Split a date/time string using the current locale's clock syntax. */
export function splitTrailingDateTimeText(raw) {
    const settings = localeClock(), normalized = localeNumberSeparators().normalizeDigits(raw.normalize('NFKC')).trim(), match = new RegExp(`^(.+?)[ T]+(?:${settings.body})$`, 'i').exec(normalized);
    return match ? [match[1], normalized.slice(match[1].length).trim()] : null;
}
