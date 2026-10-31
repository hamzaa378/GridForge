const DAY_MS = 86_400_000;
const EXCEL_EPOCH_1900 = Date.UTC(1899, 11, 31);
const EXCEL_EPOCH_AFTER_FAKE = Date.UTC(1899, 11, 30);
const EXCEL_FICTITIOUS_LEAP_DAY = Date.UTC(1900, 2, 1);
/** Convert an ISO date/time to the Excel 1900 date system, including its fake leap day. */
export function excelDateSerial(value) {
    if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}(?:T.*)?$/.test(value))
        return null;
    if (value === '1900-02-29')
        return 60;
    const day = value.slice(0, 10), dayDate = new Date(`${day}T00:00:00Z`);
    if (!Number.isFinite(dayDate.getTime()) || dayDate.toISOString().slice(0, 10) !== day)
        return null;
    const timestamp = value.includes('T') ? Date.parse(value) : dayDate.getTime();
    if (!Number.isFinite(timestamp))
        return null;
    const serial = (timestamp - EXCEL_EPOCH_1900) / DAY_MS + (timestamp >= EXCEL_FICTITIOUS_LEAP_DAY ? 1 : 0);
    return serial >= 0 && serial <= 2_958_465 ? serial : null;
}
/** Convert an Excel 1900-system serial to an ISO date label; serial 60 is 1900-02-29. */
export function excelSerialDateIso(serial) {
    if (!Number.isFinite(serial) || serial < 0 || serial > 2_958_465)
        return null;
    const whole = Math.floor(serial);
    if (whole === 60)
        return '1900-02-29';
    const date = new Date((whole < 60 ? EXCEL_EPOCH_1900 : EXCEL_EPOCH_AFTER_FAKE) + whole * DAY_MS);
    return Number.isFinite(date.getTime()) ? `${String(date.getUTCFullYear()).padStart(4, '0')}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}` : null;
}
/** Convert an Excel serial to a UTC Date for calculations. Serial 60 uses Feb 28 as its physical surrogate. */
export function excelSerialUtcDate(serial) {
    if (!Number.isFinite(serial))
        return new Date(NaN);
    const whole = Math.floor(serial), fraction = serial - whole, iso = excelSerialDateIso(whole);
    if (!iso)
        return new Date(NaN);
    const physical = iso === '1900-02-29' ? '1900-02-28' : iso;
    return new Date(Date.parse(`${physical}T00:00:00Z`) + fraction * DAY_MS);
}
