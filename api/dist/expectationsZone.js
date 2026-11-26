"use strict";
/**
 * Turns a street address into a point so the five-mile check can use a straight line.
 * The address is not written to the application log.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.geocodeAddress = geocodeAddress;
async function geocodeAddress(address) {
    const query = String(address || '').replace(/\s+/g, ' ').trim();
    if (query.length < 5)
        return null;
    const url = 'https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=us&q=' + encodeURIComponent(query);
    const res = await fetch(url, {
        headers: { Accept: 'application/json', 'User-Agent': 'EmmeDashboard/1.0' },
        signal: AbortSignal.timeout(8000),
    });
    if (!res.ok)
        return null;
    const data = await res.json().catch(function () { return []; });
    const first = Array.isArray(data) ? data[0] : null;
    if (!first)
        return null;
    const lat = Number(first.lat);
    const lon = Number(first.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon))
        return null;
    return { lat, lon };
}
