"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.readMapGeometry = readMapGeometry;
exports.geocodeAddress = geocodeAddress;
const expectationsRules_1 = require("./expectationsRules");
/**
 * Turns a street address into a point so the five-mile check can use a straight line.
 * A Google My Map link is read for the drawn area and the pins.
 * The address is not written to the application log.
 */
async function readMapGeometry(mapLink) {
    const mid = (0, expectationsRules_1.mapIdFromLink)(mapLink);
    if (!mid)
        return null;
    const url = 'https://www.google.com/maps/d/kml?mid=' + encodeURIComponent(mid) + '&forcekml=1';
    try {
        const res = await fetch(url, {
            headers: { Accept: 'application/vnd.google-earth.kml+xml, application/xml, text/xml, */*' },
            signal: AbortSignal.timeout(8000),
        });
        if (!res.ok)
            return null;
        const text = await res.text();
        if (!/<kml|<Placemark/i.test(text))
            return null;
        return (0, expectationsRules_1.parseMapKml)(text);
    }
    catch (_) {
        return null;
    }
}
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
