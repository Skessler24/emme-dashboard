/**
 * Staff expectation rules taken from Emme's answers and the 2026 cancellation policy.
 * Pure functions only. No note text and no addresses are handled here.
 */

export const ZONE_MILES = 5;
export const RATIO_LIMIT = 1.8;
export const NOTE_HOURS = 48;
export const TIME_ZONE = 'America/Denver';

export type VisitFrequency = 'weekly' | 'twice-monthly' | 'monthly';

export type CancellationClass = {
  kind: 'rescheduled' | 'no-show' | 'late' | 'advance' | 'other';
  productivity: boolean;
  incident: boolean;
};

export function denverDate(now: Date): string {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(now);
}

export function denverMonth(now: Date): string {
  return denverDate(now).slice(0, 7);
}

export function haversineMiles(lat1: number, lon1: number, lat2: number, lon2: number): number {
  const radius = 3958.8;
  const toRad = function (value: number) { return value * Math.PI / 180; };
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2)
    + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const miles = radius * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(miles * 10) / 10;
}

export function outsideZone(miles: number): boolean {
  return Number.isFinite(miles) && miles > ZONE_MILES;
}

export function mapIdFromLink(link: string): string {
  const match = String(link || '').match(/[?&]mid=([^&#]+)/i);
  if (!match) return '';
  try { return decodeURIComponent(match[1]); } catch (_) { return match[1]; }
}

export function pointInRing(lat: number, lon: number, ring: Array<{ lat: number; lon: number }>): boolean {
  if (!ring || ring.length < 3) return false;
  let inside = false;
  for (let i = 0, j = ring.length - 1; i < ring.length; j = i++) {
    const yi = ring[i].lat;
    const xi = ring[i].lon;
    const yj = ring[j].lat;
    const xj = ring[j].lon;
    const crosses = (yi > lat) !== (yj > lat);
    if (crosses && lon < ((xj - xi) * (lat - yi)) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}

function coordinatePair(token: string): { lat: number; lon: number } | null {
  const parts = token.split(',');
  const lon = Number(parts[0]);
  const lat = Number(parts[1]);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
  return { lat, lon };
}

/** Reads the drawn area and the pins from a Google My Map export. */
export function parseMapKml(kml: string): { polygons: Array<Array<{ lat: number; lon: number }>>; points: Array<{ lat: number; lon: number }> } {
  const polygons: Array<Array<{ lat: number; lon: number }>> = [];
  const points: Array<{ lat: number; lon: number }> = [];
  String(kml || '').split(/<Placemark\b/i).slice(1).forEach(function (chunk) {
    const body = chunk.split(/<\/Placemark>/i)[0] || '';
    const found = body.match(/<coordinates>([\s\S]*?)<\/coordinates>/i);
    if (!found) return;
    const pairs = found[1].trim().split(/\s+/).map(coordinatePair).filter(function (pair): pair is { lat: number; lon: number } {
      return !!pair;
    });
    if (/<Polygon\b/i.test(body) && pairs.length >= 3) polygons.push(pairs);
    else if (/<Point\b/i.test(body) && pairs.length) points.push(pairs[0]);
  });
  return { polygons, points };
}

export function mileageRatio(miles: number, hours: number): number | null {
  if (!Number.isFinite(miles) || !Number.isFinite(hours) || hours <= 0) return null;
  return Math.round((miles / hours) * 100) / 100;
}

export function ratioIsHigh(ratio: number | null): boolean {
  return ratio != null && ratio > RATIO_LIMIT;
}

export function hoursSince(iso: string, now: Date): number | null {
  const ms = Date.parse(String(iso || ''));
  if (Number.isNaN(ms)) return null;
  return Math.round(((now.getTime() - ms) / 36e5) * 10) / 10;
}

export function noteIsLate(iso: string, now: Date): boolean {
  const hours = hoursSince(iso, now);
  return hours != null && hours > NOTE_HOURS;
}

function dateParts(iso: string): { y: number; m: number; d: number } | null {
  const match = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  return { y: Number(match[1]), m: Number(match[2]), d: Number(match[3]) };
}

function isoFromUtc(y: number, m: number, d: number): string {
  return new Date(Date.UTC(y, m - 1, d)).toISOString().slice(0, 10);
}

export function fridayOfWeek(isoDate: string): string {
  const parts = dateParts(isoDate);
  if (!parts) return '';
  const utc = new Date(Date.UTC(parts.y, parts.m - 1, parts.d));
  const day = utc.getUTCDay();
  const shift = day === 6 ? -1 : 5 - day;
  utc.setUTCDate(utc.getUTCDate() + shift);
  return utc.toISOString().slice(0, 10);
}

export function pastEndOfWeek(iso: string, now: Date): boolean {
  const parts = dateParts(iso);
  if (!parts) return false;
  const friday = fridayOfWeek(isoFromUtc(parts.y, parts.m, parts.d));
  if (!friday) return false;
  return denverDate(now) > friday;
}

export function addDays(iso: string, days: number): string {
  const parts = dateParts(iso);
  if (!parts) return '';
  const utc = new Date(Date.UTC(parts.y, parts.m - 1, parts.d));
  utc.setUTCDate(utc.getUTCDate() + days);
  return utc.toISOString().slice(0, 10);
}

export function incidentLimit(frequency: VisitFrequency): number {
  if (frequency === 'monthly') return 1;
  if (frequency === 'twice-monthly') return 2;
  return 4;
}

export function cancellationClass(status: string): CancellationClass {
  const text = String(status || '').toLowerCase();
  if (/reschedul/.test(text)) return { kind: 'rescheduled', productivity: false, incident: false };
  const noShow = /no[-\s]?show|missed/.test(text);
  const late = /late/.test(text);
  const cancel = /cancel|no[-\s]?show|missed|declined/.test(text);
  if (noShow) return { kind: 'no-show', productivity: true, incident: true };
  if (late) return { kind: 'late', productivity: true, incident: true };
  if (cancel) return { kind: 'advance', productivity: true, incident: false };
  return { kind: 'other', productivity: false, incident: false };
}

export function clientKey(name: string): string {
  return String(name || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
}

export const OUTREACH_MESSAGE = 'We have missed several visits. Do you still want services? Has your availability changed, or do you need a new time or a different provider? If we do not hear back, this time will be offered to another family.';

export function isDashboardCopyName(name: string): boolean {
  return /dashboard mileage copy/i.test(String(name || ''));
}
