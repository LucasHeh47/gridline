// World data: weather (Open-Meteo), public holidays (Nager.Date) and country statistics (World Bank).
// All three are free services that don't need an account.
import { config } from '../../config';
import { external, type Resource } from '../../data/external';
import { dateToSerial, serialToParts, todaySerial } from '../../dates';
import { Errors, ErrorValue, isError, type Scalar } from '../../values';
import { argNum, argText } from '../coerce';
import { Matrix } from '../types';
import { define } from './registry';

const DAY = 86400;

// ---------- places ----------

interface Place {
  name: string;
  admin: string;
  country: string;
  lat: number;
  lon: number;
}

interface GeoResult {
  name: string;
  latitude: number;
  longitude: number;
  country?: string;
  country_code?: string;
  admin1?: string;
}

/** "Chicago", "Paris, France", "Springfield, IL" or "41.88, -87.63". */
function placeArg(text: string): { query: string; qualifier: string } | { lat: number; lon: number } | null {
  const t = text.trim();
  if (!t) return null;
  const coords = /^(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)$/.exec(t);
  if (coords) return { lat: +coords[1], lon: +coords[2] };
  const [query, ...rest] = t.split(',').map((s) => s.trim());
  return { query, qualifier: rest.join(' ').toLowerCase() };
}

const US_STATES: Record<string, string> = {
  al: 'alabama', ak: 'alaska', az: 'arizona', ar: 'arkansas', ca: 'california', co: 'colorado', ct: 'connecticut', de: 'delaware', fl: 'florida', ga: 'georgia', hi: 'hawaii', id: 'idaho', il: 'illinois', in: 'indiana', ia: 'iowa', ks: 'kansas', ky: 'kentucky', la: 'louisiana', me: 'maine', md: 'maryland', ma: 'massachusetts', mi: 'michigan', mn: 'minnesota', ms: 'mississippi', mo: 'missouri', mt: 'montana', ne: 'nebraska', nv: 'nevada', nh: 'new hampshire', nj: 'new jersey', nm: 'new mexico', ny: 'new york', nc: 'north carolina', nd: 'north dakota', oh: 'ohio', ok: 'oklahoma', or: 'oregon', pa: 'pennsylvania', ri: 'rhode island', sc: 'south carolina', sd: 'south dakota', tn: 'tennessee', tx: 'texas', ut: 'utah', vt: 'vermont', va: 'virginia', wa: 'washington', wv: 'west virginia', wi: 'wisconsin', wy: 'wyoming', dc: 'district of columbia',
};

function geocode(query: string): Resource<GeoResult[]> {
  return {
    key: `geo:${query.toLowerCase()}`,
    url: `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(query)}&count=20&language=en&format=json`,
    maxAgeSec: 90 * DAY,
    label: `the location "${query}"`,
    parse: (text) => (JSON.parse(text) as { results?: GeoResult[] }).results ?? [],
  };
}

/** Finds a place, or returns an error value (#BUSY! while looking it up). */
function findPlace(text: string): Place | ErrorValue {
  const p = placeArg(text);
  if (!p) return Errors.value('Type a place, like "Chicago", "Paris, France" or "Springfield, IL".');
  if ('lat' in p) return { name: `${p.lat}, ${p.lon}`, admin: '', country: '', lat: p.lat, lon: p.lon };
  const got = external.get(geocode(p.query));
  if (!got.ok) return got.error;
  let list = got.value;
  if (p.qualifier) {
    const q = US_STATES[p.qualifier] ?? p.qualifier;
    list = list.filter((r) => [r.country, r.country_code, r.admin1].some((x) => x && x.toLowerCase().startsWith(q)));
  }
  const r = list[0];
  if (!r) return Errors.na(`Couldn't find a place called "${text.trim()}". Try adding the state or country, like "Springfield, IL".`);
  return { name: r.name, admin: r.admin1 ?? '', country: r.country ?? '', lat: r.latitude, lon: r.longitude };
}

// ---------- weather ----------

interface Weather {
  current: { temperature_2m: number; apparent_temperature: number; relative_humidity_2m: number; precipitation: number; weather_code: number; wind_speed_10m: number; is_day: number };
  daily: { time: string[]; weather_code: number[]; temperature_2m_max: number[]; temperature_2m_min: number[]; precipitation_probability_max: (number | null)[]; precipitation_sum: number[]; sunrise: string[]; sunset: string[]; uv_index_max: (number | null)[]; wind_speed_10m_max: number[] };
}

/** Plain-English names for WMO weather codes. */
const CONDITIONS: Record<number, string> = {
  0: 'Clear',
  1: 'Mostly clear',
  2: 'Partly cloudy',
  3: 'Cloudy',
  45: 'Fog',
  48: 'Freezing fog',
  51: 'Light drizzle',
  53: 'Drizzle',
  55: 'Heavy drizzle',
  56: 'Freezing drizzle',
  57: 'Freezing drizzle',
  61: 'Light rain',
  63: 'Rain',
  65: 'Heavy rain',
  66: 'Freezing rain',
  67: 'Freezing rain',
  71: 'Light snow',
  73: 'Snow',
  75: 'Heavy snow',
  77: 'Snow grains',
  80: 'Showers',
  81: 'Showers',
  82: 'Heavy showers',
  85: 'Snow showers',
  86: 'Heavy snow showers',
  95: 'Thunderstorm',
  96: 'Thunderstorm with hail',
  99: 'Thunderstorm with hail',
};

const condition = (code: number | null | undefined) => (code == null ? '' : CONDITIONS[code] ?? 'Unknown');

export const WEATHER_FIELDS = ['Temperature', 'Feels like', 'Condition', 'High', 'Low', 'Rain chance', 'Precipitation', 'Humidity', 'Wind', 'UV index', 'Sunrise', 'Sunset', 'Place'];

function metric(): boolean {
  return config.units === 'metric';
}

function weatherRes(place: Place): Resource<Weather> {
  const m = metric();
  const q = new URLSearchParams({
    latitude: place.lat.toFixed(3),
    longitude: place.lon.toFixed(3),
    current: 'temperature_2m,apparent_temperature,relative_humidity_2m,precipitation,weather_code,wind_speed_10m,is_day',
    daily: 'weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max,precipitation_sum,sunrise,sunset,uv_index_max,wind_speed_10m_max',
    temperature_unit: m ? 'celsius' : 'fahrenheit',
    wind_speed_unit: m ? 'kmh' : 'mph',
    precipitation_unit: m ? 'mm' : 'inch',
    timezone: 'auto',
    forecast_days: '16',
  });
  return {
    key: `wx:${place.lat.toFixed(3)},${place.lon.toFixed(3)}:${m ? 'm' : 'i'}`,
    url: `https://api.open-meteo.com/v1/forecast?${q}`,
    maxAgeSec: 15 * 60,
    label: `the weather for ${place.name}`,
    parse: (text) => JSON.parse(text) as Weather,
  };
}

/** "2026-03-14T06:58" as a date-time serial. */
function isoToSerial(s: string): number | string {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/.exec(s);
  if (!m) return s;
  return dateToSerial(+m[1], +m[2], +m[3]) + (m[4] ? (+m[4] * 60 + +m[5]) / 1440 : 0);
}

function weatherValue(w: Weather, place: Place, field: string, day: number): Scalar {
  const d = w.daily;
  const i = Math.max(0, Math.min(d.time.length - 1, day));
  const now = day === 0;
  switch (field) {
    case 'temperature':
    case 'temp':
      return now ? w.current.temperature_2m : (d.temperature_2m_max[i] + d.temperature_2m_min[i]) / 2;
    case 'feels like':
    case 'feelslike':
      return now ? w.current.apparent_temperature : Errors.na('"Feels like" is only available for right now (day 0).');
    case 'condition':
    case 'conditions':
    case 'sky':
      return condition(now ? w.current.weather_code : d.weather_code[i]);
    case 'high':
    case 'max':
      return d.temperature_2m_max[i];
    case 'low':
    case 'min':
      return d.temperature_2m_min[i];
    case 'rain chance':
    case 'chance of rain':
    case 'precipitation chance':
      return d.precipitation_probability_max[i] ?? Errors.na('No rain forecast for that day.');
    case 'precipitation':
    case 'rain':
      return now ? w.current.precipitation : d.precipitation_sum[i];
    case 'humidity':
      return now ? w.current.relative_humidity_2m : Errors.na('Humidity is only available for right now (day 0).');
    case 'wind':
    case 'wind speed':
      return now ? w.current.wind_speed_10m : d.wind_speed_10m_max[i];
    case 'uv index':
    case 'uv':
      return d.uv_index_max[i] ?? Errors.na('No UV forecast for that day.');
    case 'sunrise':
      return isoToSerial(d.sunrise[i]);
    case 'sunset':
      return isoToSerial(d.sunset[i]);
    case 'place':
    case 'location':
      return [place.name, place.admin, place.country].filter(Boolean).join(', ');
    default:
      return Errors.value(`"${field}" isn't a weather field. Use one of: ${WEATHER_FIELDS.join(', ')}.`);
  }
}

define(
  'WEATHER',
  'World',
  'location, [field], [day]',
  `Current weather or the forecast for a place. Fields: ${WEATHER_FIELDS.join(', ')}. day 0 is today, 1 is tomorrow, up to 15. Example: =WEATHER("Chicago", "High", 1).`,
  (ctx, [loc, fld, dy]) => {
    const text = argText(ctx, loc);
    if (isError(text)) return text;
    const field = argText(ctx, fld, 'Temperature');
    if (isError(field)) return field;
    const day = argNum(ctx, dy, 0);
    if (isError(day)) return day;
    if (day < 0 || day > 15) return Errors.num('day must be between 0 (today) and 15.');
    const place = findPlace(text);
    if (isError(place)) return place;
    const got = external.get(weatherRes(place));
    if (!got.ok) return got.error;
    return weatherValue(got.value, place, field.trim().toLowerCase(), Math.round(day));
  },
  { external: true, lift: [0, 1, 2], series: { week: 2 }, suggest: { 1: 'weatherField' }, argHelp: { location: 'A city, "City, State", "City, Country" or "latitude, longitude"', day: '0 = today, 1 = tomorrow … 15' } },
);

define(
  'WEATHER.FORECAST',
  'World',
  'location, [days]',
  'A day-by-day forecast table: date, conditions, high, low, chance of rain and precipitation. Example: =WEATHER.FORECAST("Denver", 7).',
  (ctx, [loc, dys]) => {
    const text = argText(ctx, loc);
    if (isError(text)) return text;
    const days = argNum(ctx, dys, 7);
    if (isError(days)) return days;
    if (days < 1 || days > 16) return Errors.num('days must be between 1 and 16.');
    const place = findPlace(text);
    if (isError(place)) return place;
    const got = external.get(weatherRes(place));
    if (!got.ok) return got.error;
    const d = got.value.daily;
    const deg = metric() ? '°C' : '°F';
    const out: Scalar[][] = [['Date', 'Conditions', `High ${deg}`, `Low ${deg}`, 'Rain chance', `Precipitation (${metric() ? 'mm' : 'in'})`]];
    for (let i = 0; i < Math.min(Math.round(days), d.time.length); i++) {
      out.push([isoToSerial(d.time[i]), condition(d.weather_code[i]), d.temperature_2m_max[i], d.temperature_2m_min[i], d.precipitation_probability_max[i] == null ? '' : d.precipitation_probability_max[i]! / 100, d.precipitation_sum[i]]);
    }
    return new Matrix(out, ['ddd, mmm d', undefined, '0', '0', '0%', '0.00']);
  },
  { external: true },
);

// ---------- holidays ----------

interface NagerCountry {
  countryCode: string;
  name: string;
}

interface NagerHoliday {
  date: string;
  localName: string;
  name: string;
  global: boolean;
  types: string[];
}

const countriesRes: Resource<NagerCountry[]> = {
  key: 'holidays:countries',
  url: 'https://date.nager.at/api/v3/AvailableCountries',
  maxAgeSec: 30 * DAY,
  label: 'the list of countries with holiday data',
  parse: (t) => JSON.parse(t) as NagerCountry[],
};

const ALIASES: Record<string, string> = { usa: 'US', 'united states of america': 'US', america: 'US', uk: 'GB', 'great britain': 'GB', england: 'GB', britain: 'GB', 'south korea': 'KR', korea: 'KR', russia: 'RU', holland: 'NL' };

function holidayCountry(input: string): string | ErrorValue {
  const t = input.trim();
  const lower = t.toLowerCase();
  if (ALIASES[lower]) return ALIASES[lower];
  if (/^[A-Za-z]{2}$/.test(t)) return t.toUpperCase();
  const got = external.get(countriesRes);
  if (!got.ok) return got.error;
  const hit = got.value.find((c) => c.name.toLowerCase() === lower) ?? got.value.find((c) => c.name.toLowerCase().startsWith(lower));
  return hit ? hit.countryCode : Errors.na(`No holiday data for "${t}". Use a country name or 2-letter code like "US", "CA", "GB", "DE".`);
}

define(
  'HOLIDAYS',
  'World',
  'country, [year]',
  'Public holidays for a country and year: date, name and local name. Example: =HOLIDAYS("US", 2026) or =HOLIDAYS("Germany").',
  (ctx, [c, y]) => {
    const name = argText(ctx, c, 'US');
    if (isError(name)) return name;
    const year = argNum(ctx, y, serialToParts(todaySerial()).y);
    if (isError(year)) return year;
    if (year < 1975 || year > 2075) return Errors.num('year must be between 1975 and 2075.');
    const code = holidayCountry(name);
    if (isError(code)) return code;
    const got = external.get<NagerHoliday[]>({
      key: `holidays:${code}:${year}`,
      url: `https://date.nager.at/api/v3/PublicHolidays/${Math.round(year)}/${code}`,
      maxAgeSec: 30 * DAY,
      label: `${code} public holidays for ${year}`,
      parse: (t) => JSON.parse(t) as NagerHoliday[],
    });
    if (!got.ok) return got.error;
    const out: Scalar[][] = [['Date', 'Holiday', 'Local name', 'Nationwide']];
    for (const hol of got.value) out.push([isoToSerial(hol.date), hol.name, hol.localName, hol.global ? 'Yes' : 'Regional']);
    return new Matrix(out, ['ddd, mmm d, yyyy']);
  },
  { external: true },
);

// ---------- country statistics ----------

interface WBCountry {
  id: string;
  iso2Code: string;
  name: string;
  capitalCity: string;
  region: { value: string };
  incomeLevel: { value: string };
}

/** Friendly names for World Bank indicators. */
export const COUNTRY_INDICATORS: Record<string, { code: string; label: string }> = {
  population: { code: 'SP.POP.TOTL', label: 'Population' },
  gdp: { code: 'NY.GDP.MKTP.CD', label: 'GDP (US$)' },
  'gdp per person': { code: 'NY.GDP.PCAP.CD', label: 'GDP per person (US$)' },
  'gdp growth': { code: 'NY.GDP.MKTP.KD.ZG', label: 'GDP growth (%)' },
  inflation: { code: 'FP.CPI.TOTL.ZG', label: 'Inflation (%)' },
  unemployment: { code: 'SL.UEM.TOTL.ZS', label: 'Unemployment (%)' },
  'life expectancy': { code: 'SP.DYN.LE00.IN', label: 'Life expectancy (years)' },
  'internet users': { code: 'IT.NET.USER.ZS', label: 'Internet users (% of people)' },
  'land area': { code: 'AG.LND.TOTL.K2', label: 'Land area (km²)' },
  'urban population': { code: 'SP.URB.TOTL.IN.ZS', label: 'Urban population (%)' },
  'co2 per person': { code: 'EN.GHG.CO2.PC.CE.AR5', label: 'CO₂ per person (tonnes)' },
};

export const COUNTRY_FIELDS = ['Population', 'GDP', 'GDP per person', 'GDP growth', 'Inflation', 'Unemployment', 'Life expectancy', 'Internet users', 'Land area', 'Urban population', 'Capital', 'Region', 'Income level', 'Name'];

const wbCountriesRes: Resource<WBCountry[]> = {
  key: 'wb:countries',
  url: 'https://api.worldbank.org/v2/country?format=json&per_page=400',
  maxAgeSec: 30 * DAY,
  label: 'the list of countries',
  parse: (t) => ((JSON.parse(t) as [unknown, WBCountry[]])[1] ?? []).filter((c) => c.region.value !== 'Aggregates'),
};

function findCountry(input: string): WBCountry | ErrorValue {
  const got = external.get(wbCountriesRes);
  if (!got.ok) return got.error;
  const t = input.trim().toLowerCase();
  const alias = ALIASES[t]?.toLowerCase();
  const list = got.value;
  const hit =
    list.find((c) => c.iso2Code.toLowerCase() === (alias ?? t) || c.id.toLowerCase() === t) ??
    list.find((c) => c.name.toLowerCase() === t) ??
    list.find((c) => c.name.toLowerCase().startsWith(t)) ??
    list.find((c) => c.name.toLowerCase().includes(t));
  return hit ?? Errors.na(`Couldn't find a country called "${input.trim()}".`);
}

define(
  'COUNTRYDATA',
  'World',
  'country, field, [year]',
  `Facts and statistics for a country, from the World Bank. Fields: ${COUNTRY_FIELDS.join(', ')}. Without a year you get the most recent value. Example: =COUNTRYDATA("Japan", "Population").`,
  (ctx, [c, f, y]) => {
    const name = argText(ctx, c);
    if (isError(name)) return name;
    const field = argText(ctx, f, 'Population');
    if (isError(field)) return field;
    const country = findCountry(name);
    if (isError(country)) return country;
    const key = field.trim().toLowerCase();
    switch (key) {
      case 'capital':
        return country.capitalCity || Errors.na(`${country.name} has no capital listed.`);
      case 'region':
        return country.region.value;
      case 'income level':
      case 'income':
        return country.incomeLevel.value;
      case 'name':
        return country.name;
    }
    const ind = COUNTRY_INDICATORS[key] ?? Object.values(COUNTRY_INDICATORS).find((x) => x.label.toLowerCase().startsWith(key));
    if (!ind) return Errors.value(`"${field}" isn't a country field. Use one of: ${COUNTRY_FIELDS.join(', ')}.`);
    let year: number | null = null;
    if (y !== undefined) {
      const n = argNum(ctx, y);
      if (isError(n)) return n;
      year = Math.round(n);
    }
    const got = external.get<{ value: number | null; date: string } | null>({
      key: `wb:${country.id}:${ind.code}:${year ?? 'latest'}`,
      url: `https://api.worldbank.org/v2/country/${country.id}/indicator/${ind.code}?format=json&${year ? `date=${year}` : 'mrnev=1'}`,
      maxAgeSec: 7 * DAY,
      label: `${ind.label} for ${country.name}`,
      parse: (t) => {
        const rows = (JSON.parse(t) as [unknown, { value: number | null; date: string }[] | null])[1];
        return rows?.find((r) => r.value !== null) ?? null;
      },
    });
    if (!got.ok) return got.error;
    if (!got.value || got.value.value === null) return Errors.na(`No ${ind.label} for ${country.name}${year ? ` in ${year}` : ''}.`);
    return got.value.value;
  },
  { external: true, lift: [0, 1, 2], series: { season: 2 }, suggest: { 1: 'countryField' } },
);
