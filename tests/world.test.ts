import { beforeEach, describe, expect, it } from 'vitest';
import { config } from '../src/core/config';
import { external, FetchError } from '../src/core/data/external';
import { makeDoc, val } from './helpers';

const GEO = JSON.stringify({
  results: [
    { name: 'Springfield', latitude: 39.8, longitude: -89.64, country: 'United States', country_code: 'US', admin1: 'Illinois' },
    { name: 'Springfield', latitude: 37.21, longitude: -93.29, country: 'United States', country_code: 'US', admin1: 'Missouri' },
  ],
});

const WEATHER = JSON.stringify({
  current: { temperature_2m: 71.5, apparent_temperature: 70, relative_humidity_2m: 40, precipitation: 0, weather_code: 2, wind_speed_10m: 8, is_day: 1 },
  daily: {
    time: ['2026-10-10', '2026-10-11', '2026-10-12'],
    weather_code: [2, 61, 0],
    temperature_2m_max: [75, 62, 68],
    temperature_2m_min: [55, 50, 49],
    precipitation_probability_max: [5, 80, 0],
    precipitation_sum: [0, 0.4, 0],
    sunrise: ['2026-10-10T07:05', '2026-10-11T07:06', '2026-10-12T07:07'],
    sunset: ['2026-10-10T18:30', '2026-10-11T18:28', '2026-10-12T18:27'],
    uv_index_max: [4, 1, 5],
    wind_speed_10m_max: [10, 18, 6],
  },
});

const HOLIDAYS = JSON.stringify([
  { date: '2026-01-01', localName: "New Year's Day", name: "New Year's Day", global: true, types: ['Public'] },
  { date: '2026-07-03', localName: 'Independence Day', name: 'Independence Day', global: true, types: ['Public'] },
]);

const WB_COUNTRIES = JSON.stringify([
  { page: 1 },
  [
    { id: 'JPN', iso2Code: 'JP', name: 'Japan', capitalCity: 'Tokyo', region: { value: 'East Asia & Pacific' }, incomeLevel: { value: 'High income' } },
    { id: 'USA', iso2Code: 'US', name: 'United States', capitalCity: 'Washington D.C.', region: { value: 'North America' }, incomeLevel: { value: 'High income' } },
    { id: 'WLD', iso2Code: '1W', name: 'World', capitalCity: '', region: { value: 'Aggregates' }, incomeLevel: { value: 'Aggregates' } },
  ],
]);

let urls: string[] = [];

beforeEach(() => {
  external.reset();
  urls = [];
  config.units = 'imperial';
  external.fetcher = async (url) => {
    urls.push(url);
    if (url.includes('geocoding-api.open-meteo.com')) return url.includes('Nowhereville') ? JSON.stringify({}) : GEO;
    if (url.includes('api.open-meteo.com/v1/forecast')) return WEATHER;
    if (url.includes('date.nager.at/api/v3/PublicHolidays/2026/US')) return HOLIDAYS;
    if (url.includes('date.nager.at/api/v3/AvailableCountries')) return JSON.stringify([{ countryCode: 'US', name: 'United States' }, { countryCode: 'DE', name: 'Germany' }]);
    if (url.includes('api.worldbank.org/v2/country?')) return WB_COUNTRIES;
    if (url.includes('/country/JPN/indicator/SP.POP.TOTL') && url.includes('mrnev=1')) return JSON.stringify([{}, [{ value: 123975371, date: '2025' }]]);
    if (url.includes('/country/JPN/indicator/SP.POP.TOTL') && url.includes('date=2020')) return JSON.stringify([{}, [{ value: 126261000, date: '2020' }]]);
    throw new FetchError('Not Found', 404);
  };
});

const settle = async () => {
  for (let i = 0; i < 4; i++) await new Promise((r) => setTimeout(r, 60));
};

describe('world data functions', () => {
  it('reads current weather and the forecast', async () => {
    const d = makeDoc({ A1: '=WEATHER("Springfield, MO")', A2: '=WEATHER("Springfield, MO", "Condition", 1)', A3: '=WEATHER("Springfield", "Place")', A4: '=WEATHER("Springfield, MO", "Rain chance", 1)' });
    expect(val(d, 'A1')).toBe('#BUSY!');
    await settle();
    expect(val(d, 'A1')).toBe(71.5);
    expect(val(d, 'A2')).toBe('Light rain');
    expect(val(d, 'A3')).toBe('Springfield, Illinois, United States');
    expect(val(d, 'A4')).toBe(80);
    // "Springfield, MO" picked the Missouri coordinates
    expect(urls.some((u) => u.includes('latitude=37.210'))).toBe(true);
    expect(urls.some((u) => u.includes('temperature_unit=fahrenheit'))).toBe(true);
  });

  it('builds a forecast table and explains unknown places', async () => {
    const d = makeDoc({ A1: '=WEATHER.FORECAST("Springfield", 3)', H1: '=WEATHER("Nowhereville")', H2: '=WEATHER("Springfield", "Sparkle")' });
    await settle();
    expect(val(d, 'A1')).toBe('Date');
    expect(val(d, 'B3')).toBe('Light rain');
    expect(val(d, 'C2')).toBe(75);
    expect(val(d, 'E3')).toBeCloseTo(0.8);
    expect(val(d, 'H1')).toBe('#N/A');
    expect(val(d, 'H2')).toBe('#VALUE!');
  });

  it('lists public holidays by country name or code', async () => {
    const d = makeDoc({ A1: '=HOLIDAYS("US", 2026)', F1: '=HOLIDAYS("United States", 2026)' });
    await settle();
    expect(val(d, 'B2')).toBe("New Year's Day");
    expect(val(d, 'B3')).toBe('Independence Day');
    expect(val(d, 'G3')).toBe('Independence Day');
  });

  it('reads country facts and statistics', async () => {
    const d = makeDoc({ A1: '=COUNTRYDATA("Japan", "Population")', A2: '=COUNTRYDATA("JP", "Capital")', A3: '=COUNTRYDATA("Japan", "Population", 2020)', A4: '=COUNTRYDATA("Atlantis", "Population")', A5: '=COUNTRYDATA("USA", "Region")' });
    await settle();
    expect(val(d, 'A1')).toBe(123975371);
    expect(val(d, 'A2')).toBe('Tokyo');
    expect(val(d, 'A3')).toBe(126261000);
    expect(val(d, 'A4')).toBe('#N/A');
    expect(val(d, 'A5')).toBe('North America');
  });
});
