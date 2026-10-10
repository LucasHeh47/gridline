// Runtime options for the core (parsing, filling, live data, CSV). The UI's Settings page writes
// these; the defaults match Gridline's out-of-the-box behavior, so tests and tools need no setup.

export interface CoreConfig {
  /** How typed dates like 3/4/2025 are read. */
  dateOrder: 'MDY' | 'DMY';
  /** Turn typed dates and times into real dates. */
  recognizeDates: boolean;
  /** "00123" stays text (IDs, ZIP codes) instead of becoming 123. */
  keepLeadingZeros: boolean;
  /** Digit strings longer than 15 digits stay text instead of losing precision. */
  longDigitsAsText: boolean;
  /** Typed web addresses become clickable links. */
  autoLinkUrls: boolean;
  /** Dragging a single number counts up (1, 2, 3…) instead of copying it. */
  fillCountsSingleNumber: boolean;
  /** How many actions Undo remembers. */
  undoLimit: number;
  /** Currency used by CRYPTO when none is given. */
  defaultCurrency: string;
  /** How long current-season NFL data is reused before downloading again (hours). */
  nflCurrentSeasonHours: number;
  /** How long a stock or crypto quote is reused (seconds). */
  quoteMaxAgeSec: number;
  /** Units for WEATHER: °F, mph, inches (imperial) or °C, km/h, mm (metric). */
  units: 'imperial' | 'metric';
  csv: {
    delimiter: ',' | ';' | '\t' | '|';
    /** Write a UTF-8 byte order mark (helps Excel read accented characters). */
    bom: boolean;
    lineEnding: '\r\n' | '\n';
    /** How dates are written: as shown in the cell, or ISO 2025-03-14. */
    dates: 'display' | 'iso';
    /** Convert numbers and dates when opening CSV files (off = everything stays text). */
    detectTypes: boolean;
  };
}

export const config: CoreConfig = {
  dateOrder: 'MDY',
  recognizeDates: true,
  keepLeadingZeros: true,
  longDigitsAsText: true,
  autoLinkUrls: true,
  fillCountsSingleNumber: false,
  undoLimit: 200,
  defaultCurrency: 'USD',
  nflCurrentSeasonHours: 3,
  quoteMaxAgeSec: 120,
  units: 'imperial',
  csv: { delimiter: ',', bom: true, lineEnding: '\r\n', dates: 'display', detectTypes: true },
};
