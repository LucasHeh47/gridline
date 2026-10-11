// Recognizes API keys and similar secrets, so they are never sent to an AI model or shown in an
// error message. Used by both the app and the main process.

const PATTERNS: RegExp[] = [
  /sk-ant-[A-Za-z0-9_-]{16,}/g, // Anthropic
  /\bsk-(?:proj-|svcacct-)?[A-Za-z0-9_-]{32,}/g, // OpenAI and similar
  /\bgh[pousr]_[A-Za-z0-9]{30,}/g, // GitHub
  /\bgithub_pat_[A-Za-z0-9_]{40,}/g,
  /\bAKIA[0-9A-Z]{16}\b/g, // AWS access key id
  /\bAIza[0-9A-Za-z_-]{35}\b/g, // Google
  /\bxox[abprs]-[A-Za-z0-9-]{10,}/g, // Slack
  /\b[rs]k_live_[A-Za-z0-9]{20,}/g, // Stripe
];

export const HIDDEN = '[hidden API key]';

/** True when the text contains something that looks like an API key. */
export function containsSecret(text: string): boolean {
  return PATTERNS.some((re) => {
    re.lastIndex = 0;
    return re.test(text);
  });
}

/** Replaces anything that looks like an API key (and `exact`, when given) with a placeholder. */
export function redactSecrets(text: string, exact?: string | null): string {
  let out = text;
  if (exact && exact.length >= 8) out = out.split(exact).join(HIDDEN);
  for (const re of PATTERNS) out = out.replace(re, HIDDEN);
  return out;
}

/** A key can only be printable characters without spaces (it goes into a request header). */
export function validKeyFormat(key: string): boolean {
  return /^[\x21-\x7e]{20,512}$/.test(key);
}
