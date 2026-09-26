// Writes the manifest for the hosted copy of the add-in.
//
// The manifest in the repo root points at https://localhost:3100 and carries the
// development add-in id. The hosted copy has to differ in three ways:
//
//   1. every URL points at the deployed site, under /addin/;
//   2. it carries its own id, so a developer can keep the localhost copy
//      registered at the same time as the installed one without Excel treating
//      them as the same add-in;
//   3. the support and learn-more links point at real pages on the site.
//
// Usage: node scripts/make-manifest.mjs [--base https://example.com] [--out path]
import fs from "node:fs";
import path from "node:path";

export const HOSTED_ID = "fcf30730-1db1-4555-979a-87d048d8c363";
export const DEV_BASE = "https://localhost:3100";
export const DEFAULT_BASE = "https://mex-automate.vercel.app";

/** Reads --name value pairs, falling back to the environment then a default. */
function arg(argv, name, fallback) {
  const at = argv.indexOf(`--${name}`);
  return at >= 0 && argv[at + 1] ? argv[at + 1] : fallback;
}

/** Strips any trailing slash so we can join paths without doubling it. */
export function normalizeBase(base) {
  return base.replace(/\/+$/, "");
}

/**
 * Turns the development manifest into the hosted one. Kept as a pure string
 * transform so it can be tested without touching the file system.
 */
export function hostedManifest(xml, base, { id = HOSTED_ID } = {}) {
  const site = normalizeBase(base);
  let out = xml;

  // Every localhost URL becomes a URL under /addin/ on the site: the add-in's
  // own files (taskpane.html, commands.html, assets/*) are deployed there.
  out = out.split(`${DEV_BASE}/`).join(`${site}/addin/`);

  // The hosted copy is a separate registration from the localhost one.
  out = out.replace(
    /<Id>[^<]+<\/Id>/,
    `<Id>${id}</Id>`
  );

  // Placeholder support links become real pages.
  out = out.replace(
    /(<SupportUrl DefaultValue=")[^"]*(")/,
    `$1${site}/#faq$2`
  );
  out = out.replace(
    /(<bt:Url id="GetStarted.LearnMoreUrl" DefaultValue=")[^"]*(")/,
    `$1${site}/#tools$2`
  );

  if (out.includes("localhost") || out.includes("example.com")) {
    throw new Error("the hosted manifest still contains a placeholder URL");
  }
  return out;
}

if (import.meta.url === `file://${process.argv[1]?.split(path.sep).join("/")}`) {
  const base = arg(process.argv, "base", process.env.MEX_SITE_BASE || DEFAULT_BASE);
  const out = arg(process.argv, "out", "public/manifest.xml");
  const xml = fs.readFileSync("manifest.xml", "utf8");
  const hosted = hostedManifest(xml, base);
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, hosted);
  console.log(`wrote ${out} for ${normalizeBase(base)}`);
}
