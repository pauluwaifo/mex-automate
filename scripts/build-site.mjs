// Assembles everything the deployed site serves into public/.
//
//   public/index.html        the landing page
//   public/addin/*           the built add-in (taskpane.html, bundles, assets)
//   public/manifest.xml      the manifest people install, pointing at /addin/
//   public/install.ps1       the Windows installer
//   public/uninstall.ps1     the Windows uninstaller
//
// One deployment therefore hosts both the page that explains the add-in and the
// add-in itself, which keeps the manifest's URLs on the same origin as the site.
//
// Run `npm run build` first: this script copies dist/, it does not create it.
import fs from "node:fs";
import path from "node:path";
import { hostedManifest, MANIFESTS, normalizeBase, DEFAULT_BASE } from "./make-manifest.mjs";

const OUT = "public";

/** Where the add-in will actually be served from, once deployed. */
function siteBase() {
  if (process.env.MEX_SITE_BASE) return process.env.MEX_SITE_BASE;
  // Vercel exposes the project's production domain to the build, so a rename of
  // the project does not silently leave the manifest pointing somewhere stale.
  if (process.env.VERCEL_PROJECT_PRODUCTION_URL) {
    return `https://${process.env.VERCEL_PROJECT_PRODUCTION_URL}`;
  }
  return DEFAULT_BASE;
}

function copyDir(from, to) {
  fs.mkdirSync(to, { recursive: true });
  for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
    const src = path.join(from, entry.name);
    const dest = path.join(to, entry.name);
    if (entry.isDirectory()) copyDir(src, dest);
    else fs.copyFileSync(src, dest);
  }
}

function count(dir) {
  let n = 0;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    n += entry.isDirectory() ? count(path.join(dir, entry.name)) : 1;
  }
  return n;
}

const base = normalizeBase(siteBase());

if (!fs.existsSync("dist/taskpane.html")) {
  console.error("dist/taskpane.html is missing — run `npm run build` first.");
  process.exit(1);
}

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });

copyDir("docs", OUT);
copyDir("dist", path.join(OUT, "addin"));

// The manifest webpack copied into dist/ is replaced by the generated ones, so
// there is a single source of truth for the hosted URLs and the hosted ids.
// Both add-ins are served from one deployment and share the same bundle; the
// pane looks at which host it is running in and shows the right screen.
let hosted = "";
for (const manifest of MANIFESTS) {
  const built = hostedManifest(fs.readFileSync(manifest.source, "utf8"), base, { id: manifest.id });
  fs.writeFileSync(path.join(OUT, manifest.hosted), built);
  fs.writeFileSync(path.join(OUT, "addin", manifest.hosted), built);
  if (manifest.app === "Excel") hosted = built;
}

for (const script of ["install.ps1", "uninstall.ps1"]) {
  fs.copyFileSync(path.join("install", script), path.join(OUT, script));
}

// The installer and the landing page both need to know where the site lives.
// Rather than hard-code it in two more places, they read it from here.
fs.writeFileSync(
  path.join(OUT, "version.json"),
  JSON.stringify(
    {
      version: /<Version>([^<]+)<\/Version>/.exec(hosted)?.[1] ?? "unknown",
      base,
      manifest: `${base}/manifest.xml`,
      powerpointManifest: `${base}/manifest-powerpoint.xml`,
      built: new Date().toISOString(),
    },
    null,
    2
  ) + "\n"
);

console.log(`built ${OUT}/ for ${base} — ${count(OUT)} files`);
