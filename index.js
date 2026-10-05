#!/usr/bin/env node
/**
 * Pasticcerie Padova Scraper
 * --------------------------
 * Fetches all pasticcerie in Padova via Google Places API,
 * enriches with email by scraping each website's contact page,
 * and exports to a formatted .xlsx file.
 *
 * Requirements:
 *   npm install axios cheerio exceljs
 *
 * Usage:
 *   export GOOGLE_API_KEY="your_key_here"
 *   node pasticcerie_padova.js
 */

const axios = require("axios");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");
require("dotenv").config();

const API_KEY = process.env.GOOGLE_API_KEY || "YOUR_API_KEY_HERE";
const SEARCH_KEYWORD = process.env.SEARCH_KEYWORD || "pasticceria Padova Italia";

// Dynamically generate spreadsheet filename: pasticceria_padova_italia.xlsx
const OUTPUT_FILE = SEARCH_KEYWORD.toLowerCase()
  .replace(/[^a-z0-9]+/g, "_")
  .replace(/^_+|_+$/g, "") + ".xlsx";

const PLACES_NEW_URL = "https://places.googleapis.com/v1/places:searchText";
const FIELD_MASK = "places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,places.websiteUri,places.rating,places.userRatingCount,nextPageToken";

const CONTACT_PATHS = [
  "/contatti", "/contatti/", "/contattaci", "/contattaci/",
  "/contact", "/contact/", "/contact-us", "/contact-us/",
  "/chi-siamo", "/informazioni",
];

const EMAIL_RE = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
const EMAIL_BLACKLIST = ["example.com", "sentry.io", "wixpress.com"];

const HTTP_HEADERS = {
  "User-Agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
    "AppleWebKit/537.36 (KHTML, like Gecko) " +
    "Chrome/120.0.0.0 Safari/537.36",
};

const SOCIAL_PATTERNS = {
  facebook: /(?:facebook\.com|fb\.me|fb\.com)\/[a-zA-Z0-9._\-]+/i,
  instagram: /(?:instagram\.com)\/[a-zA-Z0-9._\-]+/i,
  tripadvisor: /(?:tripadvisor\.[a-z.]+)\/[a-zA-Z0-9._\-]+/i
};

const HIRING_KEYWORDS = [
  "lavora con noi", "lavoraconnoi", "assunzioni", "posizioni aperte", "lavora nel team",
  "careers", "jobs", "hiring", "work with us", "join our team", "cercasi personale"
];

const sleep = (ms) => new Promise((res) => setTimeout(res, ms));

const log = (msg) => console.log(`[INFO] ${msg}`);
const warn = (msg) => console.warn(`[WARN] ${msg}`);


// ---------------------------------------------------------------------------
// Google Places helpers
// ---------------------------------------------------------------------------

async function searchPlaces(query) {
  const results = [];
  let body = {
    textQuery: query,
    languageCode: "it"
  };

  while (true) {
    try {
      const response = await axios.post(PLACES_NEW_URL, body, {
        headers: {
          "Content-Type": "application/json",
          "X-Goog-Api-Key": API_KEY,
          "X-Goog-FieldMask": FIELD_MASK
        }
      });

      const data = response.data;
      results.push(...(data.places ?? []));
      log(`Fetched page, total so far: ${results.length}`);

      if (!data.nextPageToken) break;

      log("nextPageToken present, fetching next page...");
      body = {
        textQuery: query,
        languageCode: "it",
        pageToken: data.nextPageToken
      };
      await sleep(100);
    } catch (err) {
      if (err.response && err.response.data && err.response.data.error) {
        const gErr = err.response.data.error;
        throw new Error(`Places API error: ${gErr.status} (${gErr.code}) — ${gErr.message}`);
      }
      throw err;
    }
  }

  return results;
}


// ---------------------------------------------------------------------------
// Email, Social & Hiring scraping helpers
// ---------------------------------------------------------------------------

function extractInfoFromHtml(html, baseUrl) {
  const $ = cheerio.load(html);
  const emails = new Set();
  let facebook = "";
  let instagram = "";
  let tripadvisor = "";
  let hiringLink = "";
  let hiringTextFound = false;

  // Scan all links
  $("a[href]").each((_, el) => {
    const href = $(el).attr("href")?.trim();
    if (!href) return;

    // Check email mailto
    if (href.startsWith("mailto:")) {
      const addr = href.slice(7).split("?")[0].trim();
      if (addr) emails.add(addr);
      return;
    }

    // Check socials
    if (!facebook && SOCIAL_PATTERNS.facebook.test(href)) {
      facebook = href;
    }
    if (!instagram && SOCIAL_PATTERNS.instagram.test(href)) {
      instagram = href;
    }
    if (!tripadvisor && SOCIAL_PATTERNS.tripadvisor.test(href)) {
      tripadvisor = href;
    }

    // Check hiring links
    const text = $(el).text().toLowerCase();
    const hrefLower = href.toLowerCase();
    const isHiringHref = hrefLower.includes("lavora") || hrefLower.includes("career") || hrefLower.includes("job") || hrefLower.includes("assunzion");
    const isHiringText = HIRING_KEYWORDS.some(k => text.includes(k));

    if (isHiringHref || isHiringText) {
      try {
        hiringLink = href.startsWith("http") ? href : new URL(href, baseUrl).toString();
      } catch {
        hiringLink = href;
      }
    }
  });

  // Check raw text for emails
  const pageText = $.text();
  const matches = pageText.match(EMAIL_RE) ?? [];
  matches.forEach((m) => emails.add(m));

  // Check text content for hiring keywords (like "cercasi personale")
  const pageTextLower = pageText.toLowerCase();
  if (pageTextLower.includes("cercasi personale") || pageTextLower.includes("lavora con noi") || pageTextLower.includes("assumiamo")) {
    hiringTextFound = true;
  }

  return {
    emails: [...emails].filter(e => !EMAIL_BLACKLIST.some(b => e.includes(b))),
    facebook,
    instagram,
    tripadvisor,
    hiringLink,
    hiringTextFound
  };
}

async function enrichFromWebsite(website) {
  const info = {
    email: "",
    facebook: "",
    instagram: "",
    tripadvisor: "",
    hiring: "No"
  };

  if (!website) return info;

  const base = website.startsWith("http") ? website : `https://${website}`;
  const baseClean = base.replace(/\/$/, "");

  // Try home page first
  try {
    const { data, status } = await axios.get(baseClean, {
      headers: HTTP_HEADERS,
      timeout: 8000,
      maxRedirects: 5,
    });

    if (status === 200) {
      const res = extractInfoFromHtml(data, baseClean);
      if (res.emails.length > 0) info.email = res.emails[0];
      if (res.facebook) info.facebook = res.facebook;
      if (res.instagram) info.instagram = res.instagram;
      if (res.tripadvisor) info.tripadvisor = res.tripadvisor;
      if (res.hiringLink) {
        info.hiring = `Sì (${res.hiringLink})`;
      } else if (res.hiringTextFound) {
        info.hiring = "Sì (Testo)";
      }
    }
  } catch (err) {
    // Ignore, try contact pages next
  }

  // If critical fields are missing, try contact subpages
  const needsMore = !info.email || !info.facebook || !info.instagram || !info.tripadvisor;
  if (needsMore) {
    for (const path of CONTACT_PATHS) {
      const contactUrl = baseClean + path;
      try {
        const { data, status } = await axios.get(contactUrl, {
          headers: HTTP_HEADERS,
          timeout: 6000,
          maxRedirects: 5,
        });

        if (status !== 200) continue;

        const res = extractInfoFromHtml(data, contactUrl);
        if (!info.email && res.emails.length > 0) info.email = res.emails[0];
        if (!info.facebook && res.facebook) info.facebook = res.facebook;
        if (!info.instagram && res.instagram) info.instagram = res.instagram;
        if (!info.tripadvisor && res.tripadvisor) info.tripadvisor = res.tripadvisor;
        if (info.hiring === "No") {
          if (res.hiringLink) {
            info.hiring = `Sì (${res.hiringLink})`;
          } else if (res.hiringTextFound) {
            info.hiring = "Sì (Testo)";
          }
        }

        // If we found everything, we can stop early
        if (info.email && info.facebook && info.instagram && info.tripadvisor && info.hiring !== "No") {
          break;
        }
      } catch {
        // Try next path
      }
    }
  }

  // Clean up social urls if they are relative
  const cleanSocialUrl = (url, type) => {
    if (!url) return "";
    if (url.startsWith("http")) return url;
    if (url.startsWith("//")) return `https:${url}`;
    if (url.startsWith("/")) return `https://${type}.com${url}`;
    return `https://${type}.com/${url}`;
  };

  if (info.facebook) info.facebook = cleanSocialUrl(info.facebook, "facebook");
  if (info.instagram) info.instagram = cleanSocialUrl(info.instagram, "instagram");
  if (info.tripadvisor) info.tripadvisor = cleanSocialUrl(info.tripadvisor, "tripadvisor");

  if (info.email) log(`    Found email: ${info.email}`);
  if (info.facebook) log(`    Found Facebook: ${info.facebook}`);
  if (info.instagram) log(`    Found Instagram: ${info.instagram}`);
  if (info.tripadvisor) log(`    Found TripAdvisor: ${info.tripadvisor}`);
  if (info.hiring !== "No") log(`    Hiring status: ${info.hiring}`);

  return info;
}


// ---------------------------------------------------------------------------
// Data collection
// ---------------------------------------------------------------------------

async function collectPassticcerie() {
  log(`Searching Google Places for '${SEARCH_KEYWORD}'...`);
  const searchResults = await searchPlaces(SEARCH_KEYWORD);

  // Deduplicate by place id
  const seen = new Set();
  const unique = searchResults.filter(({ id }) => {
    if (!id || seen.has(id)) return false;
    seen.add(id);
    return true;
  });

  log(`Found ${unique.length} unique places, scraping websites for additional info...`);

  const pasticcerie = [];

  for (let i = 0; i < unique.length; i++) {
    const place = unique[i];
    const name = place.displayName?.text ?? "";
    log(`[${i + 1}/${unique.length}] ${name}`);

    const website = place.websiteUri ?? "";
    log(`  Scraping website: ${website || "(none)"}`);
    const info = await enrichFromWebsite(website);

    pasticcerie.push({
      name,
      address: place.formattedAddress ?? "",
      phone:   place.nationalPhoneNumber ?? "",
      email:   info.email,
      website,
      rating:  place.rating ?? "",
      reviews: place.userRatingCount ?? "",
      facebook: info.facebook,
      instagram: info.instagram,
      tripadvisor: info.tripadvisor,
      hiring:  info.hiring,
    });
  }

  return pasticcerie;
}


// ---------------------------------------------------------------------------
// XLSX export
// ---------------------------------------------------------------------------

async function saveXlsx(pasticcerie, path) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(SEARCH_KEYWORD.slice(0, 30)); // Max 31 chars for sheet name

  const COLUMNS = [
    { header: "Nome",                  key: "name",        width: 30 },
    { header: "Indirizzo",             key: "address",     width: 44 },
    { header: "Telefono",              key: "phone",       width: 20 },
    { header: "Email",                 key: "email",       width: 34 },
    { header: "Sito Web",              key: "website",     width: 32 },
    { header: "Valutazione (★)",       key: "rating",      width: 16 },
    { header: "N° Recensioni",         key: "reviews",     width: 14 },
    { header: "Facebook",              key: "facebook",    width: 34 },
    { header: "Instagram",             key: "instagram",   width: 34 },
    { header: "TripAdvisor",           key: "tripadvisor", width: 34 },
    { header: "Hiring (Assunzioni)",   key: "hiring",      width: 26 },
  ];

  ws.columns = COLUMNS;

  // Style header row
  const headerRow = ws.getRow(1);
  headerRow.height = 24;
  headerRow.eachCell((cell) => {
    cell.font      = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2B5D9C" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border    = {
      top: { style: "thin", color: { argb: "FFCCCCCC" } },
      left: { style: "thin", color: { argb: "FFCCCCCC" } },
      bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
      right: { style: "thin", color: { argb: "FFCCCCCC" } },
    };
  });

  const BORDER = {
    top:    { style: "thin", color: { argb: "FFCCCCCC" } },
    left:   { style: "thin", color: { argb: "FFCCCCCC" } },
    bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
    right:  { style: "thin", color: { argb: "FFCCCCCC" } },
  };

  // Data rows
  pasticcerie.forEach((p, i) => {
    const row = ws.addRow([
      p.name, p.address, p.phone, p.email, p.website, p.rating, p.reviews,
      p.facebook, p.instagram, p.tripadvisor, p.hiring
    ]);
    const isAlt = (i + 2) % 2 === 0;

    row.height = 18;
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: isAlt ? "FFEEF3FA" : "FFFFFFFF" } };
      cell.border = BORDER;
      cell.font   = { name: "Arial", bold: colNum === 1, size: 10 };
      cell.alignment = {
        horizontal: [6, 7, 11].includes(colNum) ? "center" : "left",
        vertical: "middle",
        wrapText: true,
      };
    });
  });

  // Freeze header, auto-filter
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.autoFilter = { from: "A1", to: `K${pasticcerie.length + 1}` };

  await wb.xlsx.writeFile(path);
  log(`Saved ${pasticcerie.length} rows to ${path}`);
}


// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

(async () => {
  if (API_KEY === "YOUR_API_KEY_HERE") {
    console.error(
      "Set your Google API key:\n" +
      "  export GOOGLE_API_KEY='AIza...'\n" +
      "or edit the API_KEY variable at the top of this file."
    );
    process.exit(1);
  }

  const pasticcerie = await collectPassticcerie();
  await saveXlsx(pasticcerie, OUTPUT_FILE);
  console.log(`\nDone — ${pasticcerie.length} records saved to ${OUTPUT_FILE}`);
})();