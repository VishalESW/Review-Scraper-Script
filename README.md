# Review Scraper → Walmart Export

Local app. Paste a product page URL → pull every customer review (rating, name,
title, text, date) → download **Excel (.xlsx)** and a **Walmart-syndication-ready CSV**.

Runs on your machine only. No cloud, no account.

## Run it

Double-click **`Start Scraper.bat`**. A browser tab opens at
`http://127.0.0.1:8765/`. Keep the black window open while you work.

Manual alternative:

```bash
pip install openpyxl
python app.py
```

## Use it

1. Paste a product URL (or click **Use example**).
2. Click **Scrape reviews** — it detects the review app and pulls all pages.
3. Download **Excel** (to read/edit) or **Walmart CSV** (to upload).

## Supported review apps

Detection is marker-scored: the tool reads the page, scores every known widget,
and tries the strongest first (a store can run more than one), falling through
if one has no data.

| App | Data source | Confidence |
|-----|-------------|-----------|
| Okendo | `api.okendo.io` JSON, paginated | ✅ Verified live (colourpop, moonjuice 678, healthfulseasons 944) |
| Yotpo | v3 storefront API (grouped) + v1 fallback | ✅ Verified live (ohnorman 175) |
| Judge.me | `cdn.judge.me/reviews/reviews_for_widget` (myshopify + product id) | ✅ Verified live (beardbrand 832, correct rating spread) |
| REVIEWS.io | `api.reviews.io/product/review` (store + SKU) | Endpoint confirmed; verify-on-demand |
| Shopify Product Reviews | Server-rendered `.spr-review` HTML | Implemented |
| Loox | `loox.io/widget/{clientId}/reviews/{productId}` | ⚠️ clientId is JS-injected on many stores → needs headless |
| Stamped.io | `stamped.io/api/widget` (pubkey + storeUrl) | ⚠️ pubkey is JS-injected on many stores → needs headless |
| **Any other** | JSON-LD `Review` blocks on the page | Partial fallback (many stores expose only the aggregate) |

**Two classes of review app** (learned from live testing):

1. **Server-scrapable** — the key/id sits in the page HTML, or the API is keyed by
   shop domain + product id. Okendo, Yotpo, Judge.me, REVIEWS.io. These the tool
   pulls completely and fast.
2. **Client-injected key** — Loox's `clientId` and Stamped's `pubkey` are written
   into the page by JavaScript after load, so a plain server fetch can't see them.
   For these the tool needs a **headless browser** (Playwright) that loads the page,
   reads the config, then hits the same API. That's an optional add-on — heavier to
   host, but it makes *every* app work. Ask to enable it.

**Detected but not yet wired** (Junip, Fera, Growave, Opinew, Rivyo, Ryviu, Ali
Reviews, Trustpilot, Kudobuzz): the tool names the app and asks for the URL —
send a live product URL and a tested handler gets added. That's exactly how
Okendo and Yotpo were locked in.

> Why not "all apps, guaranteed"? Each app hides its reviews behind a different
> private endpoint. The only way to get **every** review reliably is a tested,
> per-app handler (or a headless browser that paginates the widget — heavy to
> host). JSON-LD is the universal net, but most stores put only the star average
> there, not the review text.

## Hosting this on your website

The scraper runs server-side (no browser needed), so it deploys as a small web
service. To go from this local app to hosted:

1. Swap the stdlib `http.server` for a WSGI/ASGI app (Flask or FastAPI) — the
   scraping functions stay identical.
2. Deploy to Render / Railway / Fly.io / a VPS. Add `openpyxl` to requirements.
3. Add **rate limiting + caching** (per-URL), because a public endpoint will get
   hit hard and review APIs throttle by IP.
4. Add a job queue for large stores (900+ reviews = a few seconds; do it async).

Say the word and this gets refactored into a deployable Flask service with those
pieces. **Legal note:** only syndicate reviews you own (your store's reviews to
your Walmart listing). A public tool that scrapes arbitrary third-party stores at
scale is a different risk profile — worth a conversation before launch.

## Reading the Walmart part first (important)

Walmart has **no "upload reviews" button** in Seller Center. The legitimate route
is Walmart's free **Review Syndication Program**, which ingests a **review CSV**.

Rules that decide eligibility:

- Reviews must be **organically collected on your own site** (your Okendo reviews qualify).
  Reviews from *other* marketplaces cannot be syndicated.
- You must **moderate** reviews for fraud / policy violations.
- You can't already run another syndication service (Bazaarvoice, PowerReviews, etc.).
- Apply here: Walmart Marketplace → Review Syndication Program. Walmart emails you
  their **exact CSV template** on approval — map this file's columns 1:1 to it.

Walmart matches each review to your listing by **GTIN/UPC + Walmart Item ID + SKU**.
The scraper leaves those three columns **blank** because the source page doesn't
contain your Walmart IDs — fill them in before uploading.

## CSV columns produced

`Product Name, SKU, Walmart Item ID, GTIN/UPC, Rating (1-5), Review Title,
Review Text, Reviewer Name, Review Date (YYYY-MM-DD), Verified Purchaser,
Recommended, Source Product URL, Source Site`

The `.xlsx` has a second sheet ("All Fields") with every raw field including
helpful counts and review IDs.

## Note on syndication vs. authenticity

Syndication is meant for **your own** product's reviews moving from your site to
your Walmart listing for the **same product**. Don't paste another seller's or
another product's reviews onto a Walmart listing — that violates Walmart's content
policy and can get the listing or account actioned.
