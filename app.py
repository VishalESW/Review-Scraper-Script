#!/usr/bin/env python3
"""
Review Scraper -> Walmart-ready export
--------------------------------------
Local web app. Paste a product page URL, scrape all customer reviews
(rating + reviewer name + title + text + date), export to Excel (.xlsx)
and to a Walmart-Review-Syndication-ready CSV.

No cloud, no accounts. Runs on your machine.

Supported review apps:
  - Okendo          (fully tested on healthfulseasons.com)
  - Yotpo           (public widget API)
  - Judge.me        (public widget endpoint)
  - Stamped.io      (public widget API)
  - Loox            (detected, guidance only)

Only one external dependency: openpyxl (for .xlsx). CSV/HTTP use the stdlib.
"""

import csv
import io
import json
import os
import re
import sys
import threading
import time
import uuid
import webbrowser
from html import unescape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, quote
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    HAVE_XLSX = True
except Exception:
    HAVE_XLSX = False

PORT = int(os.environ.get("PORT", "8765"))
HOST = os.environ.get("HOST", "127.0.0.1")  # set HOST=0.0.0.0 when hosting (Docker/Dokploy)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

# CORS: set ALLOWED_ORIGIN to your site (e.g. https://esellerworld.com) when embedding
# the page on a different domain than the backend. "*" allows any origin.
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "*")

# Rate limiting for the expensive endpoints (/scrape, /support), per client IP.
RATE_MAX = int(os.environ.get("RATE_LIMIT_MAX", "20"))       # requests...
RATE_WINDOW = int(os.environ.get("RATE_LIMIT_WINDOW", "60"))  # ...per this many seconds
_RL_LOCK = threading.Lock()
_RL = {}  # ip -> [timestamps]


def rate_ok(ip):
    if RATE_MAX <= 0:            # RATE_LIMIT_MAX=0 disables rate limiting entirely
        return True
    now = time.time()
    with _RL_LOCK:
        hits = [t for t in _RL.get(ip, []) if now - t < RATE_WINDOW]
        if len(hits) >= RATE_MAX:
            _RL[ip] = hits
            return False
        hits.append(now)
        _RL[ip] = hits
        # opportunistic cleanup so the dict can't grow forever
        if len(_RL) > 5000:
            for k in [k for k, v in _RL.items() if not any(now - t < RATE_WINDOW for t in v)]:
                _RL.pop(k, None)
        return True


# token -> scrape result (kept in memory for the download buttons)
RESULTS = {}


def send_support_email(rec):
    """Email a support request when SMTP_* env vars are set. Returns (sent, error).
    Configure on the host: SMTP_HOST, SMTP_PORT (587 STARTTLS / 465 SSL), SMTP_USER,
    SMTP_PASS, SUPPORT_TO, optional SUPPORT_FROM."""
    host = os.environ.get("SMTP_HOST")
    if not host:
        return False, None
    port = int(os.environ.get("SMTP_PORT", "587"))
    user = os.environ.get("SMTP_USER", "")
    pw = os.environ.get("SMTP_PASS", "")
    to_addr = os.environ.get("SUPPORT_TO", "contact@esellerworld.com")
    from_addr = os.environ.get("SUPPORT_FROM", user or to_addr)
    try:
        import smtplib
        from email.message import EmailMessage
        msg = EmailMessage()
        msg["Subject"] = f"[Review Tool] Issue from {rec['name']}"
        msg["From"] = from_addr
        msg["To"] = to_addr
        msg.set_content(
            f"Time: {rec['time']}\nName: {rec['name']}\nURL:  {rec['url']}\n\n"
            f"{rec['description']}\n")
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=20) as s:
                if user:
                    s.login(user, pw)
                s.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=20) as s:
                s.starttls()
                if user:
                    s.login(user, pw)
                s.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)


# --------------------------------------------------------------------------
# HTTP helpers (stdlib only)
# --------------------------------------------------------------------------
def http_get(url, as_json=False, timeout=30):
    req = Request(url, headers={"User-Agent": UA,
                                "Accept": "application/json, text/html, */*"})
    with urlopen(req, timeout=timeout) as r:
        raw = r.read()
        final = r.geturl()
    text = raw.decode("utf-8", "replace")
    if as_json:
        return json.loads(text), final
    return text, final


def product_handle(path):
    m = re.search(r"/products/([^/?#]+)", path)
    return m.group(1) if m else ""


# --------------------------------------------------------------------------
# Common review record
# --------------------------------------------------------------------------
def make_review(rating, title, body, name, date_iso, verified, recommended, extra=None):
    d = ""
    if date_iso:
        d = str(date_iso)[:10]  # YYYY-MM-DD
    return {
        "rating": rating if rating is not None else "",
        "review_title": (title or "").strip(),
        "review_text": (body or "").strip(),
        "reviewer_name": (name or "").strip(),
        "review_date": d,
        "verified_purchaser": "Yes" if verified else "No",
        "recommended": "Yes" if recommended else ("No" if recommended is False else ""),
        "extra": extra or {},
    }


# --------------------------------------------------------------------------
# Provider: OKENDO
# Works on classic stores (inline settings) AND modern JS-rendered themes, where
# the subscriberId is no longer in the page HTML but is baked into the referenced
# okendo widget bundle (e.g. media-dynamic.okendo.io/images/{subscriberId}/...),
# and the product id comes from Shopify's /products/{handle}.js.
# --------------------------------------------------------------------------
# A UUID that follows an okendo.io path segment is the store's subscriberId.
OKENDO_SUB_RE = re.compile(
    r'okendo\.io/(?:[a-z0-9_-]+/)*?'
    r'([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})', re.I)


def _okendo_subscriber_ids(html):
    """Resolve the Okendo subscriberId: inline first, else from the widget bundle(s)."""
    ids = []
    m = re.search(r'"subscriberId"\s*:\s*"([0-9a-fA-F-]{36})"', html)
    if m:
        ids.append(m.group(1))
    ids += OKENDO_SUB_RE.findall(html)              # api./media-dynamic.okendo.io/.../{id}
    if not ids:
        # JS-rendered theme: the id lives inside the referenced okendo *.js bundle(s).
        assets = re.findall(r'["\']((?://|https?://)[^"\']*okendo[^"\']*\.js)', html, re.I)
        for a in list(dict.fromkeys(assets))[:4]:
            url = ("https:" + a) if a.startswith("//") else a
            try:
                js, _ = http_get(url)
            except Exception:
                continue
            ids += OKENDO_SUB_RE.findall(js)
            if ids:
                break
    return _uniq(ids)


def _okendo_product_id(html, page_url):
    """Resolve the Okendo product id (shopify-<numeric id>)."""
    for pat in (r'data-oke-[\w-]*product-id=["\']([^"\']+)["\']',
                r'"productId"\s*:\s*"(shopify-\d+)"',
                r'okendoProductId["\']?\s*[:=]\s*["\']([^"\']+)["\']'):
        m = re.search(pat, html)
        if m:
            return m.group(1)
    prod = shopify_product(page_url)               # /products/{handle}.js
    if prod.get("id"):
        return "shopify-" + prod["id"]
    # Clean/alias product URL (no /products/ in path): use the canonical product handle.
    host = urlparse(page_url).netloc
    cm = (re.search(r'rel=["\']canonical["\'][^>]+/products/([a-z0-9][a-z0-9-]*)', html, re.I)
          or re.search(r'og:url["\'][^>]+content=["\'][^"\']*/products/([a-z0-9][a-z0-9-]*)',
                       html, re.I))
    if cm:
        p = shopify_product(f"https://{host}/products/{cm.group(1)}")
        if p.get("id"):
            return "shopify-" + p["id"]
    return None


def okendo_scrape(html, host, progress, page_url):
    pid = _okendo_product_id(html, page_url)
    if not pid:
        raise RuntimeError("OKENDO_NO_PRODUCT")    # sentinel -> page-vs-product guidance
    subs = _okendo_subscriber_ids(html)
    if not subs:
        raise RuntimeError("Okendo detected but could not resolve the store's subscriber id.")

    base = "https://api.okendo.io/v1"
    sub = subs[0]
    if len(subs) > 1:                              # pick the id that actually returns reviews
        for c in subs:
            try:
                j, _ = http_get(f"{base}/stores/{c}/products/{quote(pid)}/reviews?limit=1",
                                as_json=True)
                if j.get("reviews"):
                    sub = c
                    break
            except Exception:
                continue

    url = f"{base}/stores/{sub}/products/{quote(pid)}/reviews?limit=100&orderBy=date%20desc"
    reviews, product_name, pages = [], "", 0
    while url and pages < 200:
        j, _ = http_get(url, as_json=True)
        for rv in j.get("reviews", []):
            if not product_name:
                product_name = rv.get("productName", "")
            reviews.append(make_review(
                rv.get("rating"),
                rv.get("title"),
                rv.get("body"),
                (rv.get("reviewer") or {}).get("displayName"),
                rv.get("dateCreated"),
                (rv.get("reviewer") or {}).get("isVerified"),
                rv.get("isRecommended"),
                {"helpfulCount": rv.get("helpfulCount"),
                 "isIncentivized": rv.get("isIncentivized"),
                 "reviewId": rv.get("reviewId")},
            ))
        pages += 1
        progress(len(reviews))
        nu = j.get("nextUrl")
        if not nu:
            break
        url = nu if nu.startswith("http") else base + nu
    return reviews, product_name, "Okendo"


# --------------------------------------------------------------------------
# Provider: YOTPO  (public widget API)
# --------------------------------------------------------------------------
def _uniq(seq):
    seen, out = set(), []
    for x in seq:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


def yotpo_scrape(html, host, progress):
    # Collect candidate app keys (classic staticw2 + new widgets-repository loader token)
    keys = _uniq(
        re.findall(r'staticw2\.yotpo\.com/([\w-]{8,})', html)
        + re.findall(r'cdn-widgetsrepository\.yotpo\.com/v1/loader/([\w-]{15,})', html)
        + re.findall(r'app_key["\']?\s*[:=]\s*["\']([\w-]{8,})["\']', html, re.I)
    )
    # Collect candidate product ids (Yotpo product id or Shopify product id)
    pids = _uniq(
        re.findall(r'data-yotpo-product-id=["\']?(\d{4,})', html)
        + re.findall(r'data-product-id=["\'](\d{4,})["\']', html)
        + re.findall(r'"productId"\s*:\s*"?(\d{5,})"?', html)
        + re.findall(r'"product"\s*:\s*\{\s*"id"\s*:\s*(\d{6,})', html)
        + re.findall(r'product_id["\'\s:=]+(\d{10,})', html)
    )
    if not keys:
        raise RuntimeError("Yotpo detected but no app_key found on page.")
    if not pids:
        raise RuntimeError("Yotpo detected but no product id found on page.")

    # Probe (key, pid) combos on BOTH APIs; keep the one with the highest total.
    # v3 storefront returns grouped reviews (matches the on-page count); v1 often
    # returns only the single product's reviews, which can be fewer.
    best = {"total": -1, "key": None, "pid": None, "api": None}

    def probe(api, k, pd):
        if api == "v3":
            u = (f"https://api-cdn.yotpo.com/v3/storefront/store/{k}/product/{pd}"
                 f"/reviews?perPage=1&page=1")
            j, _ = http_get(u, as_json=True)
            return (j.get("pagination", {}) or {}).get("total", 0)
        u = (f"https://api-cdn.yotpo.com/v1/widget/{k}/products/{pd}"
             f"/reviews.json?per_page=1&page=1")
        j, _ = http_get(u, as_json=True)
        return (j.get("response", {}).get("pagination", {}) or {}).get("total", 0)

    for api in ("v3", "v1"):
        for k in keys:
            for pd in pids:
                try:
                    total = probe(api, k, pd)
                except Exception:
                    continue
                if total > best["total"]:
                    best = {"total": total, "key": k, "pid": pd, "api": api}
    if not best["key"] or best["total"] <= 0:
        raise RuntimeError("Yotpo detected but no reviews returned for any app_key/product-id combo.")

    app_key, pid, api = best["key"], best["pid"], best["api"]
    reviews, product_name, page = [], "", 1

    if api == "v3":
        while page < 400:
            url = (f"https://api-cdn.yotpo.com/v3/storefront/store/{app_key}/product/{pid}"
                   f"/reviews?perPage=100&page={page}&sortBy=date&sortDirection=desc")
            j, _ = http_get(url, as_json=True)
            batch = j.get("reviews", []) or []
            if not product_name:
                product_name = (j.get("products", [{}]) or [{}])[0].get("name", "")
            if not batch:
                break
            for rv in batch:
                u = rv.get("user", {}) or {}
                reviews.append(make_review(
                    rv.get("score"), rv.get("title"), rv.get("content"),
                    u.get("displayName") or rv.get("name"),
                    rv.get("createdAt"),
                    rv.get("verifiedBuyer"), None,
                    {"reviewId": rv.get("id"),
                     "isIncentivized": rv.get("isIncentivized"),
                     "country": (rv.get("country") or {}).get("code")},
                ))
            progress(len(reviews))
            total = (j.get("pagination", {}) or {}).get("total")
            if total and len(reviews) >= total:
                break
            page += 1
        return reviews, product_name, "Yotpo"

    while page < 400:
        url = (f"https://api-cdn.yotpo.com/v1/widget/{app_key}/products/{pid}/reviews.json"
               f"?per_page=100&page={page}")
        j, _ = http_get(url, as_json=True)
        resp = j.get("response", {})
        batch = resp.get("reviews", []) or []
        if not product_name:
            product_name = (resp.get("products", [{}]) or [{}])[0].get("name", "")
        if not batch:
            break
        for rv in batch:
            u = rv.get("user", {}) or {}
            reviews.append(make_review(
                rv.get("score"), rv.get("title"), rv.get("content"),
                u.get("display_name") or rv.get("name"),
                rv.get("created_at"),
                rv.get("verified_buyer"), None,
                {"reviewId": rv.get("id")},
            ))
        progress(len(reviews))
        page += 1
        total = resp.get("pagination", {}).get("total")
        if total and len(reviews) >= total:
            break
    return reviews, product_name, "Yotpo"


# --------------------------------------------------------------------------
# Provider: JUDGE.ME  (public widget endpoint returns HTML fragments)
# --------------------------------------------------------------------------
def judgeme_scrape(html, host, progress, page_url):
    # Judge.me serves paginated HTML fragments from cdn.judge.me keyed by the
    # store's permanent myshopify domain + Shopify product id.
    m = re.search(r'([\w-]+\.myshopify\.com)', html)
    shop = m.group(1) if m else host
    prod = shopify_product(page_url)
    pid = prod.get("id")
    if not pid:
        m2 = (re.search(r'jdgm-widget[^>]+data-id=["\'](\d{6,})', html)
              or re.search(r'data-product-id=["\'](\d{6,})', html)
              or re.search(r'"product"\s*:\s*\{\s*"id"\s*:\s*(\d+)', html))
        pid = m2.group(1) if m2 else None
    if not pid:
        raise RuntimeError("Judge.me detected but Shopify product id not found.")

    reviews, name, page, total = [], prod.get("title", ""), 1, None
    while page < 500:
        # per_page is capped server-side (~30); requesting more desyncs paging.
        api = ("https://cdn.judge.me/reviews/reviews_for_widget"
               f"?url={quote(shop)}&shop_domain={quote(shop)}&platform=shopify"
               f"&page={page}&per_page=30&product_id={pid}")
        j, _ = http_get(api, as_json=True)
        if total is None:
            total = j.get("total_count")
        frag = j.get("html", "") or ""
        blocks = re.split(r"class=['\"]jdgm-rev jdgm-divider", frag)[1:]
        if not blocks:
            if page == 1:
                raise RuntimeError("Judge.me is installed but this product has no reviews yet.")
            break
        before = len(reviews)
        for b in blocks:
            rat = re.search(r"jdgm-rev__rating['\"\s][^>]*data-score=['\"](\d)", b) or \
                  re.search(r"data-score=['\"](\d)", b)
            name_m = re.search(r"class=['\"]jdgm-rev__author['\"][^>]*>([^<]+)", b)
            title = re.search(r"class=['\"]jdgm-rev__title['\"][^>]*>(.*?)</", b, re.S)
            body = re.search(r"class=['\"]jdgm-rev__body['\"][^>]*>(.*?)</div", b, re.S)
            date = re.search(r"datetime=['\"]([^'\"]+)", b)
            reviews.append(make_review(
                int(rat.group(1)) if rat else None,
                strip_tags(title.group(1)) if title else "",
                strip_tags(body.group(1)) if body else "",
                strip_tags(name_m.group(1)) if name_m else "",
                date.group(1) if date else "",
                "data-verified-buyer='true'" in b or 'data-verified-buyer="true"' in b, None,
                {"reviewId": (re.search(r"data-review-id=['\"]([^'\"]+)", b) or [None, None])[1]},
            ))
        progress(len(reviews))
        if len(reviews) == before:      # page added nothing new -> stop
            break
        if total and len(reviews) >= total:
            break
        page += 1
    if not reviews:
        raise RuntimeError("Judge.me detected but no reviews parsed from the widget feed.")
    return reviews, name, "Judge.me"


# --------------------------------------------------------------------------
# Provider: STAMPED.IO  (public widget API)
# --------------------------------------------------------------------------
def stamped_scrape(html, host, progress, page_url):
    # Live endpoint: stamped.io/api/widget?productId&apiKey(pubkey-…)&storeUrl(myshopify)
    # The public key is often injected by client-side JS; when it isn't in the HTML
    # this raises with a clear message (that store needs the browser/headless path).
    key = (re.search(r'(pubkey-[\w]+)', html)
           or re.search(r'data-api-key=["\']([^"\']+)', html)
           or re.search(r'apiKey["\']?\s*[:=]\s*["\']([\w=-]+)', html))
    ms = re.search(r'([\w-]+\.myshopify\.com)', html)
    prod = shopify_product(page_url)
    pid = prod.get("id") or (re.search(r'data-product-id=["\'](\d+)', html) or [None, None])[1]
    if not key:
        raise RuntimeError("Stamped detected but its public key (pubkey-…) isn't in the page "
                           "HTML (injected by JS on this store). Send me the URL for the "
                           "headless path.")
    if not pid:
        raise RuntimeError("Stamped detected but product id not found.")
    store_url = ms.group(1) if ms else host
    key = key.group(1)

    reviews, name, page, total = [], prod.get("title", ""), 1, None
    while page < 300:
        url = (f"https://stamped.io/api/widget?productId={pid}&apiKey={quote(key)}"
               f"&storeUrl={quote(store_url)}&take=25&page={page}"
               f"&sortReviews=recent&widgetLanguage=en")
        j, _ = http_get(url, as_json=True)
        if total is None:
            total = j.get("count") or 0
        # Reviews may arrive as JSON (data/results) or embedded in the widget HTML.
        batch = j.get("data") or j.get("results") or []
        before = len(reviews)
        if batch:
            for rv in batch:
                reviews.append(make_review(
                    rv.get("reviewRating") or rv.get("rating"),
                    rv.get("reviewTitle") or rv.get("title"),
                    rv.get("reviewMessage") or rv.get("body"),
                    (rv.get("author") or rv.get("reviewUserName") or "").strip(),
                    rv.get("reviewDate") or rv.get("dateCreated"),
                    rv.get("reviewVerifiedType") == 2, None,
                ))
        else:
            frag = j.get("widget") if isinstance(j.get("widget"), str) else ""
            for b in re.split(r'stamped-review-content\b', frag)[1:]:
                b = b[:3000]
                rat = re.search(r'data-rating=["\'](\d)', b) or re.search(r'Rated (\d)', b)
                title = re.search(r'stamped-review-header-title[^>]*>(.*?)<', b, re.S)
                body = re.search(r'stamped-review-content-body[^>]*>(.*?)</', b, re.S)
                auth = re.search(r'stamped-review-header-byline[^>]*>(.*?)<', b, re.S)
                reviews.append(make_review(
                    int(rat.group(1)) if rat else None,
                    strip_tags(title.group(1)) if title else "",
                    strip_tags(body.group(1)) if body else "",
                    strip_tags(auth.group(1)) if auth else "", "", None, None,
                ))
        progress(len(reviews))
        if len(reviews) == before:
            break
        if total and len(reviews) >= total:
            break
        page += 1
    if not reviews:
        raise RuntimeError("Stamped detected but no reviews returned (product may have none, "
                           "or reviews render client-side here).")
    return reviews, name, "Stamped.io"


# --------------------------------------------------------------------------
# Shopify helper: product .js gives id, title, and variant SKUs (no CORS server-side)
# --------------------------------------------------------------------------
def shopify_product(page_url):
    handle = product_handle(urlparse(page_url).path)
    host = urlparse(page_url).netloc
    if not handle:
        return {}
    try:
        j, _ = http_get(f"https://{host}/products/{handle}.js", as_json=True)
        return {
            "id": str(j.get("id", "")),
            "title": j.get("title", ""),
            "skus": _uniq([str(v.get("sku")) for v in j.get("variants", []) if v.get("sku")]),
        }
    except Exception:
        return {}


def shopify_resolve_handle(host, alias):
    """Map a short/alias product handle to the real one via Shopify predictive search.
    Lets clean URLs like /products/ashwagandha-capsules (real handle is longer) work."""
    if not alias:
        return None
    q = alias.replace("-", " ").strip()
    try:
        j, _ = http_get(f"https://{host}/search/suggest.json?q={quote(q)}"
                        f"&resource[type]=product&resource[limit]=10", as_json=True)
        prods = (((j.get("resources") or {}).get("results") or {}).get("products") or [])
    except Exception:
        return None
    handles = [p.get("handle", "") for p in prods if p.get("handle")]
    if not handles:
        return None
    a = alias.lower()
    for h in handles:                       # exact handle
        if h.lower() == a:
            return h
    for h in handles:                       # one is a prefix of the other
        if h.lower().startswith(a) or a.startswith(h.lower()):
            return h
    first = a.split("-")[0]
    for h in handles:                       # share the first word
        if h.lower().startswith(first):
            return h
    return handles[0]                       # else the most relevant match


def strip_tags(s):
    return unescape(re.sub(r"<[^>]+>", "", s or "")).strip()


# --------------------------------------------------------------------------
# Provider: REVIEWS.io  (public product review API)
# --------------------------------------------------------------------------
def reviewsio_scrape(html, host, progress, page_url):
    store = None
    for pat in (r'widget\.reviews\.io/[^"\']*?store[=/]([\w.\-]+)',
                r'data-store-name=["\']([\w.\-]+)',
                r'reviews\.io[^{]{0,80}?store["\']?\s*:\s*["\']([\w.\-]+)',
                r'"store"\s*:\s*"([\w.\-]+)"'):
        m = re.search(pat, html, re.I)
        if m:
            store = m.group(1)
            break
    if not store:
        raise RuntimeError("REVIEWS.io detected but store id not found on page.")

    prod = shopify_product(page_url)
    skus = prod.get("skus", [])
    ids = ";".join(skus) if skus else product_handle(urlparse(page_url).path)

    reviews, name, page = [], prod.get("title", ""), 1
    while page < 200:
        url = (f"https://api.reviews.io/product/review?store={quote(store)}"
               f"&sku={quote(ids)}&per_page=50&page={page}")
        j, _ = http_get(url, as_json=True)
        block = j.get("reviews", {}) or {}
        data = block.get("data", []) or []
        if not data:
            break
        for rv in data:
            rr = rv.get("reviewer", {}) or {}
            nm = (f"{rr.get('first_name','')} {rr.get('last_name','')}").strip()
            reviews.append(make_review(
                rv.get("rating"), rv.get("title"), rv.get("review"),
                nm, rv.get("date_created"),
                rr.get("verified_buyer"), None,
                {"reviewId": rv.get("store_review_id") or rv.get("id")},
            ))
        progress(len(reviews))
        total_pages = block.get("last_page") or 1
        if page >= total_pages:
            break
        page += 1
    if not reviews:
        raise RuntimeError("REVIEWS.io detected but no reviews returned for this product's SKU.")
    return reviews, name, "REVIEWS.io"


# --------------------------------------------------------------------------
# Provider: LOOX  (widget endpoint returns an HTML doc with review blocks)
# --------------------------------------------------------------------------
def loox_scrape(html, host, progress, page_url):
    cid = None
    for pat in (r'loox\.io/widget/([\w-]{5,})/', r'clientId["\']?\s*:\s*["\']([\w-]{5,})'):
        m = re.search(pat, html)
        if m:
            cid = m.group(1)
            break
    prod = shopify_product(page_url)
    pid = (re.search(r'id=["\']looxReviews["\'][^>]*data-product-id=["\'](\d+)', html)
           or re.search(r'data-product-id=["\'](\d+)["\'][^>]*class=["\']loox', html)
           or [None, prod.get("id")])[1] or prod.get("id")
    if not pid:
        pid = (re.search(r'"product"\s*:\s*\{\s*"id"\s*:\s*(\d+)', html)
               or re.search(r'LOOX\.productId\s*=\s*["\']?(\d+)', html) or [None, None])[1]
    if not cid:
        raise RuntimeError("Loox detected but its clientId isn't in the page HTML "
                           "(injected by JS on this store) — needs the headless path.")
    if not pid:
        raise RuntimeError("Loox detected but productId not found on page.")

    reviews, page = [], 1
    while page < 200:
        url = f"https://loox.io/widget/{cid}/reviews/{pid}?limit=100&page={page}"
        doc, _ = http_get(url)
        blocks = re.split(r'loox-review\b', doc)[1:]
        if not blocks:
            break
        got = 0
        for b in blocks[:120]:
            b = b[:4000]
            rat = (re.search(r'data-rating=["\'](\d(?:\.\d)?)', b)
                   or re.search(r'(\d)\s*out of\s*5', b))
            title = re.search(r'block title[^>]*>(.*?)</', b, re.S)
            body = re.search(r'block text[^>]*>(.*?)</', b, re.S) or \
                   re.search(r'review-text[^>]*>(.*?)</', b, re.S)
            owner = re.search(r'owner[^>]*>(.*?)</', b, re.S) or \
                    re.search(r'review-author[^>]*>(.*?)</', b, re.S)
            date = re.search(r'datetime=["\']([^"\']+)', b) or \
                   re.search(r'data-date=["\']([^"\']+)', b)
            reviews.append(make_review(
                float(rat.group(1)) if rat else None,
                strip_tags(title.group(1)) if title else "",
                strip_tags(body.group(1)) if body else "",
                strip_tags(owner.group(1)) if owner else "",
                date.group(1) if date else "",
                None, None,
            ))
            got += 1
        progress(len(reviews))
        if got < 100:
            break
        page += 1
    if not reviews:
        raise RuntimeError("Loox detected but no reviews parsed (widget markup may differ) "
                           "— send this URL so the parser can be tuned.")
    return reviews, prod.get("title", ""), "Loox"


# --------------------------------------------------------------------------
# Provider: SHOPIFY PRODUCT REVIEWS (free app, server-rendered .spr-review)
# --------------------------------------------------------------------------
def spr_scrape(html, host, progress, page_url):
    blocks = re.split(r'spr-review\b', html)[1:]
    reviews = []
    for b in blocks:
        b = b[:4000]
        rat = re.search(r'spr-starrating[^>]*aria-label=["\'](\d(?:\.\d)?)', b) or \
              re.search(r'Rated (\d(?:\.\d)?) out of 5', b)
        stars = None
        if not rat:
            stars = b.count("spr-icon-star ") - b.count("spr-icon-star-empty")
        title = re.search(r'spr-review-header-title[^>]*>(.*?)</', b, re.S)
        body = re.search(r'spr-review-content-body[^>]*>(.*?)</div', b, re.S)
        author = re.search(r'spr-review-header-byline[^>]*>(.*?)</', b, re.S)
        auth_txt = strip_tags(author.group(1)) if author else ""
        nm = re.sub(r'\s*on\s+.*$', '', auth_txt).strip()
        dt = re.search(r'on\s+([A-Za-z]+ \d{1,2},? \d{4})', auth_txt)
        reviews.append(make_review(
            float(rat.group(1)) if rat else (stars if stars and 1 <= stars <= 5 else None),
            strip_tags(title.group(1)) if title else "",
            strip_tags(body.group(1)) if body else "",
            nm, dt.group(1) if dt else "",
            None, None,
        ))
    progress(len(reviews))
    if not reviews:
        raise RuntimeError("Shopify Product Reviews detected but none parsed on the page.")
    prod = shopify_product(page_url)
    return reviews, prod.get("title", ""), "Shopify Product Reviews"


# --------------------------------------------------------------------------
# Universal fallback: JSON-LD Review objects (partial — many stores omit these)
# --------------------------------------------------------------------------
def jsonld_scrape(html, host, progress, page_url):
    reviews = []
    for blk in re.findall(r'<script[^>]+application/ld\+json[^>]*>(.*?)</script>', html, re.S):
        try:
            data = json.loads(blk.strip())
        except Exception:
            continue
        stack = data if isinstance(data, list) else [data]
        for node in stack:
            if not isinstance(node, dict):
                continue
            revs = node.get("review")
            if isinstance(revs, dict):
                revs = [revs]
            if not isinstance(revs, list):
                continue
            for rv in revs:
                if not isinstance(rv, dict):
                    continue
                rr = rv.get("reviewRating", {}) or {}
                auth = rv.get("author", {})
                nm = auth.get("name") if isinstance(auth, dict) else auth
                reviews.append(make_review(
                    rr.get("ratingValue"), rv.get("name"),
                    rv.get("reviewBody") or rv.get("description"),
                    nm, rv.get("datePublished"), None, None,
                ))
    progress(len(reviews))
    if not reviews:
        raise RuntimeError("No review app with a public feed was detected, and the page "
                           "carries no embedded review data. Send me this URL and I'll add "
                           "a handler for its review app.")
    prod = shopify_product(page_url)
    return reviews, prod.get("title", ""), "JSON-LD (partial)"


# --------------------------------------------------------------------------
# Dispatcher — priority-ordered registry
# --------------------------------------------------------------------------
# Apps detected but not yet wired to a data source: clear message instead of a guess.
UNWIRED = [
    ("junip", "Junip"), ("juniphq", "Junip"),
    ("fera.ai", "Fera"), ("data-fera", "Fera"),
    ("growave", "Growave"), ("reviewsjs.growave", "Growave"),
    ("opinew", "Opinew"),
    ("rivyo", "Rivyo"), ("thimatic-app", "Rivyo"),
    ("ryviu", "Ryviu"),
    ("ali-reviews", "Ali Reviews"), ("alireviews", "Ali Reviews"),
    ("trustpilot", "Trustpilot"),
    ("kudobuzz", "Kudobuzz"),
]


# Distinctive markers per provider (avoid bare brand names that appear in blogs/themes).
# Each handler is normalised to signature (html, host, progress, page_url).
PROVIDERS = [
    ("Okendo", ["oke-reviews-settings", "api.okendo.io", "okereviews", "data-oke-",
                "okendoreviews", "okendo.io", "media-dynamic.okendo"],
     lambda h, ho, pr, u: okendo_scrape(h, ho, pr, u)),
    ("Yotpo", ["staticw2.yotpo.com", "cdn-widgetsrepository.yotpo.com",
               "yotpo-reviews-main-widget", "yotpo-main-widget", "yotpowidgetsinstance"],
     lambda h, ho, pr, u: yotpo_scrape(h, ho, pr)),
    ("Judge.me", ["jdgm-widget", "jdgm-rev", "data-jdgm", "judge.me/reviews",
                  "judgeme_product_reviews", "jdgm-form"],
     lambda h, ho, pr, u: judgeme_scrape(h, ho, pr, u)),
    ("Loox", ["loox.io/widget", "looxreviews", "loox-reviews", "loox.io/js", "id=\"looxreviews"],
     loox_scrape),
    ("Stamped.io", ["stamped.io", "stamped-main-widget", "data-widget-type=\"stamped"],
     stamped_scrape),
    ("REVIEWS.io", ["widget.reviews.io", "reviews.io/product", "ruk_rating", "reviewsio-"],
     reviewsio_scrape),
    ("Shopify Product Reviews", ["spr-review", "shopify-product-reviews", "spr-badge"],
     spr_scrape),
]


def detect_and_scrape(page_url, progress):
    parsed = urlparse(page_url)
    host = parsed.netloc
    hl = host.lower()

    # Marketplaces are not Shopify stores and don't expose review widgets we can read.
    if "amazon." in hl:
        raise RuntimeError(
            "Amazon isn't supported. It blocks automated scraping (returns a CAPTCHA / robot "
            "page) and has no public review API. Also, Amazon reviews can't be syndicated to "
            "Walmart — only reviews you collected on your own website qualify. This tool reads "
            "Shopify review widgets: Okendo, Yotpo, Judge.me, Loox, Stamped, REVIEWS.io.")
    if any(m in hl for m in ("walmart.com", "ebay.", "etsy.", "aliexpress.")):
        raise RuntimeError(
            "That's a marketplace listing, not a Shopify store. This tool reads a store's own "
            "Shopify review widget — paste your product page URL from your website instead.")

    # Fetch the page; recover from 404s (short/alias product handles) and never let a
    # raw network error escape — always surface a clean, actionable message.
    try:
        html, final = http_get(page_url)
    except HTTPError as e:
        if e.code == 404 and "/products/" in parsed.path.lower():
            real = shopify_resolve_handle(host, product_handle(parsed.path))
            if real:
                page_url = f"https://{host}/products/{real}"
                parsed = urlparse(page_url)
                try:
                    html, final = http_get(page_url)
                except Exception:
                    raise RuntimeError("That product URL returned 404 and the closest match "
                                       "couldn't be loaded. Open the product on your store and "
                                       "paste its exact URL.")
            else:
                raise RuntimeError("That product URL returned 404 (not found) and no matching "
                                   "product could be located. Check the URL is correct and the "
                                   "product is published.")
        else:
            raise RuntimeError(f"Couldn't load that URL (HTTP {e.code}). Check the link is "
                               f"correct and publicly accessible.")
    except URLError as e:
        raise RuntimeError(f"Couldn't reach that URL ({getattr(e, 'reason', e)}). "
                           f"Check the link is correct.")
    low = html.lower()

    # Score each provider by how many distinctive markers appear; strongest wins.
    scored = sorted(
        ((sum(low.count(m) for m in markers), i, name, fn)
         for i, (name, markers, fn) in enumerate(PROVIDERS)),
        key=lambda x: (-x[0], x[1]),
    )
    reviews = name = prov = None
    candidates = [(nm, fn) for score, _, nm, fn in scored if score > 0]
    detected = [nm for nm, _ in candidates]

    # Try detected providers strongest-first; a store can host more than one widget,
    # so fall through to the next if one has no data / errors.
    for nm, fn in candidates:
        try:
            reviews, name, prov = fn(html, host, progress, page_url)
            if reviews:
                break
        except Exception:
            continue

    if not reviews:
        # Non-product URL (landing / funnel / collection / home): reviews are tied to a
        # specific product, so point the user at the product URLs found on the page.
        if "/products/" not in parsed.path.lower() and "/products/" not in urlparse(final).path.lower():
            handles = [h for h in _uniq(re.findall(r'/products/([a-z0-9][a-z0-9-]{2,})', low))
                       if not h.endswith(".js")][:6]
            if handles:
                links = "\n".join(f"  • https://{host}/products/{h}" for h in handles)
                raise RuntimeError(
                    "This looks like a landing/collection page, not a product page. Reviews are "
                    "tied to a specific product — open the product you want and paste its URL "
                    "(it has /products/… in it). Products found on this page:\n" + links)

        hit = next((label for token, label in UNWIRED if token in low), None)
        try:
            reviews, name, prov = jsonld_scrape(html, host, progress, page_url)
            if hit and hit not in detected:
                prov = f"{hit} via JSON-LD (partial)"
        except Exception:
            names = ", ".join(detected + ([hit] if hit and hit not in detected else [])) or None
            if names:
                raise RuntimeError(
                    f"Detected {names}, but couldn't pull reviews automatically here "
                    f"(the widget may be JS-rendered or misconfigured on this store). "
                    f"Send me this exact URL and I'll wire/tune the handler.")
            raise RuntimeError(
                "No review app with a public feed was detected on this page. "
                "Send me the URL and I'll add a handler for its review app.")

    # product name fallback from <title>
    if not name:
        t = re.search(r"<title[^>]*>(.*?)</title>", html, re.S | re.I)
        if t:
            name = unescape(t.group(1).strip())

    return {
        "provider": prov,
        "product_name": name,
        "product_url": final,
        "source": host,
        "reviews": reviews,
        "count": len(reviews),
    }


# --------------------------------------------------------------------------
# Exports
# --------------------------------------------------------------------------
IMPORT_COLS = [
    ("product_name", "Product Name"),
    ("sku", "SKU"),
    ("walmart_item_id", "Walmart Item ID"),
    ("gtin_upc", "GTIN / UPC"),
    ("rating", "Rating (1-5)"),
    ("review_title", "Review Title"),
    ("review_text", "Review Text"),
    ("reviewer_name", "Reviewer Name"),
    ("review_date", "Review Date (YYYY-MM-DD)"),
    ("verified_purchaser", "Verified Purchaser"),
    ("recommended", "Recommended"),
    ("product_url", "Source Product URL"),
    ("source", "Source Site"),
]


def row_for(data, rv):
    return {
        "product_name": data["product_name"],
        "sku": "",
        "walmart_item_id": "",
        "gtin_upc": "",
        "rating": rv["rating"],
        "review_title": rv["review_title"],
        "review_text": rv["review_text"],
        "reviewer_name": rv["reviewer_name"],
        "review_date": rv["review_date"],
        "verified_purchaser": rv["verified_purchaser"],
        "recommended": rv["recommended"],
        "product_url": data["product_url"],
        "source": data["source"],
    }


def export_csv_bytes(data):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow([h for _, h in IMPORT_COLS])
    for rv in data["reviews"]:
        r = row_for(data, rv)
        w.writerow([r[k] for k, _ in IMPORT_COLS])
    return buf.getvalue().encode("utf-8-sig")


def export_xlsx_bytes(data):
    if not HAVE_XLSX:
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = "Walmart Import"
    head_fill = PatternFill("solid", fgColor="2C3E50")
    head_font = Font(bold=True, color="FFFFFF")
    for c, (_, label) in enumerate(IMPORT_COLS, 1):
        cell = ws.cell(1, c, label)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for rv in data["reviews"]:
        r = row_for(data, rv)
        ws.append([r[k] for k, _ in IMPORT_COLS])
    widths = [34, 14, 16, 16, 11, 30, 60, 20, 20, 16, 13, 40, 22]
    for i, wdt in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i) if i <= 26 else "A" + chr(64 + i - 26)].width = wdt
    ws.freeze_panes = "A2"

    # raw sheet
    ws2 = wb.create_sheet("All Fields")
    extra_keys = sorted({k for rv in data["reviews"] for k in rv["extra"]})
    headers = ["rating", "review_title", "review_text", "reviewer_name",
               "review_date", "verified_purchaser", "recommended"] + extra_keys
    for c, h in enumerate(headers, 1):
        cell = ws2.cell(1, c, h)
        cell.font = Font(bold=True)
    for rv in data["reviews"]:
        base = [rv["rating"], rv["review_title"], rv["review_text"], rv["reviewer_name"],
                rv["review_date"], rv["verified_purchaser"], rv["recommended"]]
        base += [rv["extra"].get(k, "") for k in extra_keys]
        ws2.append(base)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def safe_name(data):
    base = re.sub(r"[^\w\- ]+", "", (data["product_name"] or "reviews"))[:60].strip()
    return (base or "reviews").replace(" ", "_")


# --------------------------------------------------------------------------
# Web UI + routes
# --------------------------------------------------------------------------
_HERE = os.path.dirname(os.path.abspath(__file__))


def _load_page():
    """Serve the branded tool page (tool_page.html) with fonts embedded."""
    try:
        with open(os.path.join(_HERE, "cabinet_fonts.css"), encoding="utf-8") as f:
            fonts = f.read()
    except Exception:
        fonts = ("@import url('https://api.fontshare.com/v2/css?"
                 "f[]=cabinet-grotesk@400,500,700,800&display=swap');")
    try:
        with open(os.path.join(_HERE, "tool_page.html"), encoding="utf-8") as f:
            page = f.read()
    except Exception as e:
        return ("<!doctype html><meta charset=utf-8><h1>tool_page.html missing</h1>"
                "<pre>%s</pre>" % e)
    return page.replace("/*__FONTS__*/", fonts)


INDEX_HTML = _load_page()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def client_ip(self):
        # Behind Dokploy/Traefik the real client is in X-Forwarded-For.
        xff = self.headers.get("X-Forwarded-For")
        if xff:
            return xff.split(",")[0].strip()
        return self.client_address[0]

    def _send(self, code, body, ctype="application/json", headers=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Access-Control-Allow-Origin", ALLOWED_ORIGIN)
        self.send_header("Vary", "Origin")
        if headers:
            for k, v in headers.items():
                self.send_header(k, v)
        self.end_headers()
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", ALLOWED_ORIGIN)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Max-Age", "86400")
        self.send_header("Vary", "Origin")
        self.end_headers()

    def _serve_video(self, path):
        try:
            size = os.path.getsize(path)
        except OSError:
            return self._send(404, "video not found", "text/plain")
        rng = self.headers.get("Range")
        try:
            if rng:
                m = re.match(r"bytes=(\d+)-(\d*)", rng)
                start = int(m.group(1))
                end = int(m.group(2)) if m and m.group(2) else size - 1
                end = min(end, size - 1)
                length = end - start + 1
                self.send_response(206)
                self.send_header("Content-Type", "video/mp4")
                self.send_header("Accept-Ranges", "bytes")
                self.send_header("Content-Range", f"bytes {start}-{end}/{size}")
                self.send_header("Content-Length", str(length))
                self.end_headers()
                with open(path, "rb") as f:
                    f.seek(start)
                    remaining = length
                    while remaining > 0:
                        chunk = f.read(min(65536, remaining))
                        if not chunk:
                            break
                        self.wfile.write(chunk)
                        remaining -= len(chunk)
            else:
                self.send_response(200)
                self.send_header("Content-Type", "video/mp4")
                self.send_header("Accept-Ranges", "bytes")
                self.send_header("Content-Length", str(size))
                self.end_headers()
                with open(path, "rb") as f:
                    while True:
                        chunk = f.read(65536)
                        if not chunk:
                            break
                        self.wfile.write(chunk)
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
            pass  # client seeked / closed the video — normal

    def do_GET(self):
        p = urlparse(self.path)
        if p.path == "/":
            return self._send(200, INDEX_HTML, "text/html; charset=utf-8")

        if p.path == "/about.mp4":
            return self._serve_video(os.path.join(_HERE, "about.mp4"))

        if p.path == "/scrape":
            if not rate_ok(self.client_ip()):
                return self._send(429, json.dumps({"ok": False,
                    "error": f"Too many requests — please wait a moment and try again "
                             f"(limit {RATE_MAX} per {RATE_WINDOW}s)."}),
                    headers={"Retry-After": str(RATE_WINDOW)})
            url = parse_qs(p.query).get("url", [""])[0]
            if not url:
                return self._send(200, json.dumps({"ok": False, "error": "No URL given."}))
            try:
                data = detect_and_scrape(url, lambda n: None)
            except (HTTPError, URLError) as e:
                return self._send(200, json.dumps({"ok": False, "error": f"Network error: {e}"}))
            except Exception as e:
                return self._send(200, json.dumps({"ok": False, "error": str(e)}))
            token = uuid.uuid4().hex
            RESULTS[token] = data
            rated = [r for r in data["reviews"] if isinstance(r["rating"], (int, float))]
            avg = round(sum(r["rating"] for r in rated) / len(rated), 2) if rated else 0
            ver = sum(1 for r in data["reviews"] if r["verified_purchaser"] == "Yes")
            return self._send(200, json.dumps({
                "ok": True, "token": token, "provider": data["provider"],
                "count": data["count"], "product_name": data["product_name"],
                "avg": avg, "verified": ver,
                "preview": data["reviews"][:50],
            }))

        if p.path == "/download":
            q = parse_qs(p.query)
            token = q.get("token", [""])[0]
            fmt = q.get("fmt", ["csv"])[0]
            data = RESULTS.get(token)
            if not data:
                return self._send(404, "expired", "text/plain")
            fn = safe_name(data)
            if fmt == "xlsx":
                body = export_xlsx_bytes(data)
                if body is None:
                    return self._send(200, json.dumps({"ok": False, "error": "openpyxl not installed"}))
                return self._send(200, body,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                  {"Content-Disposition": f'attachment; filename="{fn}.xlsx"'})
            body = export_csv_bytes(data)
            return self._send(200, body, "text/csv",
                              {"Content-Disposition": f'attachment; filename="{fn}_walmart.csv"'})

        return self._send(404, "not found", "text/plain")

    def do_POST(self):
        p = urlparse(self.path)
        if p.path == "/support":
            if not rate_ok(self.client_ip()):
                return self._send(429, json.dumps({"ok": False,
                    "error": "Too many requests — please wait a moment and try again."}),
                    headers={"Retry-After": str(RATE_WINDOW)})
            try:
                length = int(self.headers.get("Content-Length", 0) or 0)
                raw = self.rfile.read(length) if length else b"{}"
                data = json.loads(raw.decode("utf-8"))
            except Exception as e:
                return self._send(200, json.dumps({"ok": False, "error": f"Bad request: {e}"}))
            name = str(data.get("name", "")).strip()
            url = str(data.get("url", "")).strip()
            desc = str(data.get("description", "")).strip()
            if not (name and url and desc):
                return self._send(200, json.dumps({"ok": False, "error": "All fields are required."}))
            rec = {"time": time.strftime("%Y-%m-%d %H:%M:%S"),
                   "name": name[:200], "url": url[:500], "description": desc[:4000]}
            # Always keep a local copy as backup...
            try:
                with open(os.path.join(_HERE, "support_requests.jsonl"), "a", encoding="utf-8") as f:
                    f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            except Exception as e:
                return self._send(200, json.dumps({"ok": False, "error": f"Could not save: {e}"}))
            # ...and email it if SMTP is configured (env vars).
            emailed, err = send_support_email(rec)
            print(f"[support] {rec['time']} · {name} · {url} · emailed={emailed}"
                  + (f" ({err})" if err else ""))
            return self._send(200, json.dumps({"ok": True}))
        return self._send(404, "not found", "text/plain")


def main():
    srv = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}/"
    print("=" * 58)
    print("  Review Scraper  ->  Walmart Export")
    print(f"  Listening on {url}")
    if not HAVE_XLSX:
        print("  NOTE: openpyxl missing -> Excel disabled, CSV still works.")
        print("        install with:  pip install openpyxl")
    print("  Press Ctrl+C to stop.")
    print("=" * 58)
    # Only pop a browser when running locally on a real desktop — never on a server.
    if HOST in ("127.0.0.1", "localhost") and not os.environ.get("NO_BROWSER"):
        def _open():
            try:
                webbrowser.open(f"http://127.0.0.1:{PORT}/")
            except Exception:
                pass
        threading.Timer(0.8, _open).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
