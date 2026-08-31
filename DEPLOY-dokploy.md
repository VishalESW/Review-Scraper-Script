# Deploy to Dokploy

Dokploy runs your app as a Docker container on your own VPS and gives it a domain
+ automatic HTTPS (via Traefik). This project already ships a `Dockerfile`, so
Dokploy builds and runs it directly.

## What's in the box
- `app.py` — the server (scraper + tool page + `/about.mp4` + `/support`)
- `tool_page.html`, `cabinet_fonts.css`, `about.mp4` — the UI + assets it serves
- `Dockerfile`, `requirements.txt`, `.dockerignore` — the deploy setup

The container binds to `0.0.0.0:8765` (set by the Dockerfile). Dokploy/Traefik
proxies your domain to that port.

---

## Step 1 — Put the code in a Git repo
Dokploy deploys from Git. Push this folder to a GitHub/GitLab repo (private is fine):

```bash
git init
git add .
git commit -m "Review scraper tool"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```

> `about.mp4` (~19 MB) is committed too, since the app serves it. If your Git host
> rejects it, use Git LFS, or host the video elsewhere and change the `<video src>`
> in `tool_page.html` to that URL.

## Step 2 — Create the app in Dokploy
1. Dokploy dashboard → **Create Project** → inside it **Create Service → Application**.
2. **Provider**: connect your GitHub/GitLab account (or use a Git URL + deploy key),
   then select the repo and the `main` branch.
3. **Build Type**: **Dockerfile** (Dokploy auto-detects the `Dockerfile` in the repo
   root — leave the path as `./Dockerfile`).
4. Click **Deploy**. Watch the **Logs** tab until you see:
   `Listening on http://0.0.0.0:8765/`

## Step 3 — Domain + HTTPS
1. In the app's **Domains** tab → **Add Domain**.
2. Host: e.g. `tool.esellerworld.com`
3. **Container Port**: `8765`
4. **HTTPS**: on (Let's Encrypt). **Path**: `/`
5. In your DNS, add an **A record**: `tool.esellerworld.com → <your VPS public IP>`.
   Once DNS resolves, Dokploy issues the certificate automatically.

Visit `https://tool.esellerworld.com` — the full tool loads.

## Step 4 — Put it on your website
Two options:
- **Subdomain (simplest):** just link to `https://tool.esellerworld.com`, or embed it
  in a page with an iframe:
  `<iframe src="https://tool.esellerworld.com" style="width:100%;height:100vh;border:0"></iframe>`
- **Inline on an existing page (different origin):** paste `tool_page.html`'s markup
  into your page and point the `fetch()` calls at `https://tool.esellerworld.com`.
  This needs CORS added to `app.py` — ask and I'll wire it.

---

## Environment variables
Set these in the app's **Environment** tab in Dokploy. Only the SMTP ones need
your input; the rest have sensible defaults.

| Var | Default | Purpose |
|-----|---------|---------|
| `HOST` | `0.0.0.0` | Bind address (set by Dockerfile) |
| `PORT` | `8765` | Container port (match it in the Domains tab) |
| `NO_BROWSER` | `1` | Never open a desktop browser on the server |
| `ALLOWED_ORIGIN` | `*` | CORS. Set to `https://esellerworld.com` if you embed the page on your main domain and call this backend cross-origin. Leave `*` if serving on a subdomain. |
| `RATE_LIMIT_MAX` | `20` | Max requests per IP per window (for `/scrape` + `/support`) |
| `RATE_LIMIT_WINDOW` | `60` | Window length in seconds |

### Support-form email (optional but recommended)
`/support` always writes `support_requests.jsonl` as a backup **and** emails you if
these are set:

| Var | Example | Purpose |
|-----|---------|---------|
| `SMTP_HOST` | `smtp.gmail.com` | Your SMTP server |
| `SMTP_PORT` | `587` | `587` (STARTTLS) or `465` (SSL) |
| `SMTP_USER` | `you@esellerworld.com` | SMTP login |
| `SMTP_PASS` | *(app password)* | SMTP password / app-password |
| `SUPPORT_TO` | `contact@esellerworld.com` | Where issues are sent |
| `SUPPORT_FROM` | *(defaults to SMTP_USER)* | From address |

If `SMTP_HOST` is unset, it just logs to the file (no email). Server logs show
`emailed=True/False` per submission.

> The file backup still resets on redeploy. With email configured you don't need to
> persist it, but if you want to, mount a Dokploy **Volume** at `/app/support_requests.jsonl`.

## Hardening — already included
- **Rate limiting** per client IP on `/scrape` and `/support` (returns `429` when
  exceeded). Tune with `RATE_LIMIT_MAX` / `RATE_LIMIT_WINDOW`.
- **CORS** headers + `OPTIONS` preflight, so the page can call the backend from your
  domain. Lock it down by setting `ALLOWED_ORIGIN`.
- Real client IP is read from `X-Forwarded-For` (Traefik sets it).

## Redeploying after changes
Push to `main` and hit **Redeploy** in Dokploy (or enable auto-deploy on push).
