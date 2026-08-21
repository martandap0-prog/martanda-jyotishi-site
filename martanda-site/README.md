# Mārtaṇḍa Jyotiṣī — Website + Auto-Publishing

A Hugo static site with a Python script that turns a finished article
(.docx or .md) into a live blog post with one command.

## What's in this folder
- `hugo.toml` — site config
- `layouts/` — page templates (home, article, list) in your navy/gold/maroon Georgia brand
- `static/css/style.css` — the styling
- `content/posts/` — your articles live here as Markdown (one sample article included)
- `scripts/publish.py` — the automation script

---

## PART 1 — One-time setup (do this once, start to finish)

### Step 1: Install Git on Windows
- Download from https://git-scm.com/download/win and run the installer (defaults are fine).
- Open VS Code, open a new Terminal (Terminal menu → New Terminal).
- Set your identity (only needed once, ever):
```
git config --global user.name "Monarch"
git config --global user.email "your-email@example.com"
```

### Step 2: Install Hugo on Windows
```
winget install Hugo.Hugo.Extended
```
(Close and reopen your terminal after this so the `hugo` command is recognized.)

### Step 3: Create a GitHub account
- Go to github.com → Sign up (free).

### Step 4: Create a new (empty) repository on GitHub
- github.com → click the "+" top right → "New repository"
- Name it: `martanda-jyotishi-site`
- Leave it **Public**, do **NOT** check "Add a README" (you already have one)
- Click "Create repository" — GitHub will show you a page with commands; ignore those, use the ones below instead.

### Step 5: Push this folder to your new GitHub repo
Unzip this folder, open it in VS Code, open the Terminal, and run these five commands
one at a time (replace YOUR_USERNAME with your actual GitHub username):
```
git init
git remote add origin https://github.com/YOUR_USERNAME/martanda-jyotishi-site.git
git add -A
git commit -m "Initial site"
git branch -M main
git push -u origin main
```
The first time you push, a browser window will pop up asking you to log into GitHub —
log in there, and the push will complete.

### Step 6: Connect free hosting via Netlify
1. Go to netlify.com → sign up (use "Sign up with GitHub" — simplest option)
2. "Add new site" → "Import an existing project" → "Deploy with GitHub"
3. Pick your `martanda-jyotishi-site` repo
4. Build settings:
   - Build command: `hugo --minify`
   - Publish directory: `public`
5. Click Deploy. Wait ~1 minute. Netlify gives you a live URL like
   `random-name-123.netlify.app` — that's your site, live, right now.

### Step 7 (optional, do later): Custom domain
- Buy `martandajyotishi.com` (or similar) from Namecheap/GoDaddy (~$10-15/year)
- In Netlify: Site settings → Domain management → Add custom domain
- Netlify shows you exactly which DNS records to add at your domain registrar
  (usually just one or two lines) — follow their on-screen instructions.

**At this point your site is live on the internet. Everything below is for publishing articles.**

---

## PART 2 — Publishing a new article (the automated part)

Once Part 1 is done, publishing any future article is one command:

```
python scripts/publish.py "Understanding the Navāṃśa Chart" my_article.docx --tags Navamsa,Vivaha
```

This will:
1. Convert your `.docx` to Markdown
2. Add proper frontmatter (title, date, tags, summary)
3. Save it into `content/posts/`
4. Commit and push to GitHub
5. Netlify rebuilds automatically — live in 1-2 minutes

To save as a draft (not published yet):
```
python scripts/publish.py "Draft Article" my_article.docx --draft
```

To just generate the file locally without pushing (e.g. to preview first):
```
python scripts/publish.py "Draft Article" my_article.docx --no-push
```

You'll need `python-docx` installed once:
```
pip install python-docx
```

## Local preview before publishing
```
hugo server
```
Then open http://localhost:1313 to see exactly how it'll look, live, before you push.
