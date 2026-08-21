#!/usr/bin/env python3
"""
Mārtaṇḍa Jyotiṣī — one-command article publisher.

Usage:
    python scripts/publish.py "My Article Title" path/to/article.docx --tags Navamsa,Vivaha
    python scripts/publish.py "My Article Title" path/to/article.md --tags Navamsa
    python scripts/publish.py "Draft Title" article.docx --draft

What it does:
  1. Reads your .docx (or .md) article
  2. Converts it to Markdown with proper Hugo frontmatter
  3. Saves it into content/posts/
  4. Commits and pushes to GitHub
  5. Netlify/GitHub Pages picks up the push and rebuilds the live site automatically
"""

import argparse
import datetime
import re
import subprocess
import sys
from pathlib import Path

SITE_ROOT = Path(__file__).resolve().parent.parent
POSTS_DIR = SITE_ROOT / "content" / "posts"


def slugify(title: str) -> str:
    s = title.lower()
    s = re.sub(r"[^a-z0-9\s-]", "", s)
    s = re.sub(r"\s+", "-", s.strip())
    return s


def docx_to_markdown(path: Path) -> str:
    from docx import Document
    doc = Document(str(path))
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = (para.style.name or "").lower()
        if "heading 1" in style or "title" in style:
            lines.append(f"# {text}")
        elif "heading 2" in style:
            lines.append(f"## {text}")
        elif "heading 3" in style:
            lines.append(f"### {text}")
        elif "quote" in style:
            lines.append(f"> {text}")
        else:
            lines.append(text)
        lines.append("")
    return "\n".join(lines)


def build_frontmatter(title: str, tags: list[str], summary: str, draft: bool) -> str:
    date_str = datetime.date.today().isoformat()
    tags_str = ", ".join(f'"{t.strip()}"' for t in tags if t.strip())
    fm = [
        "---",
        f'title: "{title}"',
        f"date: {date_str}",
        f"tags: [{tags_str}]",
        f'summary: "{summary}"',
        f"draft: {str(draft).lower()}",
        "---",
        "",
    ]
    return "\n".join(fm)


def run(cmd, cwd):
    print(f"$ {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if result.stdout.strip():
        print(result.stdout.strip())
    if result.returncode != 0:
        print(result.stderr.strip(), file=sys.stderr)
    return result.returncode == 0


def main():
    p = argparse.ArgumentParser(description="Publish a Jyotiṣa article to the Mārtaṇḍa site.")
    p.add_argument("title", help="Article title, in quotes")
    p.add_argument("source", help="Path to the .docx or .md source file")
    p.add_argument("--tags", default="", help="Comma-separated tags, e.g. Navamsa,Vivaha")
    p.add_argument("--summary", default="", help="One-line summary (optional; auto-generated if omitted)")
    p.add_argument("--draft", action="store_true", help="Save as a draft (not published live)")
    p.add_argument("--no-push", action="store_true", help="Skip git commit/push, just generate the file locally")
    args = p.parse_args()

    source = Path(args.source)
    if not source.exists():
        print(f"File not found: {source}")
        sys.exit(1)

    if source.suffix.lower() == ".docx":
        body = docx_to_markdown(source)
    else:
        body = source.read_text(encoding="utf-8")

    summary = args.summary or " ".join(body.split()[:30]) + "..."
    tags = args.tags.split(",") if args.tags else []

    slug = slugify(args.title)
    out_path = POSTS_DIR / f"{slug}.md"
    frontmatter = build_frontmatter(args.title, tags, summary, args.draft)
    out_path.write_text(frontmatter + body, encoding="utf-8")
    print(f"Written: {out_path}")

    if args.no_push:
        print("Skipped git push (--no-push set). Run `hugo server` to preview locally.")
        return

    run(["git", "add", "-A"], cwd=SITE_ROOT)
    status = "draft" if args.draft else "published"
    run(["git", "commit", "-m", f"Add article: {args.title} ({status})"], cwd=SITE_ROOT)
    ok = run(["git", "push"], cwd=SITE_ROOT)
    if ok:
        print(f"\nPushed. Your site will rebuild automatically — live in 1-2 minutes.")
    else:
        print("\nPush failed — check that 'git remote' is set to your GitHub repo (see README).")


if __name__ == "__main__":
    main()
