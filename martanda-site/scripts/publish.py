#!/usr/bin/env python3
"""
Mārtaṇḍa Jyotiṣī — one-command article publisher.

Usage:
    python scripts/publish.py "My Article Title" path/to/article.docx --tags Navamsa,Vivaha --level beginner
    python scripts/publish.py "My Article Title" path/to/article.md --tags Navamsa
    python scripts/publish.py "Draft Title" article.docx --draft

What it does:
  1. Reads your .docx (or .md) article
  2. Converts it to Markdown with proper Hugo frontmatter
     - Preserves paragraph structure, headings, quotes
     - Extracts embedded images and saves them to static/images/posts/<slug>/
     - Converts Word tables (including tables containing images, like chart
       comparison rows) into HTML tables so they render correctly on the site
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
IMAGES_ROOT = SITE_ROOT / "static" / "images" / "posts"


def slugify(title: str) -> str:
    s = title.lower()
    s = re.sub(r"[^a-z0-9\s-]", "", s)
    s = re.sub(r"\s+", "-", s.strip())
    return s


def docx_to_markdown(path: Path, slug: str) -> str:
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(str(path))
    img_dir = IMAGES_ROOT / slug
    img_dir.mkdir(parents=True, exist_ok=True)
    img_counter = [0]

    def extract_images(xml_element) -> list[str]:
        """Find every image inside this XML element, save it, return web paths."""
        urls = []
        for blip in xml_element.findall(".//" + qn("a:blip")):
            rId = blip.get(qn("r:embed"))
            if not rId:
                continue
            try:
                part = doc.part.related_parts[rId]
            except KeyError:
                continue
            img_counter[0] += 1
            ext = part.content_type.split("/")[-1]
            ext = "jpg" if ext == "jpeg" else ext
            fname = f"img{img_counter[0]}.{ext}"
            (img_dir / fname).write_bytes(part.blob)
            urls.append(f"/images/posts/{slug}/{fname}")
        return urls

    def para_text(p_element) -> str:
        return "".join(node.text or "" for node in p_element.findall(".//" + qn("w:t"))).strip()

    def para_style_name(p_element) -> str:
        pPr = p_element.find(qn("w:pPr"))
        if pPr is None:
            return ""
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None:
            return ""
        val = pStyle.get(qn("w:val")) or ""
        style_map = {
            "Title": "title", "Heading1": "heading 1", "Heading2": "heading 2",
            "Heading3": "heading 3", "Quote": "quote", "IntenseQuote": "quote",
        }
        return style_map.get(val, val.lower())

    def heading_level_from_formatting(p_element) -> int:
        """Fallback heuristic: some docs use direct run formatting (size/bold/italic)
        instead of Word's built-in heading styles. Detect common size/weight patterns."""
        runs = p_element.findall(qn("w:r"))
        if not runs:
            return 0
        rPr = runs[0].find(qn("w:rPr"))
        if rPr is None:
            return 0
        sz_el = rPr.find(qn("w:sz"))
        b_el = rPr.find(qn("w:b"))
        i_el = rPr.find(qn("w:i"))
        if sz_el is None or b_el is None:
            return 0
        try:
            sz = int(sz_el.get(qn("w:val")))
        except (TypeError, ValueError):
            return 0
        if sz >= 32:
            return 2
        if sz >= 24 and i_el is not None:
            return 3
        return 0

    def paragraph_to_md(p_element, text=None, images=None) -> str:
        if text is None:
            text = para_text(p_element)
        if images is None:
            images = extract_images(p_element)
        lines = []
        for url in images:
            lines.append(f"![]({url})")
        if text:
            style = para_style_name(p_element)
            if "heading 1" in style or style == "title":
                lines.append(f"## {text}")
            elif "heading 2" in style:
                lines.append(f"### {text}")
            elif "heading 3" in style:
                lines.append(f"#### {text}")
            elif "quote" in style:
                lines.append(f"> {text}")
            else:
                level = heading_level_from_formatting(p_element)
                if level == 2:
                    lines.append(f"## {text}")
                elif level == 3:
                    lines.append(f"### {text}")
                else:
                    lines.append(text)
        return "\n\n".join(lines)

    def table_to_html(tbl_element) -> str:
        rows_html = []
        for tr in tbl_element.findall(qn("w:tr")):
            cells_html = []
            for tc in tr.findall(qn("w:tc")):
                cell_parts = []
                for p in tc.findall(qn("w:p")):
                    text = para_text(p)
                    images = extract_images(p)
                    for url in images:
                        cell_parts.append(f'<img src="{url}" style="max-width:100%;">')
                    if text:
                        cell_parts.append(text)
                cell_content = "<br>".join(cell_parts) if cell_parts else "&nbsp;"
                cells_html.append(f"<td style=\"padding:8px; border:1px solid #ccc; text-align:center; vertical-align:middle;\">{cell_content}</td>")
            rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
        return '<table style="border-collapse:collapse; width:100%; margin: 1.5rem 0;">\n' + "\n".join(rows_html) + "\n</table>"

    body = doc.element.body
    blocks = []
    first_heading_skipped = False
    is_first_text_para = True
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            text = para_text(child)
            images = extract_images(child)
            if not first_heading_skipped and text and not images:
                level = heading_level_from_formatting(child)
                style = para_style_name(child)
                if level == 2 or "heading 1" in style or style == "title":
                    first_heading_skipped = True
                    continue
            if is_first_text_para and text and not images and first_heading_skipped:
                is_first_text_para = False
                blocks.append(f"*{text}*")
                continue
            if text or images:
                is_first_text_para = False
            md = paragraph_to_md(child, text=text, images=images)
            if md:
                blocks.append(md)
        elif child.tag == qn("w:tbl"):
            is_first_text_para = False
            blocks.append(table_to_html(child))

    return "\n\n".join(blocks) + "\n"


def build_frontmatter(title: str, tags: list[str], summary: str, draft: bool, level: str) -> str:
    date_str = datetime.date.today().isoformat()
    tags_str = ", ".join(f'"{t.strip()}"' for t in tags if t.strip())
    fm = [
        "---",
        f'title: "{title}"',
        f"date: {date_str}",
        f"tags: [{tags_str}]",
        f'levels: ["{level}"]',
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
    p.add_argument("--level", default="beginner", choices=["beginner", "intermediate", "research"],
                   help="Difficulty tier: beginner, intermediate, or research (default: beginner)")
    p.add_argument("--draft", action="store_true", help="Save as a draft (not published live)")
    p.add_argument("--no-push", action="store_true", help="Skip git commit/push, just generate the file locally")
    args = p.parse_args()

    source = Path(args.source)
    if not source.exists():
        print(f"File not found: {source}")
        sys.exit(1)

    slug = slugify(args.title)

    if source.suffix.lower() == ".docx":
        body = docx_to_markdown(source, slug)
    else:
        body = source.read_text(encoding="utf-8")

    plain_for_summary = re.sub(r"[#>*!\[\]()<>]", "", body)
    plain_for_summary = re.sub(r"http\S+", "", plain_for_summary)
    summary = args.summary or " ".join(plain_for_summary.split()[:30]) + "..."
    summary = summary.replace('"', "'")
    tags = args.tags.split(",") if args.tags else []

    out_path = POSTS_DIR / f"{slug}.md"
    frontmatter = build_frontmatter(args.title, tags, summary, args.draft, args.level)
    out_path.write_text(frontmatter + body, encoding="utf-8")
    print(f"Written: {out_path}")
    if (IMAGES_ROOT / slug).exists():
        n_images = len(list((IMAGES_ROOT / slug).iterdir()))
        if n_images:
            print(f"Saved {n_images} image(s) to: {IMAGES_ROOT / slug}")

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