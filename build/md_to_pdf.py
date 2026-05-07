#!/usr/bin/env python3
"""Render a markdown file (with mermaid blocks) to PDF using pandoc + mmdc + weasyprint."""
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
BUILD.mkdir(exist_ok=True)

ORIENTATION = {
    "NEB-ALCARRIA": "portrait",
}


def css_for(orientation: str) -> Path:
    path = BUILD / f"style_{orientation}.css"
    path.write_text(
        f"""
@page {{ size: A4 {orientation}; margin: 2cm; }}
body {{ font-family: Helvetica, Arial, sans-serif; font-size: 11pt; line-height: 1.5; color: #222; }}
h1 {{ color: #1f3a93; border-bottom: 2px solid #1f3a93; padding-bottom: 4px; }}
h2 {{ color: #1f3a93; margin-top: 1.4em; }}
h3 {{ color: #2c3e50; }}
code {{ background: #f4f4f4; padding: 1px 4px; border-radius: 3px; font-size: 0.92em; }}
pre {{ background: #f4f4f4; padding: 8px; border-radius: 4px; }}
table {{ border-collapse: collapse; width: 100%; margin: 0.6em 0; font-size: 0.92em; }}
th, td {{ border: 1px solid #ccc; padding: 4px 8px; text-align: left; vertical-align: top; }}
th {{ background: #eef2f7; }}
blockquote {{ border-left: 3px solid #1f3a93; background: #f4f6fa; padding: 6px 12px; margin: 0.6em 0; }}
img {{ max-width: 100%; height: auto; }}
hr {{ border: none; border-top: 1px solid #ddd; margin: 1.2em 0; }}
a {{ color: #1f3a93; text-decoration: none; }}
"""
    )
    return path


def render_mermaid_blocks(md_text: str, stem: str) -> str:
    img_dir = BUILD / f"{stem}_img"
    img_dir.mkdir(exist_ok=True)

    pattern = re.compile(r"```mermaid\n(.*?)\n```", re.DOTALL)

    def replace(match: re.Match) -> str:
        idx = replace.counter
        replace.counter += 1
        mmd_path = img_dir / f"diagram_{idx:02d}.mmd"
        png_path = img_dir / f"diagram_{idx:02d}.png"
        mmd_path.write_text(match.group(1))
        config_path = img_dir / "config.json"
        config_path.write_text('{"theme":"default","themeVariables":{"fontSize":"14px"}}')
        subprocess.run(
            [
                "npx", "-y", "@mermaid-js/mermaid-cli",
                "-i", str(mmd_path),
                "-o", str(png_path),
                "-w", "2000", "-H", "1800",
                "-b", "white",
                "-c", str(config_path),
            ],
            check=True,
        )
        return f"![]({png_path.as_posix()})"

    replace.counter = 1
    return pattern.sub(replace, md_text)


def build_pdf(md_file: Path) -> Path:
    orientation = ORIENTATION.get(md_file.stem, "portrait")
    text = md_file.read_text()
    rendered = render_mermaid_blocks(text, md_file.stem)
    tmp_md = BUILD / f"{md_file.stem}.preprocessed.md"
    tmp_md.write_text(rendered)

    html_path = BUILD / f"{md_file.stem}.html"
    subprocess.run(
        [
            "pandoc", str(tmp_md),
            "-o", str(html_path),
            "--standalone",
            "--metadata", f"title={md_file.stem.replace('_', ' ')}",
            "--css", str(css_for(orientation)),
            "--toc",
            "--toc-depth=2",
        ],
        check=True,
    )

    pdf_path = ROOT / f"{md_file.stem}.pdf"
    subprocess.run(["weasyprint", str(html_path), str(pdf_path)], check=True)
    return pdf_path


if __name__ == "__main__":
    targets = sys.argv[1:] or [
        str(ROOT / "NEB-ALCARRIA.md"),
    ]
    for t in targets:
        print(f"→ Generando PDF de {t}")
        out = build_pdf(Path(t))
        print(f"  ✓ {out} ({ORIENTATION.get(Path(t).stem, 'portrait')})")
