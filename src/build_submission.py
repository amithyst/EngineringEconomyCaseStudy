#!/usr/bin/env python3

from pathlib import Path

import markdown


REPO_ROOT = Path(__file__).resolve().parents[1]
REPORT_MD = REPO_ROOT / "docs" / "analysis" / "final-report.md"
OUTPUT_DIR = REPO_ROOT / "outputs" / "submission"


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>Case Study Report</title>
  <style>
    @page {{
      size: A4;
      margin: 1.5cm;
    }}
    body {{
      font-family: "Noto Serif CJK SC", "Source Han Serif SC", "Songti SC", serif;
      font-size: 10.5pt;
      line-height: 1.45;
      color: #111;
    }}
    h1 {{
      font-size: 18pt;
      margin: 0 0 10pt 0;
    }}
    h2 {{
      font-size: 13pt;
      margin: 12pt 0 6pt 0;
      border-bottom: 1px solid #bbb;
      padding-bottom: 2pt;
    }}
    h3 {{
      font-size: 11pt;
      margin: 8pt 0 4pt 0;
    }}
    p {{
      margin: 4pt 0;
      text-align: justify;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin: 6pt 0 8pt 0;
      font-size: 9.2pt;
    }}
    th, td {{
      border: 1px solid #777;
      padding: 4pt 5pt;
    }}
    th {{
      background: #f0f0f0;
    }}
    code {{
      font-family: "DejaVu Sans Mono", monospace;
      font-size: 9pt;
      background: #f7f7f7;
      padding: 0 2pt;
    }}
  </style>
</head>
<body>
{body}
</body>
</html>
"""


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    markdown_text = REPORT_MD.read_text(encoding="utf-8")
    html_body = markdown.markdown(markdown_text, extensions=["tables"])
    html = HTML_TEMPLATE.format(body=html_body)
    (OUTPUT_DIR / "case-study-report.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    main()
