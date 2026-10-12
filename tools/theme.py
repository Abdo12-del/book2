#!/usr/bin/env python3
"""تحميل نظام التصميم (design/theme.yaml) وتحويل وحداته إلى نقاط."""
from __future__ import annotations

import os

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MM = 72.0 / 25.4  # 1mm = 2.8346pt

PAGE_SIZES = {  # بالمليمتر (عرض × ارتفاع)
    "A5": (148, 210),
    "B5": (176, 250),
    "A4": (210, 297),
    "B4": (250, 353),
    "A6": (105, 148),
    "LETTER": (216, 279),
}

DEFAULTS: dict = {
    "page": {"size": "A5", "custom_size_mm": [170, 240], "mirrored": True, "gutter_mm": 8,
             "margin_top_mm": 20, "margin_bottom_mm": 20, "margin_outer_mm": 15,
             "header_from_top_mm": 10, "footer_from_bottom_mm": 10},
    "direction": "rtl",
    # خطّ حديث غير أدبي: Noto Sans Arabic للعربية + Noto Sans للاتيني (تقسيم بالمحتوى)
    "fonts": {"body": "NotoSansArabic", "body_bold": "NotoSansArabic-Bold",
              "body_italic": "NotoSansArabic-Light",
              "heading": "NotoSansArabic-Bold", "heading_bold": "NotoSansArabic-Bold",
              "mono": "DejaVuSansMono", "latin": "NotoSans", "latin_bold": "NotoSans-Bold",
              "latin_italic": "NotoSans-Italic"},
    "sizes": {"body_pt": 11.5, "body_leading_pt": 19, "h1_pt": 20, "h2_pt": 15,
              "h3_pt": 12.5, "h4_pt": 11.5, "caption_pt": 10, "toc_pt": 11,
              "header_pt": 9, "footer_pt": 9.5, "table_pt": 10.5},
    "colors": {"text": "#1A1A1A", "heading": "#16324F", "accent": "#8C6D3F", "muted": "#6B7280",
               "rule": "#C9CDD3", "callout_bg": "#F5F1E8", "callout_border": "#8C6D3F",
               "table_header_bg": "#E8E5DC", "table_zebra_bg": "#F7F5F0"},
    "layout": {"space_before_h1_pt": 0, "space_after_h1_pt": 14, "space_before_h2_pt": 16,
               "space_after_h2_pt": 8, "space_before_h3_pt": 12, "space_after_h3_pt": 6,
               "space_after_p_pt": 8, "first_line_indent_pt": 0, "keep_headings_with_next": True,
               "chapter_starts_on": "right", "chapter_head_space_before_pt": 40},
    "running": {"odd_header": "chapter", "even_header": "title", "footer": "page_number",
                "page_number_style": "arabic-indic", "front_matter_numbering": "roman",
                "header_rule": True},
    "toc": {"enabled": True, "depth": 2, "dotted_leader": True},
    "objects": {"caption_prefix_table": "جدول", "caption_prefix_figure": "شكل",
                "caption_numbering": "chapter", "table_style": "grid",
                "image_max_width_ratio": 0.92},
}


def _deep_merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in (over or {}).items():
        out[k] = _deep_merge(base[k], v) if isinstance(v, dict) and isinstance(base.get(k), dict) else v
    return out


def load(path: str | None = None) -> dict:
    path = path or os.path.join(ROOT, "design", "theme.yaml")
    data = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    theme = _deep_merge(DEFAULTS, data)

    size_key = str(theme["page"]["size"]).upper()
    if size_key in PAGE_SIZES:
        w_mm, h_mm = PAGE_SIZES[size_key]
    else:
        w_mm, h_mm = theme["page"]["custom_size_mm"]
    theme["page"]["width_pt"] = w_mm * MM
    theme["page"]["height_pt"] = h_mm * MM
    theme["page"]["size_name"] = size_key
    return theme


def hex_to_rgb(h: str) -> tuple[float, float, float]:
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))  # type: ignore[return-value]


if __name__ == "__main__":
    t = load()
    print(f"الصفحة: {t['page']['size_name']}  {t['page']['width_pt']:.1f}×{t['page']['height_pt']:.1f}pt")
    print(f"المتن: {t['fonts']['body']} {t['sizes']['body_pt']}pt / تباعد {t['sizes']['body_leading_pt']}pt")
