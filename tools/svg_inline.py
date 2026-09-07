"""Inline a <style> block's class rules as presentation attributes.

The governorate icons from syrian.zone style every shape through CSS classes.
PyMuPDF's SVG reader ignores CSS, so without this every icon rasterises solid
black. This flattens `.st0 { fill: #04018c }` + `class="st0"` into `fill="..."`.
"""
from __future__ import annotations

import re

PRESENTATION_PROPS = {
    "fill", "stroke", "stroke-width", "stroke-linecap", "stroke-linejoin",
    "stroke-dasharray", "stroke-opacity", "fill-opacity", "opacity",
    "fill-rule", "stroke-miterlimit",
}

_STYLE_BLOCK = re.compile(r"<style[^>]*>(.*?)</style>", re.S | re.I)
_RULE = re.compile(r"([^{}]+)\{([^{}]*)\}", re.S)


def parse_rules(css: str) -> dict[str, dict[str, str]]:
    """class name -> {prop: value}. Later rules win, matching CSS cascade."""
    out: dict[str, dict[str, str]] = {}
    for selectors, body in _RULE.findall(css):
        props = {}
        for decl in body.split(";"):
            if ":" not in decl:
                continue
            prop, _, value = decl.partition(":")
            prop, value = prop.strip().lower(), value.strip()
            if prop in PRESENTATION_PROPS and value:
                props[prop] = value
        if not props:
            continue
        for sel in selectors.split(","):
            sel = sel.strip()
            if sel.startswith("."):
                out.setdefault(sel[1:], {}).update(props)
    return out


def inline_styles(svg: str) -> str:
    """Rewrite class="..." into explicit attributes and drop the <style> block."""
    blocks = _STYLE_BLOCK.findall(svg)
    if not blocks:
        return svg

    rules = parse_rules("\n".join(blocks))
    if not rules:
        return _STYLE_BLOCK.sub("", svg)

    def apply(match: re.Match) -> str:
        tag = match.group(0)
        names = match.group(1).split()
        props: dict[str, str] = {}
        for n in names:
            props.update(rules.get(n, {}))
        if not props:
            return tag
        # An existing inline attribute is more specific than a class - keep it.
        attrs = " ".join(
            f'{k}="{v}"' for k, v in props.items()
            if not re.search(rf'\b{re.escape(k)}\s*=', tag)
        )
        if not attrs:
            return tag
        body = tag[:-1].rstrip()          # strip the closing ">"
        selfclose = body.endswith("/")     # <path ... /> must stay self-closing
        if selfclose:
            body = body[:-1].rstrip()
        return f'{body} {attrs}{"/>" if selfclose else ">"}'

    svg = re.sub(r'<[^>]*\bclass="([^"]+)"[^>]*>', apply, svg)
    return _STYLE_BLOCK.sub("", svg)
