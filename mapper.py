"""
Content mapper: transforms raw extracted slide data into
the structured dicts that template builders expect.
"""

import re


def _first_title(slide):
    """Get the most likely title from a slide's text boxes.
    Priority: 1) placeholder_idx==0, 2) largest font, 3) topmost box."""
    boxes = slide.get("raw_boxes", [])
    if not boxes:
        texts = slide.get("all_text", [])
        return texts[0] if texts else "Untitled"

    # First: look for an explicit title placeholder (idx 0)
    for b in boxes:
        if b.get("placeholder_idx") == 0 and b["paragraphs"]:
            return b["paragraphs"][0]

    # Second: topmost single-paragraph box (likely a title)
    top_boxes = sorted(boxes, key=lambda b: b.get("top", 0) or 0)
    for b in top_boxes:
        if b["para_count"] == 1 and b["paragraphs"]:
            return b["paragraphs"][0]

    # Third: largest font
    titled = sorted(boxes, key=lambda b: b.get("max_font_size", 0), reverse=True)
    if titled and titled[0]["max_font_size"] > 16:
        return titled[0]["paragraphs"][0]

    # Fallback
    texts = slide.get("all_text", [])
    return texts[0] if texts else "Untitled"


def _title_box(slide):
    """Return the box dict identified as the title (for exclusion)."""
    boxes = slide.get("raw_boxes", [])
    if not boxes:
        return None

    # Placeholder idx 0
    for b in boxes:
        if b.get("placeholder_idx") == 0 and b["paragraphs"]:
            return b

    # Topmost single-paragraph box
    top_boxes = sorted(boxes, key=lambda b: b.get("top", 0) or 0)
    for b in top_boxes:
        if b["para_count"] == 1 and b["paragraphs"]:
            return b

    # Largest font
    titled = sorted(boxes, key=lambda b: b.get("max_font_size", 0), reverse=True)
    if titled and titled[0]["max_font_size"] > 16:
        return titled[0]

    return None


def _body_texts(slide, skip_first=True):
    """Get body text items, skipping the title box's paragraphs."""
    texts = slide.get("all_text", [])
    if not skip_first:
        return texts

    tb = _title_box(slide)
    if tb:
        title_paras = set(tb["paragraphs"])
        return [t for t in texts if t not in title_paras]

    # Fallback: skip first item
    return texts[1:] if len(texts) > 1 else texts


def _split_pairs(texts):
    """Try to split texts into finding/rec pairs based on → or similar."""
    pairs = []
    for t in texts:
        if "→" in t or "->" in t:
            parts = re.split(r'\s*[→\->]+\s*', t, maxsplit=1)
            if len(parts) == 2:
                pairs.append({"finding": parts[0].strip(), "recommendation": parts[1].strip()})
                continue
        pairs.append({"finding": t.strip(), "recommendation": ""})
    return pairs


def _split_kv(texts):
    """Try to split texts into key:value pairs."""
    fields = []
    for t in texts:
        if ":" in t:
            parts = t.split(":", maxsplit=1)
            fields.append({"label": parts[0].strip(), "value": parts[1].strip()})
        else:
            fields.append({"label": "", "value": t.strip()})
    return fields


def _find_big_number(slide):
    """Find a big standalone number/percentage."""
    for box in slide.get("raw_boxes", []):
        text = box["text"].strip()
        if re.match(r'^[\d,.]+[%×xX]?$', text):
            return text
    # Fallback: look for numbers in text
    texts = slide.get("all_text", [])
    for t in texts:
        m = re.search(r'(\d+[%×xX])', t)
        if m:
            return m.group(1)
    return "—"


def _extract_quote(slide):
    """Extract quote text and attribution."""
    texts = slide.get("all_text", [])
    quote = ""
    attribution = ""
    context = ""

    for t in texts:
        # Look for text in quotes
        m = re.search(r'["\u201c](.+?)["\u201d]', t, re.DOTALL)
        if m and len(m.group(1)) > len(quote):
            quote = m.group(1)
        elif "\u2014" in t or "—" in t:
            attribution = t.replace("\u2014", "").replace("—", "").strip()
        elif t.startswith("-") or t.startswith("–"):
            attribution = t.lstrip("-–").strip()

    if not quote:
        # Just use the longest text as the quote
        body = _body_texts(slide)
        if body:
            quote = max(body, key=len)

    return quote, attribution, context


def _split_columns(slide):
    """Try to detect left/right column content based on text box positions."""
    boxes = slide.get("raw_boxes", [])
    # Exclude the title box
    tb = _title_box(slide)
    body_boxes = [b for b in boxes if b is not tb] if tb else boxes

    if len(body_boxes) < 2:
        texts = _body_texts(slide)
        mid = len(texts) // 2
        return texts[:mid] or [""], texts[mid:] or [""]

    # Sort by horizontal position
    sorted_boxes = sorted(body_boxes, key=lambda b: b.get("left", 0) or 0)
    mid_x = sum(b.get("left", 0) or 0 for b in sorted_boxes) / len(sorted_boxes)

    left = []
    right = []
    for b in sorted_boxes:
        x = b.get("left", 0) or 0
        for p in b["paragraphs"]:
            if x < mid_x:
                left.append(p)
            else:
                right.append(p)

    return left or [""], right or [""]


def map_slide(slide, slide_type):
    """Map extracted slide data to the template-ready dict for a given type."""
    title = _first_title(slide)
    body = _body_texts(slide)
    texts = slide.get("all_text", [])

    if slide_type == "title":
        subtitle = body[0] if body else ""
        author = body[1] if len(body) > 1 else ""
        date_str = body[2] if len(body) > 2 else ""
        return {
            "title": title,
            "subtitle": subtitle,
            "author": author,
            "date": date_str,
        }

    elif slide_type == "closer":
        subtitle = body[0] if body else ""
        contact = body[1] if len(body) > 1 else ""
        return {
            "title": title,
            "subtitle": subtitle,
            "contact": contact,
        }

    elif slide_type == "section_divider":
        subtitle = body[0] if body else ""
        return {
            "title": title,
            "subtitle": subtitle,
            "sectionNumber": None,
        }

    elif slide_type == "agenda":
        items = []
        for t in body:
            if ":" in t or "\t" in t or "  " in t:
                parts = re.split(r'[:\t]|\s{2,}', t, maxsplit=1)
                if len(parts) == 2:
                    items.append({"title": parts[0].strip(), "detail": parts[1].strip()})
                    continue
            items.append({"title": t, "detail": ""})
        return {"title": title, "items": items}

    elif slide_type == "in_brief":
        return {"title": title, "bullets": body}

    elif slide_type == "stat_callout":
        stat = _find_big_number(slide)
        remaining = [t for t in body if t.strip() != stat]
        headline = remaining[0] if remaining else ""
        detail = remaining[1] if len(remaining) > 1 else ""
        source = remaining[2] if len(remaining) > 2 else ""
        return {
            "title": title,
            "stat": stat,
            "headline": headline,
            "detail": detail,
            "source": source,
        }

    elif slide_type == "quote":
        quote, attribution, context = _extract_quote(slide)
        return {
            "title": title,
            "quote": quote,
            "attribution": attribution,
            "context": context,
        }

    elif slide_type == "comparison":
        left_items, right_items = _split_columns(slide)
        return {
            "title": title,
            "leftLabel": left_items[0] if left_items else "Option A",
            "rightLabel": right_items[0] if right_items else "Option B",
            "leftItems": left_items[1:] if len(left_items) > 1 else left_items,
            "rightItems": right_items[1:] if len(right_items) > 1 else right_items,
        }

    elif slide_type == "text_graph":
        # Can't extract chart data from python-pptx easily,
        # so we pass the text and a placeholder chart
        return {
            "title": title,
            "text": body,
            "chartType": "bar",
            "chartData": [{"name": "Series 1", "labels": ["A", "B", "C"], "values": [30, 50, 40]}],
            "chartColors": ["E5E5E5", "3880F3", "368727"],
            "note": "(Chart data from original not transferred — update manually)",
        }

    elif slide_type == "process_flow":
        steps = []
        for t in body:
            # Try to split "Step title: detail"
            if ":" in t:
                parts = t.split(":", maxsplit=1)
                steps.append({"title": parts[0].strip(), "detail": parts[1].strip()})
            elif ". " in t[:20]:
                parts = t.split(". ", maxsplit=1)
                steps.append({"title": parts[0].strip(), "detail": parts[1].strip()})
            else:
                steps.append({"title": t, "detail": ""})
        return {"title": title, "steps": steps[:5]}

    elif slide_type == "matrix":
        quads = []
        for t in body[:4]:
            if ":" in t:
                parts = t.split(":", maxsplit=1)
                quads.append({"label": parts[0].strip(), "detail": parts[1].strip()})
            else:
                quads.append({"label": t[:30], "detail": t})
        while len(quads) < 4:
            quads.append({"label": "", "detail": ""})
        return {
            "title": title,
            "quadrants": quads[:4],
            "xAxis": "",
            "yAxis": "",
        }

    elif slide_type == "methods":
        fields = _split_kv(body)
        return {"title": title, "fields": fields}

    elif slide_type == "hypotheses":
        hyps = []
        for t in body:
            status = ""
            for s in ["Confirmed", "Rejected", "Partial", "Supported", "Not Supported"]:
                if s.lower() in t.lower():
                    status = s
                    t = re.sub(re.escape(s), "", t, flags=re.IGNORECASE).strip(" -–—:•")
                    break
            hyps.append({"text": t, "status": status})
        return {"title": title, "hypotheses": hyps}

    elif slide_type in ("wsn_dense", "wsn_reveal"):
        # Try to split into What / So What / Now What sections
        what = {"headline": "", "detail": ""}
        so_what = {"headline": "", "detail": ""}
        now_what = {"headline": "", "detail": ""}

        current = what
        for t in body:
            tl = t.lower().strip()
            if tl.startswith("so what") or tl.startswith("sowhat"):
                current = so_what
                t = re.sub(r'^so\s*what\s*[:—\-]?\s*', '', t, flags=re.IGNORECASE)
            elif tl.startswith("now what") or tl.startswith("nowwhat"):
                current = now_what
                t = re.sub(r'^now\s*what\s*[:—\-]?\s*', '', t, flags=re.IGNORECASE)
            elif tl.startswith("what"):
                t = re.sub(r'^what\s*[:—\-]?\s*', '', t, flags=re.IGNORECASE)

            if not current["headline"]:
                current["headline"] = t
            else:
                current["detail"] = (current["detail"] + " " + t).strip()

        # If we didn't find explicit sections, split evenly
        if not so_what["headline"] and not now_what["headline"]:
            n = len(body)
            if n >= 3:
                third = n // 3
                what = {"headline": body[0], "detail": " ".join(body[1:third])}
                so_what = {"headline": body[third], "detail": " ".join(body[third+1:2*third])}
                now_what = {"headline": body[2*third], "detail": " ".join(body[2*third+1:])}
            elif n == 2:
                what = {"headline": body[0], "detail": ""}
                so_what = {"headline": body[1], "detail": ""}
            elif n == 1:
                what = {"headline": body[0], "detail": ""}

        return {
            "title": title,
            "what": what,
            "soWhat": so_what,
            "nowWhat": now_what,
        }

    elif slide_type == "findings_recs":
        pairs = _split_pairs(body)
        return {"title": title, "items": pairs[:5]}

    elif slide_type == "findings_recs_dense":
        pairs = _split_pairs(body)
        return {"title": title, "items": pairs[:8]}

    elif slide_type == "open_questions":
        questions = [t for t in body if t.strip().endswith("?")]
        if not questions:
            questions = body[:4]
        return {"title": title, "questions": questions[:4]}

    elif slide_type == "progressive_reveal":
        takeaways = []
        for t in body:
            if ":" in t:
                parts = t.split(":", maxsplit=1)
                takeaways.append({
                    "headline": parts[0].strip(),
                    "detail": parts[1].strip(),
                    "summary": parts[0].strip()[:60],
                })
            else:
                takeaways.append({
                    "headline": t,
                    "detail": "",
                    "summary": t[:60],
                })
        return {"title": title, "takeaways": takeaways[:5]}

    # Fallback
    return {"title": title, "bullets": body}
