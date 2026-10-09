import argparse
import json
import math
import re
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path


TEXT_TYPES = {"TEXT", "MTEXT", "ATTRIB", "ATTDEF", "DIMENSION", "MLEADER", "LEADER"}
SKIP_TYPES = {"VIEWPORT", "SEQEND"}


def decode_codes(codes):
    if not isinstance(codes, list) or not codes:
        return ""
    values = [int(value) for value in codes]
    direct = "".join(chr(value) for value in values)
    if any(value > 255 for value in values):
        return direct
    try:
        legacy = bytes(values).decode("cp950")
    except UnicodeDecodeError:
        return direct
    legacy_cjk = sum("㐀" <= char <= "鿿" for char in legacy)
    direct_cjk = sum("㐀" <= char <= "鿿" for char in direct)
    return legacy if legacy_cjk > direct_cjk else direct


def load(path):
    rows = []
    raw = Path(path).read_bytes()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("cp950", errors="replace")
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            row = json.loads(line)
        except json.JSONDecodeError:
            continue
        for key in ("layer", "space", "text", "block"):
            code_key = f"{key}_codes"
            if code_key in row:
                row[key] = decode_codes(row.pop(code_key))
        if row.get("type") not in SKIP_TYPES:
            rows.append(row)
    return rows


def clean(value):
    value = value or ""
    value = re.sub(r"\\[A-Za-z][^;]*;", "", value)
    value = value.replace("\\P", " ").replace("\n", " ").replace("\r", " ")
    value = value.replace("{", "").replace("}", "")
    return re.sub(r"\s+", " ", value).strip()


def rounded(value, digits=2):
    return None if value is None else round(float(value), digits)


def content_key(row):
    return (
        row.get("type", ""), row.get("layer", ""), clean(row.get("text")),
        row.get("block", ""), row.get("shape", ""),
        rounded(row.get("rotation"), 4), rounded(row.get("measurement"), 3),
    )


def point(row):
    x, y = row.get("x"), row.get("y")
    return None if x is None or y is None else (float(x), float(y))


def distance(a, b):
    pa, pb = point(a), point(b)
    if pa is None or pb is None:
        return float("inf")
    return math.dist(pa, pb)


def pair_identical(old, new):
    new_by_key = defaultdict(list)
    for index, row in enumerate(new):
        new_by_key[content_key(row)].append(index)
    used, pairs, remaining_old = set(), [], []
    for old_row in old:
        candidates = [i for i in new_by_key.get(content_key(old_row), []) if i not in used]
        if not candidates:
            remaining_old.append(old_row)
            continue
        same_space = [i for i in candidates if new[i].get("space") == old_row.get("space")]
        best = min(same_space or candidates, key=lambda i: distance(old_row, new[i]))
        used.add(best)
        pairs.append((old_row, new[best]))
    return pairs, remaining_old, [row for i, row in enumerate(new) if i not in used]


def is_text(row):
    return row.get("type") in TEXT_TYPES or bool(clean(row.get("text")))


def is_direct_space(row):
    space = row.get("space", "")
    return (
        space.startswith("*Model_Space")
        or space.startswith("*Paper_Space")
        or space.startswith("Layout:")
    )


def pair_modified_text(old, new, max_distance):
    candidates = []
    for oi, old_row in enumerate(old):
        if not is_text(old_row):
            continue
        for ni, new_row in enumerate(new):
            if not is_text(new_row):
                continue
            if old_row.get("type") != new_row.get("type") or old_row.get("layer") != new_row.get("layer"):
                continue
            if old_row.get("space") != new_row.get("space"):
                continue
            item_distance = distance(old_row, new_row)
            if math.isfinite(item_distance) and item_distance > max_distance:
                continue
            similarity = SequenceMatcher(None, clean(old_row.get("text")), clean(new_row.get("text"))).ratio()
            if similarity < 0.1 and item_distance > max_distance / 4:
                continue
            cost = (item_distance / max_distance if math.isfinite(item_distance) else 1) + (1 - similarity)
            candidates.append((cost, oi, ni))
    candidates.sort()
    used_old, used_new, pairs = set(), set(), []
    for _, oi, ni in candidates:
        if oi in used_old or ni in used_new:
            continue
        used_old.add(oi)
        used_new.add(ni)
        pairs.append((old[oi], new[ni]))
    return pairs, [r for i, r in enumerate(old) if i not in used_old], [r for i, r in enumerate(new) if i not in used_new]


def display_space(space):
    if space == "*Model_Space":
        return "模型空間"
    if space.startswith("*Paper_Space"):
        return "圖紙空間"
    if space.startswith("Layout:"):
        return f"圖紙空間 {space[7:]}"
    return f"圖塊 {space}"


def category(row):
    entity_type, layer = row.get("type", ""), row.get("layer", "").upper()
    if entity_type == "DIMENSION" or "DIM" in layer:
        return "尺寸"
    if entity_type == "INSERT":
        return "圖塊"
    if is_text(row):
        return "文字／標註"
    if entity_type in {"HATCH", "SOLID"}:
        return "填充"
    return "幾何圖元"


def subject(row):
    return clean(row.get("text")) or row.get("block") or row.get("type", "圖元")


def insertion_index(rows):
    result = defaultdict(list)
    for row in rows:
        if row.get("type") == "INSERT" and row.get("block"):
            result[row["block"]].append(row)
    return result


def callout(row, inserts):
    p, space = point(row), row.get("space", "")
    if space.startswith("*Model_Space") or space.startswith("*Paper_Space") or space.startswith("Layout:"):
        return p
    candidates = inserts.get(space, [])
    return point(candidates[0]) if candidates else None


def extent(rows):
    points = [point(r) for r in rows]
    points = [p for p in points if p and all(math.isfinite(v) for v in p)]
    if not points:
        return None
    xs, ys = [p[0] for p in points], [p[1] for p in points]
    return [min(xs), min(ys), max(xs), max(ys)]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("old")
    parser.add_argument("new")
    parser.add_argument("--output", required=True)
    parser.add_argument("--max-text-distance", type=float, default=600.0)
    parser.add_argument("--max-items", type=int, default=250)
    args = parser.parse_args()
    old_all, new_all = load(args.old), load(args.new)
    identical, old_remaining, new_remaining = pair_identical(old_all, new_all)
    moved = [(a, b) for a, b in identical if distance(a, b) > 0.01]
    modified, old_remaining, new_remaining = pair_modified_text(old_remaining, new_remaining, args.max_text_distance)
    old_inserts, new_inserts = insertion_index(old_all), insertion_index(new_all)
    changes = []

    def add(row, description, kind, location_row=None):
        target = location_row or row
        p = callout(target, new_inserts if kind != "removed" else old_inserts)
        changes.append({
            "item": 0, "drawing_area": display_space(row.get("space", "")),
            "location": f"X={target.get('x'):.2f}, Y={target.get('y'):.2f}" if target.get("x") is not None and target.get("y") is not None else "未定位",
            "category": category(row), "description": description, "notes": "",
            "kind": kind, "layer": row.get("layer", ""),
            "callouts": [{"x": p[0], "y": p[1]}] if p else [],
        })

    direct_modified = [(old_row, new_row) for old_row, new_row in modified if is_direct_space(new_row)]
    block_modified = [(old_row, new_row) for old_row, new_row in modified if not is_direct_space(new_row)]

    for old_row, new_row in direct_modified:
        old_subject, new_subject = subject(old_row), subject(new_row)
        if old_subject != new_subject:
            description = f"由「{old_subject}」修改為「{new_subject}」。"
        elif rounded(old_row.get("measurement"), 3) != rounded(new_row.get("measurement"), 3):
            description = f"尺寸由 {old_row.get('measurement')} 修改為 {new_row.get('measurement')}。"
        else:
            description = f"{new_subject} 的屬性或幾何定義已修改。"
        add(new_row, description, "modified")

    block_modified_groups = defaultdict(list)
    for old_row, new_row in block_modified:
        key = (new_row.get("space"), new_row.get("layer"), new_row.get("type"))
        block_modified_groups[key].append((old_row, new_row))
    for key, pairs in sorted(block_modified_groups.items()):
        samples = []
        for old_row, new_row in pairs[:3]:
            old_subject, new_subject = subject(old_row), subject(new_row)
            samples.append(f"{old_subject}→{new_subject}" if old_subject != new_subject else new_subject)
        detail = "；".join(samples)
        suffix = f"（例：{detail}）" if detail else ""
        add(pairs[0][1], f"圖塊內 {key[2]} 修改 {len(pairs)} 處{suffix}。", "modified")

    for row in new_remaining:
        if is_direct_space(row) and (is_text(row) or row.get("type") == "INSERT"):
            add(row, f"新增「{subject(row)}」。", "added")
    for row in old_remaining:
        if is_direct_space(row) and (is_text(row) or row.get("type") == "INSERT"):
            add(row, f"刪除「{subject(row)}」。", "removed")

    old_geometry = [
        r for r in old_remaining
        if is_direct_space(r) and not is_text(r) and r.get("type") != "INSERT"
    ]
    new_geometry = [
        r for r in new_remaining
        if is_direct_space(r) and not is_text(r) and r.get("type") != "INSERT"
    ]
    old_groups = defaultdict(list)
    new_groups = defaultdict(list)
    for row in old_geometry:
        old_groups[(row.get("space"), row.get("layer"))].append(row)
    for row in new_geometry:
        new_groups[(row.get("space"), row.get("layer"))].append(row)
    for key in sorted(set(old_groups) | set(new_groups)):
        removed_rows, added_rows = old_groups[key], new_groups[key]
        sample_row = (added_rows or removed_rows)[0]
        type_counts = []
        for entity_type in sorted({r.get("type") for r in removed_rows + added_rows}):
            removed_count = sum(r.get("type") == entity_type for r in removed_rows)
            added_count = sum(r.get("type") == entity_type for r in added_rows)
            type_counts.append(f"{entity_type} +{added_count}/-{removed_count}")
        add(sample_row, f"幾何圖元變更：{'、'.join(type_counts)}。", "geometry")

    old_block = [r for r in old_remaining if not is_direct_space(r)]
    new_block = [r for r in new_remaining if not is_direct_space(r)]
    old_block_groups = defaultdict(list)
    new_block_groups = defaultdict(list)
    for row in old_block:
        old_block_groups[row.get("space")].append(row)
    for row in new_block:
        new_block_groups[row.get("space")].append(row)
    for key in sorted(set(old_block_groups) | set(new_block_groups)):
        removed_rows, added_rows = old_block_groups[key], new_block_groups[key]
        sample_row = (added_rows or removed_rows)[0]
        sample_subjects = [subject(row) for row in (added_rows + removed_rows) if is_text(row)]
        sample_subjects = list(dict.fromkeys(sample_subjects))[:3]
        detail = f"（例：{'；'.join(sample_subjects)}）" if sample_subjects else ""
        type_counts = []
        for entity_type in sorted({r.get("type") for r in removed_rows + added_rows}):
            removed_count = sum(r.get("type") == entity_type for r in removed_rows)
            added_count = sum(r.get("type") == entity_type for r in added_rows)
            type_counts.append(f"{entity_type} +{added_count}/-{removed_count}")
        add(
            sample_row,
            f"圖塊內容變更：{'、'.join(type_counts)}{detail}。",
            "geometry",
        )

    merged_changes = []
    merged_by_key = {}
    for change in changes:
        key = (
            change["drawing_area"], change["kind"], change["layer"],
            change["category"], change["description"],
        )
        existing = merged_by_key.get(key)
        if existing is None:
            merged_by_key[key] = change
            merged_changes.append(change)
            continue
        existing["callouts"].extend(change["callouts"])
        existing["callouts"] = [
            {"x": x, "y": y}
            for x, y in dict.fromkeys((p["x"], p["y"]) for p in existing["callouts"])
        ]
        existing["location"] = f"共 {len(existing['callouts'])} 處（見項次標示快照）"

    changes = merged_changes[: args.max_items]
    for index, row in enumerate(changes, 1):
        row["item"] = index
    output = {"meta": {
        "old_fact_count": len(old_all), "new_fact_count": len(new_all),
        "identical_count": len(identical), "pure_move_count": len(moved),
        "candidate_count": len(changes), "extent": extent(new_all), "review_required": True,
    }, "changes": changes}
    Path(args.output).write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(output["meta"], ensure_ascii=False))


if __name__ == "__main__":
    main()
