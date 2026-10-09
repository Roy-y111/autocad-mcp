import argparse
import json
import math
import re
import statistics
from collections import Counter, defaultdict
from pathlib import Path

from compare_facts import clean, is_direct_space, load


SYSTEM_BLOCKS = {"ARROW", "DOT", "SLASH", "_ARCHTICK", "圖框"}
GENERIC_SUBJECTS = {
    "DIMENSION", "LEADER", "MULTILEADER", "{<>}", "<>", "X", "Y", "UP",
    "3~6", "3~7BEAM",
}
THEME_ORDER = {
    "revision": 0,
    "general": 1,
    "axis_dim": 2,
    "column": 3,
    "beam": 4,
    "slab": 5,
    "elevation": 6,
    "wall_opening": 7,
    "detail": 8,
    "other": 9,
}
THEME_META = {
    "revision": ("修訂註記", "各張圖名／修訂欄"),
    "general": ("一般說明", "一般附註／圖說文字"),
    "axis_dim": ("軸線／尺寸", "全平面及外圍尺寸鏈"),
    "column": ("柱", "柱位（多處）"),
    "beam": ("梁", "梁線及梁標註（多處）"),
    "slab": ("版／基礎／開口", "各版區、基礎及開口附近"),
    "elevation": ("高程註記", "高程註記附近"),
    "wall_opening": ("牆／開口／幾何", "牆線、開口及填充範圍"),
    "detail": ("詳圖", "新增／調整詳圖位置"),
    "other": ("其他註記", "圖說標註位置"),
}


def chinese_number(value):
    digits = {"零": 0, "〇": 0, "一": 1, "二": 2, "三": 3, "四": 4,
              "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    if value.isdigit():
        return int(value)
    if value == "十":
        return 10
    if "十" in value:
        left, right = value.split("十", 1)
        return (digits.get(left, 1) * 10) + digits.get(right, 0)
    return digits.get(value)


def floor_label(text):
    value = clean(text).upper().replace("－", "-").replace("～", "~")
    if "屋突頂層" in value or "屋突頂版" in value:
        return "屋突頂層"
    match = re.search(r"屋突(?:第)?([零〇一二三四五六七八九十\d]+)層", value)
    if match:
        number = chinese_number(match.group(1))
        return f"R{number}F" if number else None
    match = re.search(r"地下([零〇一二三四五六七八九十\d]+)層", value)
    if match:
        number = chinese_number(match.group(1))
        return f"B{number}F" if number else None
    match = re.search(r"([零〇一二三四五六七八九十\d]+)\s*(?:至|~|-)\s*([零〇一二三四五六七八九十\d]+)層", value)
    if match:
        start, end = chinese_number(match.group(1)), chinese_number(match.group(2))
        return f"{start}F–{end}F" if start is not None and end is not None else None
    match = re.search(r"(?:^|[^A-Z0-9])B(\d+)F(?:L)?(?:[^A-Z0-9]|$)", value)
    if match:
        return f"B{int(match.group(1))}F"
    match = re.search(r"(?:^|[^A-Z0-9])R(\d+)F(?:L)?(?:[^A-Z0-9]|$)", value)
    if match:
        return f"R{int(match.group(1))}F"
    match = re.search(r"(?:^|[^A-Z0-9])(\d+)F(?:L)?(?:[^A-Z0-9]|$)", value)
    if match:
        return f"{int(match.group(1))}F"
    match = re.search(r"([零〇一二三四五六七八九十]+)層", value)
    if match:
        number = chinese_number(match.group(1))
        return f"{number}F" if number is not None else None
    return None


def text_height(row):
    if row.get("text_height") is not None:
        return float(row["text_height"])
    match = re.search(r"(?:^|;)40:([-+0-9.]+)", row.get("shape", ""))
    return float(match.group(1)) if match else None


def find_regions(facts):
    candidates = []
    for row in facts:
        text = clean(row.get("text"))
        label = floor_label(text)
        height = text_height(row)
        if not label or row.get("space") != "*Model_Space" or row.get("type") != "TEXT":
            continue
        if not re.search(r"平面圖|配置圖|\bPLAN\b", text, re.I):
            continue
        if row.get("x") is None or row.get("y") is None or not height:
            continue
        candidates.append({"label": label, "x": float(row["x"]), "y": float(row["y"]), "height": height})
    if not candidates:
        return []
    rounded_heights = Counter(round(row["height"], 3) for row in candidates)
    modal_height = sorted(rounded_heights, key=lambda value: (-rounded_heights[value], abs(value - 100)))[0]
    selected = [row for row in candidates if abs(row["height"] - modal_height) <= max(1.0, modal_height * 0.12)]
    deduplicated = []
    for row in sorted(selected, key=lambda item: (item["y"], item["x"])):
        if not any(row["label"] == prior["label"] and math.dist((row["x"], row["y"]), (prior["x"], prior["y"])) < modal_height for prior in deduplicated):
            deduplicated.append(row)
    return deduplicated


def region_spacing(regions):
    distances = []
    for index, region in enumerate(regions):
        others = [math.dist((region["x"], region["y"]), (other["x"], other["y"])) for offset, other in enumerate(regions) if offset != index]
        if others:
            distances.append(min(others))
    return statistics.median(distances) if distances else float("inf")


def assign_region(point, regions, spacing):
    if not regions or point is None:
        return "模型空間"
    nearest = min(regions, key=lambda row: math.dist(point, (row["x"], row["y"])))
    distance = math.dist(point, (nearest["x"], nearest["y"]))
    if math.isfinite(spacing) and distance > spacing * 1.55:
        return "全圖／其他詳圖"
    return nearest["label"]


def normalize_subject(value):
    result = clean(value).replace("%%U", "").replace("'", "′")
    result = re.sub(r"^\{(.+)\}$", r"\1", result)
    result = result.strip(" 。；")
    if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", result):
        number = float(result)
        return f"{number:.3f}".rstrip("0").rstrip(".")
    return result


def subject_is_generic(value):
    normalized = normalize_subject(value)
    return (
        normalized.upper() in GENERIC_SUBJECTS
        or re.fullmatch(r"\d+(?:\.\d+)?", normalized) is not None
        or normalized.upper().startswith("A$")
    )


def internal_noise(change):
    area = change.get("drawing_area", "")
    block = area.removeprefix("圖塊 ").strip().upper()
    description = change.get("description", "")
    if area.startswith("圖塊 A$") or block in SYSTEM_BLOCKS:
        return True
    if "「A$" in description:
        return True
    if any(f"「{name}」" in description for name in ("3~6", "3~7BEAM")):
        return True
    if "新增「X」" in description or "新增「Y」" in description or "刪除「UP」" in description:
        return True
    return False


def meaningful_change(change):
    description = change.get("description", "")
    match = re.search(r"(?:新增|刪除)「(.+?)」", description)
    if match and subject_is_generic(match.group(1)):
        return False
    if re.match(r"^(?:DIMENSION|LEADER|MULTILEADER|圖例：).+的屬性或幾何定義已修改", description):
        return False
    return True


def theme_for(change):
    description = change.get("description", "")
    layer = change.get("layer", "").upper()
    area = change.get("drawing_area", "")
    upper = description.upper()
    if "第1次變更設計" in description or "修訂" in description:
        return "revision"
    if any(token in description for token in ("混凝土強度", "設計活載重", "一般說明", "kgf/cm", "施工")):
        return "general"
    if "詳圖" in description or "詳圖" in area:
        return "detail"
    if change.get("kind") == "geometry" or change.get("category") in {"幾何圖元", "填充"}:
        return "wall_opening"
    if re.search(r"[「\s][+-]\d", description) or "ELEV" in layer:
        return "elevation"
    if "COL" in layer or re.search(r"(?:^|[^A-Z])(?:CF|SC|C)\d", upper):
        return "column"
    if "SLAB" in layer or re.search(r"(?:^|[^A-Z])(?:BS|FS|CS|CES|ES|S)\d", upper):
        return "slab"
    if "BEAM" in layer or re.search(r"(?:^|[^A-Z])(?:WG|SG|SB|CB|BX|B|G)\d", upper):
        return "beam"
    if change.get("category") == "尺寸" or "GRID" in layer or "DIM" in layer:
        return "axis_dim"
    return "other"


def evidence_values(descriptions):
    pairs, added, removed = [], [], []
    for description in descriptions:
        match = re.search(r"由「(.+?)」修改為「(.+?)」", description)
        if match:
            old, new = normalize_subject(match.group(1)), normalize_subject(match.group(2))
            if old and new and old != new:
                pairs.append((old, new))
            continue
        match = re.search(r"尺寸由\s*([-+0-9.]+)\s*修改為\s*([-+0-9.]+)", description)
        if match:
            pairs.append((normalize_subject(match.group(1)), normalize_subject(match.group(2))))
            continue
        match = re.search(r"新增「(.+?)」", description)
        if match and not subject_is_generic(match.group(1)):
            added.append(normalize_subject(match.group(1)))
            continue
        match = re.search(r"刪除「(.+?)」", description)
        if match and not subject_is_generic(match.group(1)):
            removed.append(normalize_subject(match.group(1)))
    return list(dict.fromkeys(pairs)), list(dict.fromkeys(added)), list(dict.fromkeys(removed))


def summarize(theme, evidence):
    descriptions = [row["description"] for row in evidence]
    combined = " ".join(descriptions)
    if theme == "general":
        summaries = []
        if "FS版" in combined and "350kgf/cm" in combined and "5FL版" in combined:
            summaries.append("混凝土強度分界調整：350 kgf/cm² 適用上限由 2FL 版改為 5FL 版，280 kgf/cm² 適用起點同步調整")
        if "梁編號加「*」者表樓梯版下梁" in combined:
            summaries.append("設計活載重說明增列「梁編號加『*』者表樓梯版下梁」")
        if any(token in combined for token in ("連續壁部份", "SD420W", "CNS560")):
            summaries.append("一般附註增列連續壁、鋼筋材料強度及結構開孔協調等施工規定")
        if summaries:
            return "；".join(summaries) + "。"
    if theme == "detail":
        _, added, _ = evidence_values(descriptions)
        named_details = [value for value in added if "詳圖" in value]
        if named_details:
            return f"新增／調整 {'、'.join(named_details[:8])}；相關詳圖幾何及標註同步調整。"
    pairs, added, removed = evidence_values(descriptions)
    prefixes = {
        "revision": "修訂註記調整",
        "general": "一般說明調整",
        "axis_dim": "軸線與尺寸鏈調整",
        "column": "柱標註／尺寸調整",
        "beam": "梁標註與梁線調整",
        "slab": "版、基礎及開口標註調整",
        "elevation": "高程註記調整",
        "wall_opening": "牆線、開口及填充範圍調整",
        "detail": "詳圖內容調整",
        "other": "圖說標註調整",
    }
    details = []
    if pairs:
        details.append("；".join(f"{old}→{new}" for old, new in pairs[:8]))
    if added:
        details.append(f"新增 {'、'.join(added[:10])}")
    if removed:
        details.append(f"刪除 {'、'.join(removed[:10])}")
    has_geometry = any(row.get("kind") == "geometry" for row in evidence)
    if has_geometry:
        details.append("相關線位或幾何範圍有調整")
    if not details:
        specific = [description for description in descriptions if "的屬性或幾何定義已修改" not in description]
        if specific:
            details.append(normalize_subject(specific[0]))
        else:
            details.append("相關內容有調整，位置依項次標示快照")
    return f"{prefixes[theme]}：{'；'.join(details)}。"


LABEL_PATTERN = re.compile(r"^([A-Za-z一-鿿][A-Za-z0-9一-鿿′]*?)(?:[\(（]|$)")
AXIS_LABEL_PATTERN = re.compile(r"^[0-9]{0,2}[A-Za-z]?′?$")


def base_label(value):
    match = LABEL_PATTERN.match(value)
    label = match.group(1) if match else value
    return label.strip()


def distinct_base_labels(values, limit=12):
    result = []
    for value in values:
        label = base_label(value)
        if label and label not in result:
            result.append(label)
        if len(result) >= limit:
            break
    return result


GRID_LAYER = "S-GRID"
GRID_LETTER_PATTERN = re.compile(r"^[A-Z]′?$")
GRID_NUMBER_PATTERN = re.compile(r"^[0-9]+(?:-[0-9]+)?′?$")


def assign_region_by_y(y, regions):
    """Floor plans in this drawing convention stack vertically at a roughly
    constant X band, with the floor title sitting at the BOTTOM of each
    plan and its grid-axis labels running along the TOP margin — often
    1500-2500 units above the title, which can be closer in absolute
    distance to the *next* floor's title than to its own (nearest-title-y
    picked the wrong floor for e.g. B3F's and R1F's own axis numbers,
    which sit far enough above their title to be nearer B2F's/R2F's).
    Banding by "largest title y at or below this point" instead treats
    each floor as the open interval [its own title y, next title y) and
    is robust to that offset as long as labels don't overshoot into the
    next floor's own band, which they don't in this drawing."""
    if not regions:
        return "模型空間"
    ordered = sorted(regions, key=lambda row: row["y"])
    current = ordered[0]["label"]
    for region in ordered:
        if region["y"] <= y:
            current = region["label"]
        else:
            break
    return current


def transform_local_point(local_x, local_y, insert):
    """World coordinates of a block-local point for one INSERT instance.
    Our fact schema only captures Y-scale (DXF group 42) as `measurement`,
    not X-scale separately, so non-uniform scaling isn't reproduced — every
    INSERT observed in this drawing set uses rotation=0, scale=1 anyway."""
    rotation = math.radians(insert.get("rotation") or 0.0)
    scale = insert.get("measurement")
    scale = 1.0 if scale is None else float(scale)
    cos_r, sin_r = math.cos(rotation), math.sin(rotation)
    world_x = insert["x"] + scale * (local_x * cos_r - local_y * sin_r)
    world_y = insert["y"] + scale * (local_x * sin_r + local_y * cos_r)
    return world_x, world_y


def collect_grid_axes(facts, regions):
    """Bucket S-GRID axis labels per floor region: letters mark horizontal
    grid lines (keyed by their y), numbers mark vertical grid lines (keyed
    by their x). Lets wall/slab rows report a real 'near axis 2/C' location
    instead of a generic phrase, using coordinates the drawing already has.

    Some floors (e.g. a shared "typical floor" block reused across several
    storeys) keep their grid labels inside a BLOCK definition rather than
    directly in model space. Block-definition coordinates are local to the
    block, so each INSERT of that block is a separate real-world placement
    (e.g. the same block inserted once for a 3F-5F typical plan and again,
    at a different Y, for 6F) — every insert must be walked separately to
    recover the real per-floor position instead of only the first one."""
    axes = defaultdict(lambda: {"letters": [], "numbers": []})
    inserts_by_block = defaultdict(list)
    for row in facts:
        if row.get("type") == "INSERT" and row.get("block") and is_direct_space(row) and row.get("x") is not None:
            inserts_by_block[row["block"]].append(row)

    def add_axis(x, y, label, floor):
        if GRID_LETTER_PATTERN.match(label):
            axes[floor]["letters"].append((y, label))
        elif GRID_NUMBER_PATTERN.match(label):
            axes[floor]["numbers"].append((x, label))

    for row in facts:
        if row.get("layer") != GRID_LAYER or row.get("type") != "TEXT":
            continue
        label = clean(row.get("text")).replace("'", "′")
        x, y = row.get("x"), row.get("y")
        if x is None or y is None:
            continue
        if is_direct_space(row):
            add_axis(float(x), float(y), label, assign_region_by_y(float(y), regions))
        else:
            for insert in inserts_by_block.get(row.get("space"), []):
                world_x, world_y = transform_local_point(float(x), float(y), insert)
                # Bucket by the INSERT's own placement, not the label's transformed
                # y: a label sitting at the block's margin can drift, in world
                # space, past the midpoint into the next floor's bucket even
                # though the whole instance belongs to one floor.
                floor = assign_region_by_y(insert["y"], regions)
                add_axis(world_x, world_y, label, floor)
    return axes


def nearest_axis_pair(point, floor, grid_axes):
    bucket = grid_axes.get(floor)
    if not bucket or not bucket["letters"] or not bucket["numbers"]:
        return None
    number = min(bucket["numbers"], key=lambda item: abs(item[0] - point[0]))[1]
    letter = min(bucket["letters"], key=lambda item: abs(item[0] - point[1]))[1]
    return f"{number}/{letter}"


def refine_location(theme, location, evidence, area=None, grid_axes=None):
    descriptions = [row["description"] for row in evidence]
    pairs, added, removed = evidence_values(descriptions)
    if theme in ("column", "beam"):
        labels = distinct_base_labels([v for pair in pairs for v in pair] + added + removed)
        if labels:
            return f"{'、'.join(labels)} {location}"
    elif theme == "axis_dim":
        axis_labels = [value for value in added if 0 < len(value) <= 4 and AXIS_LABEL_PATTERN.match(value)]
        if axis_labels:
            return f"全平面；{'、'.join(dict.fromkeys(axis_labels))}軸及外圍尺寸鏈"
    elif theme in ("wall_opening", "slab") and grid_axes is not None:
        axis_pairs = []
        for row in evidence:
            for callout in row.get("callouts", []):
                pair = nearest_axis_pair((float(callout["x"]), float(callout["y"])), area, grid_axes)
                if pair and pair not in axis_pairs:
                    axis_pairs.append(pair)
                if len(axis_pairs) >= 6:
                    break
            if len(axis_pairs) >= 6:
                break
        if axis_pairs:
            return f"{'、'.join(axis_pairs)} 軸附近；{location}"
    return location


def floor_sort_key(label):
    if label == "全圖":
        return (-100, 0)
    if label.startswith("B") and label.endswith("F") and label[1:-1].isdigit():
        return (-int(label[1:-1]), 0)
    match = re.match(r"(\d+)F(?:–(\d+)F)?", label)
    if match:
        return (int(match.group(1)), int(match.group(2) or match.group(1)))
    if label.startswith("R") and label.endswith("F") and label[1:-1].isdigit():
        return (100 + int(label[1:-1]), 0)
    if label == "屋突頂層":
        return (199, 0)
    return (999, label)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("candidates_json")
    parser.add_argument("new_facts")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    candidates = json.loads(Path(args.candidates_json).read_text(encoding="utf-8"))
    facts = load(args.new_facts)
    regions = find_regions(facts)
    spacing = region_spacing(regions)
    grid_axes = collect_grid_axes(facts, regions)
    groups = defaultdict(list)
    excluded = []

    for change in candidates.get("changes", []):
        if internal_noise(change) or not meaningful_change(change):
            excluded.append(change)
            continue
        theme = theme_for(change)
        if theme == "revision":
            areas = {"全圖": change.get("callouts", [])}
        elif theme == "general" and (
            change.get("kind") == "geometry"
            or change.get("drawing_area", "").startswith("圖塊 ")
            or any(token in change.get("description", "") for token in ("FS版", "混凝土強度", "350kgf/cm", "連續壁部份", "SD420W", "CNS560"))
        ):
            areas = {"全圖": change.get("callouts", [])}
        else:
            areas = defaultdict(list)
            for callout in change.get("callouts", []):
                point = (float(callout["x"]), float(callout["y"]))
                areas[assign_region(point, regions, spacing)].append(callout)
            if not areas:
                areas[change.get("drawing_area", "模型空間")] = []
        for area, callouts in areas.items():
            evidence = dict(change)
            evidence["callouts"] = callouts
            groups[(area, theme)].append(evidence)

    for area in {key[0] for key in groups}:
        slab_key = (area, "slab")
        beam_key = (area, "beam")
        if beam_key in groups and slab_key in groups:
            pairs, added, removed = evidence_values([row["description"] for row in groups[beam_key]])
            if len(pairs) + len(added) + len(removed) <= 1:
                groups[slab_key].extend(groups.pop(beam_key))
        elevation_key, axis_key = (area, "elevation"), (area, "axis_dim")
        if elevation_key in groups and axis_key in groups:
            groups[axis_key].extend(groups.pop(elevation_key))

    detail_key = ("全圖／其他詳圖", "detail")
    if detail_key in groups:
        for key in list(groups):
            if key[0] == "全圖／其他詳圖" and key[1] not in {"detail", "general", "revision"}:
                groups[detail_key].extend(groups.pop(key))

    changes = []
    for (area, theme), evidence in groups.items():
        category, location = THEME_META[theme]
        location = refine_location(theme, location, evidence, area=area, grid_axes=grid_axes)
        callouts = []
        for row in evidence:
            callouts.extend(row.get("callouts", []))
        callouts = [
            {"x": x, "y": y}
            for x, y in dict.fromkeys((round(float(point["x"]), 6), round(float(point["y"]), 6)) for point in callouts)
        ][:24]
        notes = "需依原生比較快照複核幾何範圍" if any(row.get("kind") == "geometry" for row in evidence) else ""
        changes.append({
            "item": 0,
            "drawing_area": area,
            "location": location,
            "category": category,
            "description": summarize(theme, evidence),
            "notes": notes,
            "kind": "curated",
            "layer": "／".join(dict.fromkeys(row.get("layer", "") for row in evidence if row.get("layer")))[:200],
            "callouts": callouts,
            "evidence_count": len(evidence),
            "_theme": theme,
        })

    changes.sort(key=lambda row: (floor_sort_key(row["drawing_area"]), THEME_ORDER.get(row["_theme"], 99)))
    for index, change in enumerate(changes, 1):
        change["item"] = index
        change.pop("_theme", None)

    warnings = []
    if not regions:
        warnings.append("未辨識到樓層／平面圖標題，需人工補正圖面區域。")
    if any("�" in json.dumps(change, ensure_ascii=False) for change in changes):
        warnings.append("輸出仍含無法解碼字元。")
    if any(len(change.get("description", "")) > 500 or re.search(r"\\[A-Za-z]|[{}]", change.get("description", "")) for change in changes):
        warnings.append("輸出仍含過長或未清理的 CAD 文字格式碼。")
    max_expected = max(12, len(regions) * 5 + 6)
    delivery_ready = bool(changes) and len(changes) <= max_expected and not warnings
    output = {
        "meta": {
            **candidates.get("meta", {}),
            "candidate_count": len(candidates.get("changes", [])),
            "deliverable_count": len(changes),
            "excluded_candidate_count": len(excluded),
            "region_count": len(regions),
            "regions": regions,
            "delivery_ready": delivery_ready,
            "quality_warnings": warnings,
            "review_required": True,
        },
        "changes": changes,
    }
    Path(args.output).write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(output["meta"], ensure_ascii=False))


if __name__ == "__main__":
    main()
