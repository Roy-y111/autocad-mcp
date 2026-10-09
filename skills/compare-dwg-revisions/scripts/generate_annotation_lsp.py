import argparse
import json
import math
from pathlib import Path


def declutter(marks, radius, gap_factor=1.15):
    """Marks whose real callout points sit closer than 2*radius (the marker
    circles would touch or overlap) are grouped and spread evenly on a small
    ring around their shared centroid, spaced far enough apart that circles
    of this radius no longer collide. Isolated marks are left untouched."""
    min_dist = 2 * radius * gap_factor
    clusters = []
    for label, x, y in marks:
        target = None
        for cluster in clusters:
            if math.hypot(x - cluster["cx"], y - cluster["cy"]) < min_dist:
                target = cluster
                break
        if target is None:
            clusters.append({"points": [(label, x, y)], "cx": x, "cy": y})
            continue
        target["points"].append((label, x, y))
        count = len(target["points"])
        target["cx"] += (x - target["cx"]) / count
        target["cy"] += (y - target["cy"]) / count

    result = []
    for cluster in clusters:
        points = cluster["points"]
        if len(points) == 1:
            result.append(points[0])
            continue
        count = len(points)
        spread_radius = max(radius / max(math.sin(math.pi / count), 0.05), min_dist / 2)
        for index, (label, _, _) in enumerate(points):
            angle = 2 * math.pi * index / count
            result.append((
                label,
                cluster["cx"] + spread_radius * math.cos(angle),
                cluster["cy"] + spread_radius * math.sin(angle),
            ))
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("changes_json")
    parser.add_argument("output_lsp")
    parser.add_argument("--text-height", type=float, default=35.0)
    parser.add_argument("--radius", type=float, default=None, help="Defaults to text-height / 0.65 (same bubble proportion as before).")
    parser.add_argument("--color", type=int, default=1, help="AutoCAD Color Index — 1=red, 2=yellow.")
    args = parser.parse_args()
    data = json.loads(Path(args.changes_json).read_text(encoding="utf-8"))
    text_height = args.text_height
    radius = args.radius if args.radius is not None else text_height / 0.65
    color = args.color
    marks = []
    for change in data.get("changes", []):
        for point in change.get("callouts", []):
            marks.append((str(change["item"]), float(point["x"]), float(point["y"])))
    marks = declutter(marks, radius)
    mark_lines = "\n".join(f'      ("{label}" {x:.8f} {y:.8f})' for label, x, y in marks)
    content = f'''(defun cdr:add-mark (label x y / p)
  (setq p (list x y 0.0))
  (entmakex (list '(0 . "CIRCLE") '(100 . "AcDbEntity") '(8 . "COMPARE-ITEM-NO") '(62 . {color}) '(370 . 50) '(100 . "AcDbCircle") (cons 10 p) (cons 40 {radius:.8f})))
  (entmakex (list '(0 . "TEXT") '(100 . "AcDbEntity") '(8 . "COMPARE-ITEM-NO") '(62 . {color}) '(100 . "AcDbText") (cons 10 p) (cons 11 p) (cons 40 {text_height:.8f}) (cons 1 label) '(50 . 0.0) '(41 . 0.85) '(72 . 1) '(73 . 2) '(7 . "Standard")))
)

(defun c:CDRANNOTATE (/ marks)
  (if (not (tblsearch "LAYER" "COMPARE-ITEM-NO"))
    (entmakex (list '(0 . "LAYER") '(100 . "AcDbSymbolTableRecord") '(100 . "AcDbLayerTableRecord") '(2 . "COMPARE-ITEM-NO") '(70 . 0) '(62 . {color}) '(6 . "Continuous") '(370 . 50))))
  (setq marks '(
{mark_lines}
  ))
  (foreach m marks (cdr:add-mark (car m) (cadr m) (caddr m)))
  (princ (strcat "\\nADDED_CALLOUTS=" (itoa (length marks))))
  (princ)
)
(princ)
'''
    Path(args.output_lsp).write_text(content, encoding="utf-8")
    print(json.dumps({"marks": len(marks), "radius": radius}, ensure_ascii=False))


if __name__ == "__main__":
    main()
