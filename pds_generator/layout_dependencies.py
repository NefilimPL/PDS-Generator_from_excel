from __future__ import annotations

import math


DEFAULT_GRID_SIZE = 5

DEPENDENCY_DIRECTION_ABOVE = "above"
DEPENDENCY_DIRECTION_BELOW = "below"
DEPENDENCY_DIRECTION_LEFT = "left"
DEPENDENCY_DIRECTION_RIGHT = "right"

DEPENDENCY_DIRECTIONS = (
    DEPENDENCY_DIRECTION_ABOVE,
    DEPENDENCY_DIRECTION_BELOW,
    DEPENDENCY_DIRECTION_LEFT,
    DEPENDENCY_DIRECTION_RIGHT,
)

DEPENDENCY_DIRECTION_LABELS = {
    DEPENDENCY_DIRECTION_ABOVE: "nad",
    DEPENDENCY_DIRECTION_BELOW: "pod",
    DEPENDENCY_DIRECTION_LEFT: "z lewej",
    DEPENDENCY_DIRECTION_RIGHT: "z prawej",
}

DEPENDENCY_STATUS_APPLIED = "applied"
DEPENDENCY_STATUS_ALREADY_ALIGNED = "already_aligned"
DEPENDENCY_STATUS_SKIPPED = "skipped"
DEPENDENCY_STATUS_BLOCKED = "blocked"

_DIRECTION_ALIASES = {
    DEPENDENCY_DIRECTION_ABOVE: DEPENDENCY_DIRECTION_ABOVE,
    DEPENDENCY_DIRECTION_BELOW: DEPENDENCY_DIRECTION_BELOW,
    DEPENDENCY_DIRECTION_LEFT: DEPENDENCY_DIRECTION_LEFT,
    DEPENDENCY_DIRECTION_RIGHT: DEPENDENCY_DIRECTION_RIGHT,
    "nad": DEPENDENCY_DIRECTION_ABOVE,
    "pod": DEPENDENCY_DIRECTION_BELOW,
    "z lewej": DEPENDENCY_DIRECTION_LEFT,
    "z prawej": DEPENDENCY_DIRECTION_RIGHT,
}


def _normalize_name_list(values):
    if values is None:
        return []
    if isinstance(values, str):
        values = [values]
    result = []
    seen = set()
    for raw in values:
        name = str(raw or "").strip()
        if not name or name in seen:
            continue
        seen.add(name)
        result.append(name)
    return result


def _normalize_gap_steps(value):
    try:
        steps = int(float(value))
    except (TypeError, ValueError):
        steps = 0
    return max(0, steps)


def normalize_dependency(raw_dependency):
    if not isinstance(raw_dependency, dict):
        return None

    anchor_names = _normalize_name_list(
        raw_dependency.get("anchor_names")
        or raw_dependency.get("anchors")
        or raw_dependency.get("anchor")
    )
    mover_names = _normalize_name_list(
        raw_dependency.get("mover_names")
        or raw_dependency.get("movers")
        or raw_dependency.get("mover")
    )

    direction = str(raw_dependency.get("direction", DEPENDENCY_DIRECTION_BELOW) or "").strip().lower()
    direction = _DIRECTION_ALIASES.get(direction, DEPENDENCY_DIRECTION_BELOW)

    dependency = {
        "anchor_names": anchor_names,
        "mover_names": mover_names,
        "direction": direction,
        "gap_steps": _normalize_gap_steps(
            raw_dependency.get(
                "gap_steps",
                raw_dependency.get("gap", raw_dependency.get("distance", 0)),
            )
        ),
    }
    if not dependency["anchor_names"] or not dependency["mover_names"]:
        return None
    return dependency


def normalize_dependencies(raw_dependencies):
    normalized = []
    for raw_dependency in raw_dependencies or []:
        dependency = normalize_dependency(raw_dependency)
        if dependency is not None:
            normalized.append(dependency)
    return normalized


def _bounding_box(elements, names):
    left = top = right = bottom = None
    for name in names:
        element = elements.get(name)
        if element is None:
            continue
        x1 = float(getattr(element, "x", 0))
        y1 = float(getattr(element, "y", 0))
        x2 = x1 + float(getattr(element, "width", 0))
        y2 = y1 + float(getattr(element, "height", 0))
        if left is None:
            left, top, right, bottom = x1, y1, x2, y2
            continue
        left = min(left, x1)
        top = min(top, y1)
        right = max(right, x2)
        bottom = max(bottom, y2)
    if left is None:
        return None
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
    }


def bounding_box_for_names(elements, names):
    return _bounding_box(elements, names)


def _element_box(element):
    if element is None:
        return None
    x1 = float(getattr(element, "x", 0))
    y1 = float(getattr(element, "y", 0))
    return {
        "left": x1,
        "top": y1,
        "right": x1 + float(getattr(element, "width", 0)),
        "bottom": y1 + float(getattr(element, "height", 0)),
    }


def _bounding_box_from_boxes(boxes_by_name, names):
    left = top = right = bottom = None
    for name in names:
        box = boxes_by_name.get(name)
        if box is None:
            continue
        if left is None:
            left = box["left"]
            top = box["top"]
            right = box["right"]
            bottom = box["bottom"]
            continue
        left = min(left, box["left"])
        top = min(top, box["top"])
        right = max(right, box["right"])
        bottom = max(bottom, box["bottom"])
    if left is None:
        return None
    return {
        "left": left,
        "top": top,
        "right": right,
        "bottom": bottom,
    }


def _shift_box(box, dx, dy):
    return {
        "left": box["left"] + dx,
        "top": box["top"] + dy,
        "right": box["right"] + dx,
        "bottom": box["bottom"] + dy,
    }


def _boxes_overlap(box_a, box_b):
    if box_a is None or box_b is None:
        return False
    return (
        box_a["left"] < box_b["right"]
        and box_a["right"] > box_b["left"]
        and box_a["top"] < box_b["bottom"]
        and box_a["bottom"] > box_b["top"]
    )


def _find_collision_names(boxes_by_name, mover_names, hidden_names, dx, dy):
    mover_set = set(mover_names)
    proposed_boxes = {}
    for name in mover_names:
        current_box = boxes_by_name.get(name)
        if current_box is None:
            continue
        proposed_boxes[name] = _shift_box(current_box, dx, dy)
    collisions = []
    for other_name, other_box in boxes_by_name.items():
        if other_name in hidden_names or other_name in mover_set:
            continue
        for mover_name, mover_box in proposed_boxes.items():
            if _boxes_overlap(mover_box, other_box):
                collisions.append(other_name)
                break
    return collisions


def _snap_shift(value, step, direction):
    if not step:
        return value
    if direction in (DEPENDENCY_DIRECTION_BELOW, DEPENDENCY_DIRECTION_RIGHT):
        return math.ceil(value / step) * step
    if direction in (DEPENDENCY_DIRECTION_ABOVE, DEPENDENCY_DIRECTION_LEFT):
        return math.floor(value / step) * step
    return value


def _resolve_first_valid_shift(
    boxes_by_name,
    mover_names,
    hidden_names,
    mover_box,
    direction,
    dx,
    dy,
    gap,
    step,
):
    current_dx = dx
    current_dy = dy
    collision_names = _find_collision_names(
        boxes_by_name,
        mover_names,
        hidden_names,
        current_dx,
        current_dy,
    )
    if not collision_names:
        return current_dx, current_dy, collision_names

    for _ in range(max(1, len(boxes_by_name) + 5)):
        next_dx = current_dx
        next_dy = current_dy
        if direction == DEPENDENCY_DIRECTION_BELOW:
            for name in collision_names:
                blocker = boxes_by_name.get(name)
                if blocker is None:
                    continue
                next_dy = max(
                    next_dy,
                    blocker["bottom"] + gap - mover_box["top"],
                )
            next_dy = _snap_shift(next_dy, step, direction)
        elif direction == DEPENDENCY_DIRECTION_ABOVE:
            for name in collision_names:
                blocker = boxes_by_name.get(name)
                if blocker is None:
                    continue
                next_dy = min(
                    next_dy,
                    blocker["top"] - gap - mover_box["bottom"],
                )
            next_dy = _snap_shift(next_dy, step, direction)
        elif direction == DEPENDENCY_DIRECTION_RIGHT:
            for name in collision_names:
                blocker = boxes_by_name.get(name)
                if blocker is None:
                    continue
                next_dx = max(
                    next_dx,
                    blocker["right"] + gap - mover_box["left"],
                )
            next_dx = _snap_shift(next_dx, step, direction)
        elif direction == DEPENDENCY_DIRECTION_LEFT:
            for name in collision_names:
                blocker = boxes_by_name.get(name)
                if blocker is None:
                    continue
                next_dx = min(
                    next_dx,
                    blocker["left"] - gap - mover_box["right"],
                )
            next_dx = _snap_shift(next_dx, step, direction)

        if next_dx == current_dx and next_dy == current_dy:
            break

        current_dx = next_dx
        current_dy = next_dy
        collision_names = _find_collision_names(
            boxes_by_name,
            mover_names,
            hidden_names,
            current_dx,
            current_dy,
        )
        if not collision_names:
            break

    return current_dx, current_dy, collision_names


def apply_layout_dependencies(
    elements,
    hidden_names=None,
    dependencies=None,
    grid_step=DEFAULT_GRID_SIZE,
    return_statuses=False,
):
    if not elements or not dependencies:
        return (elements, []) if return_statuses else elements

    hidden = set(hidden_names or [])
    step = float(grid_step or DEFAULT_GRID_SIZE)
    statuses = []
    boxes_by_name = {
        name: _element_box(element) for name, element in (elements or {}).items()
    }

    for index, dependency in enumerate(normalize_dependencies(dependencies)):
        anchor_names = [
            name
            for name in dependency["anchor_names"]
            if name in elements and name not in hidden
        ]
        mover_names = [
            name
            for name in dependency["mover_names"]
            if name in elements and name not in hidden
        ]
        status = {
            "index": index,
            "dependency": dict(dependency),
            "anchor_names": list(anchor_names),
            "mover_names": list(mover_names),
            "direction": dependency["direction"],
            "gap_steps": dependency["gap_steps"],
            "requested_dx": 0.0,
            "requested_dy": 0.0,
            "dx": 0.0,
            "dy": 0.0,
            "anchor_box": None,
            "mover_box_before": None,
            "mover_box_after": None,
            "blocked_by": [],
            "adjusted": False,
            "status": DEPENDENCY_STATUS_SKIPPED,
        }
        if not anchor_names or not mover_names:
            statuses.append(status)
            continue
        if set(anchor_names).intersection(mover_names):
            statuses.append(status)
            continue

        anchor_box = _bounding_box_from_boxes(boxes_by_name, anchor_names)
        mover_box = _bounding_box_from_boxes(boxes_by_name, mover_names)
        if anchor_box is None or mover_box is None:
            statuses.append(status)
            continue

        gap = float(dependency["gap_steps"]) * step
        dx = 0.0
        dy = 0.0
        direction = dependency["direction"]
        if direction == DEPENDENCY_DIRECTION_BELOW:
            dy = anchor_box["bottom"] + gap - mover_box["top"]
        elif direction == DEPENDENCY_DIRECTION_ABOVE:
            dy = anchor_box["top"] - gap - mover_box["bottom"]
        elif direction == DEPENDENCY_DIRECTION_RIGHT:
            dx = anchor_box["right"] + gap - mover_box["left"]
        elif direction == DEPENDENCY_DIRECTION_LEFT:
            dx = anchor_box["left"] - gap - mover_box["right"]

        status["requested_dx"] = dx
        status["requested_dy"] = dy
        dx, dy, collision_names = _resolve_first_valid_shift(
            boxes_by_name,
            mover_names,
            hidden,
            mover_box,
            direction,
            dx,
            dy,
            gap,
            step,
        )
        status["dx"] = dx
        status["dy"] = dy
        status["anchor_box"] = anchor_box
        status["mover_box_before"] = mover_box
        status["mover_box_after"] = _shift_box(mover_box, dx, dy)
        status["adjusted"] = (
            dx != status["requested_dx"] or dy != status["requested_dy"]
        )

        if not dx and not dy:
            status["status"] = DEPENDENCY_STATUS_ALREADY_ALIGNED
            statuses.append(status)
            continue

        if collision_names:
            status["status"] = DEPENDENCY_STATUS_BLOCKED
            status["blocked_by"] = collision_names
            statuses.append(status)
            continue

        for name in mover_names:
            element = elements.get(name)
            if element is None:
                continue
            element.x = float(getattr(element, "x", 0)) + dx
            element.y = float(getattr(element, "y", 0)) + dy
            current_box = boxes_by_name.get(name)
            if current_box is not None:
                boxes_by_name[name] = _shift_box(current_box, dx, dy)
            else:
                boxes_by_name[name] = _element_box(element)
        status["status"] = DEPENDENCY_STATUS_APPLIED
        statuses.append(status)

    return (elements, statuses) if return_statuses else elements
