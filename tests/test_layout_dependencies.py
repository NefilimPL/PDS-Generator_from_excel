from types import SimpleNamespace

from pds_generator.layout_dependencies import (
    DEPENDENCY_STATUS_APPLIED,
    apply_layout_dependencies,
)


def _element(x, y, width=40, height=20):
    return SimpleNamespace(x=x, y=y, width=width, height=height)


def test_dependency_uses_visible_anchor_range_when_hidden_elements_shrink_it():
    elements = {
        "A": _element(0, 0),
        "B": _element(0, 30),
        "C": _element(0, 60),
    }

    apply_layout_dependencies(
        elements,
        hidden_names={"B"},
        dependencies=[
            {
                "anchor_names": ["A", "B"],
                "mover_names": ["C"],
                "direction": "below",
                "gap_steps": 1,
            }
        ],
        grid_step=5,
    )

    assert elements["C"].y == 25


def test_dependency_moves_multiple_blocks_as_one_range():
    elements = {
        "Header": _element(0, 0),
        "Block1": _element(0, 60),
        "Block2": _element(30, 100),
    }

    apply_layout_dependencies(
        elements,
        dependencies=[
            {
                "anchor_names": ["Header"],
                "mover_names": ["Block1", "Block2"],
                "direction": "below",
                "gap_steps": 1,
            }
        ],
        grid_step=5,
    )

    assert elements["Block1"].y == 25
    assert elements["Block2"].y == 65
    assert elements["Block2"].x == 30


def test_dependencies_are_applied_in_declared_order():
    elements = {
        "A": _element(0, 0),
        "B": _element(0, 70),
        "C": _element(0, 140),
    }

    apply_layout_dependencies(
        elements,
        dependencies=[
            {
                "anchor_names": ["A"],
                "mover_names": ["B"],
                "direction": "below",
                "gap_steps": 1,
            },
            {
                "anchor_names": ["B"],
                "mover_names": ["C"],
                "direction": "below",
                "gap_steps": 1,
            },
        ],
        grid_step=5,
    )

    assert elements["B"].y == 25
    assert elements["C"].y == 50


def test_dependency_moves_to_first_free_position_when_exact_target_collides():
    elements = {
        "Header": _element(0, 0),
        "Middle": _element(0, 30),
        "Footer": _element(0, 90),
    }

    _, statuses = apply_layout_dependencies(
        elements,
        dependencies=[
            {
                "anchor_names": ["Header"],
                "mover_names": ["Footer"],
                "direction": "below",
                "gap_steps": 1,
            }
        ],
        grid_step=5,
        return_statuses=True,
    )

    assert elements["Footer"].y == 55
    assert statuses[0]["status"] == DEPENDENCY_STATUS_APPLIED
    assert statuses[0]["adjusted"] is True


def test_hidden_blocks_do_not_prevent_dosuwanie_in_simulation():
    elements = {
        "Header": _element(0, 0),
        "Optional": _element(0, 30),
        "Footer": _element(0, 90),
    }

    _, statuses = apply_layout_dependencies(
        elements,
        hidden_names={"Optional"},
        dependencies=[
            {
                "anchor_names": ["Header"],
                "mover_names": ["Footer"],
                "direction": "below",
                "gap_steps": 1,
            }
        ],
        grid_step=5,
        return_statuses=True,
    )

    assert elements["Footer"].y == 25
    assert statuses[0]["status"] == DEPENDENCY_STATUS_APPLIED
    assert statuses[0]["adjusted"] is False


def test_dependency_keeps_configured_gap_from_nearest_visible_blocker():
    elements = {
        "Header": _element(0, 0),
        "MiddleA": _element(0, 30),
        "MiddleB": _element(0, 60),
        "Footer": _element(0, 140),
    }

    _, statuses = apply_layout_dependencies(
        elements,
        dependencies=[
            {
                "anchor_names": ["Header"],
                "mover_names": ["Footer"],
                "direction": "below",
                "gap_steps": 2,
            }
        ],
        grid_step=5,
        return_statuses=True,
    )

    assert elements["Footer"].y == 90
    assert statuses[0]["status"] == DEPENDENCY_STATUS_APPLIED
    assert statuses[0]["adjusted"] is True
