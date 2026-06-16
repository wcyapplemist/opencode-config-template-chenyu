"""Shared fixtures for ppt_builder chart tests."""
import sys
from pathlib import Path

import pytest

_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_SCRIPTS_DIR))


@pytest.fixture
def template_path():
    return str(_SCRIPTS_DIR / "templates" / "template.pptx")


@pytest.fixture
def output_path(tmp_path):
    return str(tmp_path / "test_output.pptx")


@pytest.fixture
def bar_chart_data():
    return {
        "slide_type": "chart_slide",
        "title": "Market Growth",
        "chart_type": "bar",
        "categories": ["2020", "2021", "2022", "2023"],
        "series": [
            {"name": "Revenue", "values": [8.5, 11.2, 14.8, 19.5]},
        ],
        "chart_options": {
            "legend_position": "bottom",
            "show_data_labels": True,
            "y_axis_min": 0,
            "y_axis_max": 25,
        },
        "notes": "Revenue growing steadily.",
    }


@pytest.fixture
def pie_chart_data():
    return {
        "slide_type": "chart_slide",
        "title": "Adoption Rates",
        "chart_type": "pie",
        "categories": ["BIM", "IoT", "Drones", "AI"],
        "series": [
            {"name": "Adoption", "values": [68, 45, 52, 28]},
        ],
        "chart_options": {
            "legend_position": "right",
        },
        "notes": "BIM leads adoption.",
    }


@pytest.fixture
def line_chart_data():
    return {
        "slide_type": "chart_slide",
        "title": "Performance Trends",
        "chart_type": "line_markers",
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
            {"name": "Cost Savings", "values": [5, 8, 12, 16]},
            {"name": "Efficiency", "values": [3, 6, 10, 14]},
        ],
        "chart_options": {
            "legend_position": "bottom",
            "y_axis_min": 0,
            "y_axis_max": 20,
        },
        "notes": "Both metrics trending up.",
    }


@pytest.fixture
def mixed_deck(bar_chart_data, pie_chart_data, line_chart_data):
    return [
        {"slide_type": "title_slide", "title": "Test Deck", "subtitle": "2026"},
        {"slide_type": "content_slide", "title": "Overview", "body": "**Point A** - desc\n**Point B** - desc"},
        bar_chart_data,
        pie_chart_data,
        line_chart_data,
        {"slide_type": "closing_slide", "title": "Thanks", "subtitle": "Q&A"},
    ]
