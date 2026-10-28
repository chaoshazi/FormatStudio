from __future__ import annotations

import csv
from collections import Counter
from pathlib import Path


def test_builtin_template_gallery(client):
    templates = client.get("/api/v1/templates").json()
    counts = Counter(item["format"] for item in templates if item.get("builtin"))
    assert counts == {"docx": 10, "xlsx": 10, "pptx": 10}
    for item in templates:
        if not item.get("builtin"):
            continue
        assert item["fields"]
        assert Path(item["path"]).is_file()


def test_builtin_templates_cannot_be_deleted(client):
    template = next(item for item in client.get("/api/v1/templates").json() if item.get("builtin"))
    response = client.delete(f"/api/v1/templates/{template['id']}")
    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "BUILTIN_TEMPLATE"


def test_style_gallery_has_ten_per_format(client):
    styles = client.get("/api/v1/styles").json()
    counts = Counter(fmt for style in styles for fmt in style["applies_to"])
    assert counts["docx"] >= 10
    assert counts["xlsx"] >= 10
    assert counts["pptx"] >= 10


def test_generate_from_builtin_template(client, work_dir: Path):
    template = next(
        item
        for item in client.get("/api/v1/templates").json()
        if item.get("builtin") and item["format"] == "docx"
    )
    data_path = work_dir / "data.csv"
    fields = [field["name"] for field in template["fields"]]
    with data_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(fields)
        writer.writerow([f"示例{v}" for v in range(len(fields))])
    with data_path.open("rb") as handle:
        response = client.post(
            f"/api/v1/templates/{template['id']}/generate",
            files=[("data_file", (data_path.name, handle, "text/csv"))],
        )
    assert response.status_code == 200, response.text
    payload = response.json()
    assert payload["generated"] == 1
    assert payload["files"][0]["name"].endswith(".docx")