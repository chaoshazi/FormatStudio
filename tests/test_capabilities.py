from app.capabilities import capability_for, full_matrix


def test_supported_pairs():
    assert capability_for("txt", "pdf") is not None
    assert capability_for("pdf", "docx")["modes"] == ["editable", "layout"]
    assert capability_for("pdf", "pptx")["modes"] == ["editable", "layout"]
    assert capability_for("csv", "xlsx") is not None
    assert capability_for("docx", "xlsx") is not None
    assert capability_for("pdf", "xlsx")["status"] == "beta"


def test_disabled_pairs():
    assert capability_for("docx", "pptx") is None
    assert capability_for("pptx", "docx") is None
    assert capability_for("pdf", "csv") is None
    assert capability_for("wps", "xlsx") is None


def test_full_matrix_shape():
    rows = full_matrix()
    assert len(rows) == 8 * 7
    assert any(row["status"] == "disabled" for row in rows)
    assert any(row["status"] == "supported" for row in rows)