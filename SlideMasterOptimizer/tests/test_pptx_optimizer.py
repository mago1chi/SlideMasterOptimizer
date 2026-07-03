from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from slidemasteroptimizer.core.pptx_optimizer import analyze_pptx, optimize_pptx


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("", CT_NS)
ET.register_namespace("", REL_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)


def test_analyze_finds_unused_master(tmp_path: Path) -> None:
    pptx_path = tmp_path / "sample.pptx"
    _write_sample_pptx(pptx_path, include_unused_master=True)

    result = analyze_pptx(pptx_path)

    assert result.slide_count == 1
    assert result.total_masters == 2
    assert result.unused_master_count == 1
    assert result.unused_masters[0].part_name == "ppt/slideMasters/slideMaster2.xml"
    assert result.warnings == ()


def test_optimize_removes_unused_master_and_layout(tmp_path: Path) -> None:
    pptx_path = tmp_path / "sample.pptx"
    output_path = tmp_path / "sample_optimized.pptx"
    _write_sample_pptx(pptx_path, include_unused_master=True)

    result = optimize_pptx(pptx_path, output_path)

    assert result.removed_master_count == 1
    assert result.removed_layout_count == 1
    with ZipFile(output_path) as pptx:
        names = set(pptx.namelist())
        assert "ppt/slideMasters/slideMaster2.xml" not in names
        assert "ppt/slideLayouts/slideLayout2.xml" not in names
        assert "ppt/slideMasters/slideMaster1.xml" in names
        assert "ppt/slideLayouts/slideLayout1.xml" in names

    reanalysis = analyze_pptx(output_path)
    assert reanalysis.unused_master_count == 0
    assert reanalysis.total_masters == 1


def test_optimize_removes_unused_layout_from_used_master(tmp_path: Path) -> None:
    pptx_path = tmp_path / "sample.pptx"
    output_path = tmp_path / "sample_optimized.pptx"
    _write_sample_pptx(
        pptx_path,
        include_unused_master=False,
        include_unused_layout_in_used_master=True,
    )

    analysis = analyze_pptx(pptx_path)
    assert analysis.unused_master_count == 0
    assert analysis.unused_layout_count == 1
    assert analysis.unused_layouts[0].part_name == "ppt/slideLayouts/slideLayout2.xml"

    result = optimize_pptx(pptx_path, output_path)

    assert result.removed_master_count == 0
    assert result.removed_layout_count == 1
    with ZipFile(output_path) as pptx:
        names = set(pptx.namelist())
        assert "ppt/slideMasters/slideMaster1.xml" in names
        assert "ppt/slideLayouts/slideLayout1.xml" in names
        assert "ppt/slideLayouts/slideLayout2.xml" not in names
        rels = pptx.read("ppt/slideMasters/_rels/slideMaster1.xml.rels").decode()
        assert "slideLayout2.xml" not in rels

    reanalysis = analyze_pptx(output_path)
    assert reanalysis.unused_master_count == 0
    assert reanalysis.unused_layout_count == 0


def test_no_unused_master_copies_file(tmp_path: Path) -> None:
    pptx_path = tmp_path / "sample.pptx"
    output_path = tmp_path / "sample_optimized.pptx"
    _write_sample_pptx(pptx_path, include_unused_master=False)

    result = optimize_pptx(pptx_path, output_path)

    assert result.removed_master_count == 0
    assert result.removed_layout_count == 0
    assert output_path.exists()
    assert analyze_pptx(output_path).unused_master_count == 0


def test_broken_relationship_reports_warning(tmp_path: Path) -> None:
    pptx_path = tmp_path / "broken.pptx"
    _write_sample_pptx(pptx_path, include_unused_master=False, broken_layout=True)

    result = analyze_pptx(pptx_path)

    assert result.warnings
    assert "Missing slide layout part" in result.warnings[0]


def _write_sample_pptx(
    path: Path,
    *,
    include_unused_master: bool,
    include_unused_layout_in_used_master: bool = False,
    broken_layout: bool = False,
) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as pptx:
        pptx.writestr(
            "[Content_Types].xml",
            _content_types(include_unused_master, include_unused_layout_in_used_master),
        )
        pptx.writestr("ppt/presentation.xml", _presentation(include_unused_master))
        pptx.writestr("ppt/_rels/presentation.xml.rels", _presentation_rels(include_unused_master))
        pptx.writestr("ppt/slides/slide1.xml", "<p:sld xmlns:p='%s'/>" % P_NS)
        pptx.writestr("ppt/slides/_rels/slide1.xml.rels", _slide_rels(broken_layout))
        pptx.writestr("ppt/slideLayouts/slideLayout1.xml", "<p:sldLayout xmlns:p='%s'/>" % P_NS)
        pptx.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", _layout_rels(1))
        pptx.writestr("ppt/slideMasters/slideMaster1.xml", _master_xml([1, 2] if include_unused_layout_in_used_master else [1]))
        pptx.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", _master_rels([1, 2] if include_unused_layout_in_used_master else [1]))
        if include_unused_master or include_unused_layout_in_used_master:
            pptx.writestr("ppt/slideLayouts/slideLayout2.xml", "<p:sldLayout xmlns:p='%s'/>" % P_NS)
            pptx.writestr("ppt/slideLayouts/_rels/slideLayout2.xml.rels", _layout_rels(2))
        if include_unused_master:
            pptx.writestr("ppt/slideMasters/slideMaster2.xml", "<p:sldMaster xmlns:p='%s'/>" % P_NS)
            pptx.writestr("ppt/slideMasters/_rels/slideMaster2.xml.rels", _master_rels([2]))


def _content_types(include_unused_master: bool, include_unused_layout_in_used_master: bool = False) -> str:
    overrides = [
        "/ppt/presentation.xml",
        "/ppt/slides/slide1.xml",
        "/ppt/slideLayouts/slideLayout1.xml",
        "/ppt/slideMasters/slideMaster1.xml",
    ]
    if include_unused_master or include_unused_layout_in_used_master:
        overrides.append("/ppt/slideLayouts/slideLayout2.xml")
    if include_unused_master:
        overrides.append("/ppt/slideMasters/slideMaster2.xml")
    body = "".join(
        f'<Override PartName="{part}" ContentType="application/xml"/>'
        for part in overrides
    )
    return f'<Types xmlns="{CT_NS}"><Default Extension="rels" ContentType="application/xml"/>{body}</Types>'


def _presentation(include_unused_master: bool) -> str:
    master_2 = '<p:sldMasterId id="2147483649" r:id="rIdMaster2"/>' if include_unused_master else ""
    return (
        f'<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}">'
        '<p:sldMasterIdLst>'
        '<p:sldMasterId id="2147483648" r:id="rIdMaster1"/>'
        f"{master_2}"
        '</p:sldMasterIdLst>'
        '<p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>'
        '</p:presentation>'
    )


def _presentation_rels(include_unused_master: bool) -> str:
    rels = [
        '<Relationship Id="rIdSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>',
        '<Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>',
    ]
    if include_unused_master:
        rels.append(
            '<Relationship Id="rIdMaster2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster2.xml"/>'
        )
    return f'<Relationships xmlns="{REL_NS}">{"".join(rels)}</Relationships>'


def _slide_rels(broken_layout: bool) -> str:
    target = "../slideLayouts/missing.xml" if broken_layout else "../slideLayouts/slideLayout1.xml"
    return (
        f'<Relationships xmlns="{REL_NS}">'
        f'<Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="{target}"/>'
        '</Relationships>'
    )


def _layout_rels(index: int) -> str:
    return (
        f'<Relationships xmlns="{REL_NS}">'
        f'<Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster{index}.xml"/>'
        '</Relationships>'
    )


def _master_xml(layout_indexes: list[int]) -> str:
    layouts = "".join(
        f'<p:sldLayoutId id="{2147483648 + index}" r:id="rIdLayout{index}"/>'
        for index in layout_indexes
    )
    return (
        f'<p:sldMaster xmlns:p="{P_NS}" xmlns:r="{R_NS}">'
        f"<p:sldLayoutIdLst>{layouts}</p:sldLayoutIdLst>"
        "</p:sldMaster>"
    )


def _master_rels(indexes: list[int]) -> str:
    rels = "".join(
        f'<Relationship Id="rIdLayout{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout{index}.xml"/>'
        for index in indexes
    )
    return (
        f'<Relationships xmlns="{REL_NS}">'
        f"{rels}"
        '</Relationships>'
    )
