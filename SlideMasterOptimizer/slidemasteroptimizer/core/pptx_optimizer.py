from __future__ import annotations

from dataclasses import dataclass, field
import posixpath
from pathlib import Path, PurePosixPath
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NS = {"p": P_NS, "r": R_NS}
ET.register_namespace("", CT_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)


SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
SLIDE_LAYOUT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
SLIDE_MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"


class PptxOptimizationError(RuntimeError):
    """Raised when a pptx cannot be optimized safely."""


@dataclass(frozen=True)
class MasterCandidate:
    part_name: str
    relationship_id: str | None = None


@dataclass(frozen=True)
class LayoutCandidate:
    part_name: str
    master_part_name: str
    relationship_id: str | None = None


@dataclass(frozen=True)
class AnalysisResult:
    input_path: Path
    slide_count: int
    total_masters: int
    total_layouts: int
    used_masters: tuple[str, ...]
    used_layouts: tuple[str, ...]
    unused_masters: tuple[MasterCandidate, ...]
    unused_layouts: tuple[LayoutCandidate, ...]
    warnings: tuple[str, ...] = field(default_factory=tuple)

    @property
    def used_master_count(self) -> int:
        return len(self.used_masters)

    @property
    def unused_master_count(self) -> int:
        return len(self.unused_masters)

    @property
    def used_layout_count(self) -> int:
        return len(self.used_layouts)

    @property
    def unused_layout_count(self) -> int:
        return len(self.unused_layouts)

    @property
    def removal_candidate_count(self) -> int:
        return self.unused_master_count + self.unused_layout_count


@dataclass(frozen=True)
class OptimizeResult:
    input_path: Path
    output_path: Path
    removed_master_count: int
    removed_layout_count: int
    warnings: tuple[str, ...] = field(default_factory=tuple)


@dataclass(frozen=True)
class Relationship:
    rel_id: str
    rel_type: str
    target: str
    target_mode: str | None = None


def analyze_pptx(path: Path | str) -> AnalysisResult:
    input_path = Path(path)
    _validate_input_path(input_path)

    warnings: list[str] = []
    with ZipFile(input_path) as pptx:
        names = set(pptx.namelist())
        _require_part(names, "ppt/presentation.xml")
        _require_part(names, "ppt/_rels/presentation.xml.rels")

        presentation_rels = _read_relationships(pptx, "ppt/_rels/presentation.xml.rels")
        slide_rels = [rel for rel in presentation_rels if rel.rel_type == SLIDE_REL_TYPE]
        master_rels = [rel for rel in presentation_rels if rel.rel_type == SLIDE_MASTER_REL_TYPE]

        all_masters_by_part: dict[str, str] = {}
        for rel in master_rels:
            part = _resolve_target("ppt/presentation.xml", rel.target)
            if part in names:
                all_masters_by_part[part] = rel.rel_id
            else:
                warnings.append(f"Missing slide master part referenced by {rel.rel_id}: {part}")

        used_masters: set[str] = set()
        used_layouts: set[str] = set()
        for slide_rel in slide_rels:
            slide_part = _resolve_target("ppt/presentation.xml", slide_rel.target)
            if slide_part not in names:
                warnings.append(f"Missing slide part referenced by {slide_rel.rel_id}: {slide_part}")
                continue

            slide_rels_path = _rels_path_for_part(slide_part)
            if slide_rels_path not in names:
                warnings.append(f"Missing relationship file for slide: {slide_part}")
                continue

            layout_rel = _first_relationship_of_type(
                _read_relationships(pptx, slide_rels_path), SLIDE_LAYOUT_REL_TYPE
            )
            if layout_rel is None:
                warnings.append(f"Slide has no slide layout relationship: {slide_part}")
                continue

            layout_part = _resolve_target(slide_part, layout_rel.target)
            if layout_part not in names:
                warnings.append(f"Missing slide layout part referenced by {slide_part}: {layout_part}")
                continue
            used_layouts.add(layout_part)

            layout_rels_path = _rels_path_for_part(layout_part)
            if layout_rels_path not in names:
                warnings.append(f"Missing relationship file for slide layout: {layout_part}")
                continue

            master_rel = _first_relationship_of_type(
                _read_relationships(pptx, layout_rels_path), SLIDE_MASTER_REL_TYPE
            )
            if master_rel is None:
                warnings.append(f"Slide layout has no slide master relationship: {layout_part}")
                continue

            master_part = _resolve_target(layout_part, master_rel.target)
            if master_part not in names:
                warnings.append(f"Missing slide master part referenced by {layout_part}: {master_part}")
                continue
            used_masters.add(master_part)

        all_layouts_by_master = _layout_relationships_by_master(
            pptx, names, set(all_masters_by_part)
        )
        unused = [
            MasterCandidate(part_name=part, relationship_id=rel_id)
            for part, rel_id in sorted(all_masters_by_part.items())
            if part not in used_masters
        ]
        unused_master_parts = {candidate.part_name for candidate in unused}
        unused_layouts = [
            LayoutCandidate(
                part_name=layout_part,
                master_part_name=master_part,
                relationship_id=rel.rel_id,
            )
            for master_part, relationships in sorted(all_layouts_by_master.items())
            if master_part not in unused_master_parts
            for rel, layout_part in relationships
            if layout_part not in used_layouts
        ]

    return AnalysisResult(
        input_path=input_path,
        slide_count=len(slide_rels),
        total_masters=len(all_masters_by_part),
        total_layouts=sum(len(relationships) for relationships in all_layouts_by_master.values()),
        used_masters=tuple(sorted(used_masters)),
        used_layouts=tuple(sorted(used_layouts)),
        unused_masters=tuple(unused),
        unused_layouts=tuple(unused_layouts),
        warnings=tuple(warnings),
    )


def optimize_pptx(input_path: Path | str, output_path: Path | str) -> OptimizeResult:
    source = Path(input_path)
    destination = Path(output_path)
    analysis = analyze_pptx(source)

    if analysis.warnings:
        raise PptxOptimizationError(
            "The presentation has unresolved references. No output was written."
        )

    destination.parent.mkdir(parents=True, exist_ok=True)
    unused_master_parts = {candidate.part_name for candidate in analysis.unused_masters}
    unused_layout_parts = {candidate.part_name for candidate in analysis.unused_layouts}
    if not unused_master_parts and not unused_layout_parts:
        _copy_pptx(source, destination)
        return OptimizeResult(
            input_path=source,
            output_path=destination,
            removed_master_count=0,
            removed_layout_count=0,
        )

    with ZipFile(source) as src:
        names = set(src.namelist())
        presentation_xml = _remove_master_ids(
            src.read("ppt/presentation.xml"), analysis.unused_masters
        )
        presentation_rels_xml = _remove_relationships_by_targets(
            src.read("ppt/_rels/presentation.xml.rels"),
            "ppt/presentation.xml",
            unused_master_parts,
        )

        layout_parts = _layout_parts_for_masters(src, names, unused_master_parts) | unused_layout_parts
        initial_delete_parts = set(unused_master_parts) | layout_parts
        initial_delete_parts.update(_rels_path_for_part(part) for part in list(initial_delete_parts))
        delete_parts = _collect_only_deleted_references(src, names, initial_delete_parts)
        content_types_xml = _remove_content_type_overrides(
            src.read("[Content_Types].xml"), delete_parts
        )

        with ZipFile(destination, "w", ZIP_DEFLATED) as dst:
            written: set[str] = set()
            for info in src.infolist():
                name = info.filename
                if name in delete_parts:
                    continue
                if name == "ppt/presentation.xml":
                    _write_once(dst, info, presentation_xml, written)
                    continue
                if name == "ppt/_rels/presentation.xml.rels":
                    _write_once(dst, info, presentation_rels_xml, written)
                    continue
                if name == "[Content_Types].xml":
                    _write_once(dst, info, content_types_xml, written)
                    continue
                if name in analysis.used_masters:
                    _write_once(
                        dst,
                        info,
                        _remove_layout_ids(src.read(name), name, analysis.unused_layouts),
                        written,
                    )
                    continue
                if name.endswith(".rels"):
                    source_part = _part_for_rels_path(name)
                    if source_part in analysis.used_masters:
                        _write_once(
                            dst,
                            info,
                            _remove_layout_relationships(
                                src.read(name), source_part, analysis.unused_layouts
                            ),
                            written,
                        )
                        continue
                _write_once(dst, info, src.read(name), written)

    integrity_warnings = _validate_package_integrity(destination)
    if integrity_warnings:
        raise PptxOptimizationError(
            "The optimized presentation has unresolved package references: "
            + "; ".join(integrity_warnings[:5])
        )

    return OptimizeResult(
        input_path=source,
        output_path=destination,
        removed_master_count=len(unused_master_parts),
        removed_layout_count=len(layout_parts),
    )


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_optimized{input_path.suffix}")


def _validate_input_path(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(path)
    if not path.is_file():
        raise PptxOptimizationError("Input must be a .pptx file.")
    if path.suffix.lower() != ".pptx":
        raise PptxOptimizationError("Only .pptx files are supported.")


def _require_part(names: set[str], part_name: str) -> None:
    if part_name not in names:
        raise PptxOptimizationError(f"Required pptx part is missing: {part_name}")


def _read_relationships(pptx: ZipFile, rels_path: str) -> list[Relationship]:
    root = ET.fromstring(pptx.read(rels_path))
    relationships: list[Relationship] = []
    for elem in root.findall(f"{{{REL_NS}}}Relationship"):
        relationships.append(
            Relationship(
                rel_id=elem.attrib.get("Id", ""),
                rel_type=elem.attrib.get("Type", ""),
                target=elem.attrib.get("Target", ""),
                target_mode=elem.attrib.get("TargetMode"),
            )
        )
    return relationships


def _first_relationship_of_type(
    relationships: Iterable[Relationship], rel_type: str
) -> Relationship | None:
    return next((rel for rel in relationships if rel.rel_type == rel_type), None)


def _resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    base = PurePosixPath(source_part).parent
    return posixpath.normpath(str(PurePosixPath(base, target)))


def _rels_path_for_part(part_name: str) -> str:
    part = PurePosixPath(part_name)
    return str(part.parent / "_rels" / f"{part.name}.rels")


def _remove_master_ids(data: bytes, unused_masters: Iterable[MasterCandidate]) -> bytes:
    rel_ids = {candidate.relationship_id for candidate in unused_masters if candidate.relationship_id}
    root = etree.fromstring(data)
    master_list = root.find("p:sldMasterIdLst", NS)
    if master_list is not None:
        for elem in list(master_list):
            if elem.attrib.get(f"{{{R_NS}}}id") in rel_ids:
                master_list.remove(elem)
    return etree.tostring(root, encoding="utf-8", xml_declaration=True)


def _remove_relationships_by_targets(data: bytes, source_part: str, targets: set[str]) -> bytes:
    root = etree.fromstring(data)
    for rel in list(root):
        target = rel.attrib.get("Target", "")
        if rel.attrib.get("TargetMode") == "External":
            continue
        if _resolve_target(source_part, target) in targets:
            root.remove(rel)
    return etree.tostring(root, encoding="utf-8", xml_declaration=True)


def _layout_parts_for_masters(
    pptx: ZipFile, names: set[str], master_parts: set[str]
) -> set[str]:
    layout_parts: set[str] = set()
    for master_part in master_parts:
        rels_path = _rels_path_for_part(master_part)
        if rels_path not in names:
            continue
        for rel in _read_relationships(pptx, rels_path):
            if rel.rel_type == SLIDE_LAYOUT_REL_TYPE and rel.target_mode != "External":
                layout_part = _resolve_target(master_part, rel.target)
                if layout_part in names:
                    layout_parts.add(layout_part)
    return layout_parts


def _layout_relationships_by_master(
    pptx: ZipFile, names: set[str], master_parts: set[str]
) -> dict[str, list[tuple[Relationship, str]]]:
    layouts: dict[str, list[tuple[Relationship, str]]] = {}
    for master_part in master_parts:
        rels_path = _rels_path_for_part(master_part)
        layouts[master_part] = []
        if rels_path not in names:
            continue
        for rel in _read_relationships(pptx, rels_path):
            if rel.rel_type != SLIDE_LAYOUT_REL_TYPE or rel.target_mode == "External":
                continue
            layout_part = _resolve_target(master_part, rel.target)
            if layout_part in names:
                layouts[master_part].append((rel, layout_part))
    return layouts


def _remove_layout_ids(
    data: bytes, master_part: str, unused_layouts: Iterable[LayoutCandidate]
) -> bytes:
    rel_ids = {
        candidate.relationship_id
        for candidate in unused_layouts
        if candidate.master_part_name == master_part and candidate.relationship_id
    }
    if not rel_ids:
        return data
    root = etree.fromstring(data)
    layout_list = root.find("p:sldLayoutIdLst", NS)
    if layout_list is not None:
        for elem in list(layout_list):
            if elem.attrib.get(f"{{{R_NS}}}id") in rel_ids:
                layout_list.remove(elem)
    return etree.tostring(root, encoding="utf-8", xml_declaration=True)


def _remove_layout_relationships(
    data: bytes, master_part: str | None, unused_layouts: Iterable[LayoutCandidate]
) -> bytes:
    rel_ids = {
        candidate.relationship_id
        for candidate in unused_layouts
        if candidate.master_part_name == master_part and candidate.relationship_id
    }
    if not rel_ids:
        return data
    root = etree.fromstring(data)
    for rel in list(root):
        if rel.attrib.get("Id") in rel_ids:
            root.remove(rel)
    return etree.tostring(root, encoding="utf-8", xml_declaration=True)


def _collect_only_deleted_references(
    pptx: ZipFile, names: set[str], initial_delete_parts: set[str]
) -> set[str]:
    delete_parts = set(initial_delete_parts)
    while True:
        incoming_from_kept = _incoming_reference_counts(pptx, names, delete_parts)
        next_delete = set(delete_parts)
        for part in list(delete_parts):
            rels_path = _rels_path_for_part(part)
            if rels_path not in names:
                continue
            for rel in _read_relationships(pptx, rels_path):
                if rel.target_mode == "External":
                    continue
                target = _resolve_target(part, rel.target)
                if target in names and incoming_from_kept.get(target, 0) == 0:
                    next_delete.add(target)
                    next_delete.add(_rels_path_for_part(target))
        if next_delete == delete_parts:
            return delete_parts
        delete_parts = next_delete


def _incoming_reference_counts(
    pptx: ZipFile, names: set[str], delete_parts: set[str]
) -> dict[str, int]:
    counts: dict[str, int] = {}
    for rels_path in (name for name in names if name.endswith(".rels")):
        source_part = _part_for_rels_path(rels_path)
        if source_part is None or source_part in delete_parts or rels_path in delete_parts:
            continue
        for rel in _read_relationships(pptx, rels_path):
            if rel.target_mode == "External":
                continue
            target = _resolve_target(source_part, rel.target)
            counts[target] = counts.get(target, 0) + 1
    return counts


def _part_for_rels_path(rels_path: str) -> str | None:
    path = PurePosixPath(rels_path)
    if path.parent.name != "_rels" or not path.name.endswith(".rels"):
        return None
    return str(path.parent.parent / path.name.removesuffix(".rels"))


def _remove_content_type_overrides(data: bytes, delete_parts: set[str]) -> bytes:
    root = etree.fromstring(data)
    for override in list(root):
        part_name = override.attrib.get("PartName", "").lstrip("/")
        if part_name in delete_parts:
            root.remove(override)
    return etree.tostring(root, encoding="utf-8", xml_declaration=True)


def _validate_package_integrity(path: Path) -> tuple[str, ...]:
    warnings: list[str] = []
    with ZipFile(path) as pptx:
        names = set(pptx.namelist())
        for rels_path in sorted(name for name in names if name.endswith(".rels")):
            source_part = _part_for_rels_path(rels_path)
            if source_part is None:
                continue
            for rel in _read_relationships(pptx, rels_path):
                if rel.target_mode == "External":
                    continue
                target = _resolve_target(source_part, rel.target)
                if target not in names:
                    warnings.append(f"{rels_path}:{rel.rel_id}->{target}")
        for part_name in sorted(_xml_part_names(names)):
            rel_ids = {
                rel.rel_id
                for rel in _read_relationships(pptx, _rels_path_for_part(part_name))
            } if _rels_path_for_part(part_name) in names else set()
            for rel_id in _relationship_ids_used_in_part(pptx.read(part_name)):
                if rel_id not in rel_ids:
                    warnings.append(f"{part_name}:missing relationship {rel_id}")
    return tuple(warnings)


def _xml_part_names(names: set[str]) -> Iterable[str]:
    for name in names:
        if name.endswith(".xml") and not name.endswith(".rels") and name != "[Content_Types].xml":
            yield name


def _relationship_ids_used_in_part(data: bytes) -> set[str]:
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return set()
    rel_attr_name = f"{{{R_NS}}}id"
    return {
        value
        for elem in root.iter()
        for attr_name, value in elem.attrib.items()
        if attr_name == rel_attr_name
    }


def _copy_pptx(source: Path, destination: Path) -> None:
    with ZipFile(source) as src, ZipFile(destination, "w", ZIP_DEFLATED) as dst:
        written: set[str] = set()
        for info in src.infolist():
            _write_once(dst, info, src.read(info.filename), written)


def _write_once(dst: ZipFile, info, data: bytes, written: set[str]) -> None:
    if info.filename in written:
        return
    dst.writestr(info, data)
    written.add(info.filename)
