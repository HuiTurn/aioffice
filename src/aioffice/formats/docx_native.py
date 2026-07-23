"""Lower semantic operations into minimal mutations of a native DOCX part."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from copy import deepcopy
from typing import Any
from xml.etree import ElementTree as ET

from pydantic import ValidationError

from aioffice.core.errors import NativePackageError
from aioffice.formats.docx_style import (
    patch_paragraph_mark_text_style,
    patch_paragraph_style,
    patch_paragraph_style_ref,
    patch_text_style,
)
from aioffice.formats.docx_named_styles import (
    find_named_style,
    format_named_style,
    upsert_named_style,
)
from aioffice.native import (
    FidelityReport,
    MANIFEST_PART_URI,
    NativePackage,
    build_identity_manifest,
    native_ref_for_elements,
    serialize_identity_manifest,
)
from aioffice.native.xml import parse_xml, serialize_xml
from aioffice.operations.text import resolve_text_selection
from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    Heading,
    NamedStyle,
    NativeRef,
    ParagraphStyle,
    TextStyle,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _target_id(value: Any) -> str:
    if not isinstance(value, str) or not value:
        raise NativePackageError("Native DOCX patch target must be a node ID.")
    return value[1:] if value.startswith("#") else value


def _find_source_ref(spec: AiOfficeDocumentSpec, target: Any) -> NativeRef:
    target_id = _target_id(target)
    matches = [node for node in spec.content if node.id == target_id]
    if len(matches) != 1:
        raise NativePackageError(
            f"Native DOCX target #{target_id} matched {len(matches)} semantic nodes."
        )
    source_ref = matches[0].source_ref
    if not isinstance(source_ref, NativeRef) or source_ref.format != "docx":
        raise NativePackageError(
            f"Semantic node {target_id!r} has no editable DOCX source reference."
        )
    if source_ref.part_uri != "/word/document.xml" or (
        source_ref.element_index is None and not source_ref.element_indices
    ):
        raise NativePackageError(
            f"Semantic node {target_id!r} is not mapped to the DOCX main document part."
        )
    return source_ref


def _source_indices(source_ref: NativeRef) -> list[int]:
    if source_ref.element_indices:
        return list(source_ref.element_indices)
    if source_ref.element_index is not None:
        return [source_ref.element_index]
    raise NativePackageError("DOCX source reference has no native element indices.")


def _occurrences(text: str, search: str, replace_all: bool) -> list[tuple[int, int]]:
    found: list[tuple[int, int]] = []
    start = 0
    while True:
        index = text.find(search, start)
        if index < 0:
            break
        found.append((index, index + len(search)))
        if not replace_all:
            break
        start = index + len(search)
    return found


def _replace_text_nodes(
    element: ET.Element,
    search: str,
    replacement: str,
    *,
    replace_all: bool,
) -> int:
    text_nodes = list(element.iter(_q(W, "t")))
    values = [node.text or "" for node in text_nodes]
    joined = "".join(values)
    matches = _occurrences(joined, search, replace_all)
    if not matches:
        raise NativePackageError(f"Search text {search!r} was not found in native DOCX node.")

    spans: list[tuple[int, int]] = []
    cursor = 0
    for value in values:
        spans.append((cursor, cursor + len(value)))
        cursor += len(value)

    def locate(position: int) -> tuple[int, int]:
        for node_index, (start, end) in enumerate(spans):
            if start <= position < end:
                return node_index, position - start
        raise NativePackageError("Could not map semantic text to native DOCX runs.")

    for start, end in reversed(matches):
        start_node_index, start_offset = locate(start)
        end_node_index, end_offset_last = locate(end - 1)
        end_offset = end_offset_last + 1
        if start_node_index == end_node_index:
            node = text_nodes[start_node_index]
            current = node.text or ""
            node.text = current[:start_offset] + replacement + current[end_offset:]
        else:
            start_node = text_nodes[start_node_index]
            end_node = text_nodes[end_node_index]
            start_node.text = (start_node.text or "")[:start_offset] + replacement
            for middle_index in range(start_node_index + 1, end_node_index):
                text_nodes[middle_index].text = ""
            end_node.text = (end_node.text or "")[end_offset:]

    for node in text_nodes:
        value = node.text or ""
        xml_space = _q(XML, "space")
        if value[:1].isspace() or value[-1:].isspace() or "  " in value:
            node.set(xml_space, "preserve")
        else:
            node.attrib.pop(xml_space, None)
    return len(matches)


def _set_xml_space(text: ET.Element) -> None:
    value = text.text or ""
    xml_space = _q(XML, "space")
    if value[:1].isspace() or value[-1:].isspace() or "  " in value:
        text.set(xml_space, "preserve")
    else:
        text.attrib.pop(xml_space, None)


def _run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.iter(_q(W, "t")))


def _clone_run_segment(
    run: ET.Element,
    text_children: Sequence[ET.Element],
    *,
    preserve_tail: bool,
) -> ET.Element:
    clone = ET.Element(run.tag, dict(run.attrib))
    clone.text = run.text
    properties = run.find(_q(W, "rPr"))
    if properties is not None:
        clone.append(deepcopy(properties))
    clone.extend(text_children)
    clone.tail = run.tail if preserve_tail else None
    return clone


def _split_and_format_run(
    run: ET.Element,
    *,
    run_start: int,
    selection_start: int,
    selection_end: int,
    style: TextStyle,
    fields: set[str],
    parent_map: Mapping[ET.Element, ET.Element],
) -> int:
    children = list(run)
    unsupported = [
        child.tag.rsplit("}", 1)[-1]
        for child in children
        if child.tag not in {_q(W, "rPr"), _q(W, "t")}
    ]
    if unsupported:
        raise NativePackageError(
            "A partial text range crosses a complex native run containing "
            f"{', '.join(sorted(set(unsupported)))}; refusing to duplicate or "
            "drop unknown inline content."
        )
    text_nodes = [child for child in children if child.tag == _q(W, "t")]
    segments: list[tuple[bool, list[ET.Element]]] = []
    cursor = run_start
    for text_node in text_nodes:
        value = text_node.text or ""
        node_start = cursor
        node_end = cursor + len(value)
        cursor = node_end
        cuts = {
            0,
            len(value),
            max(0, min(len(value), selection_start - node_start)),
            max(0, min(len(value), selection_end - node_start)),
        }
        ordered = sorted(cuts)
        for left, right in zip(ordered, ordered[1:]):
            if left == right:
                continue
            piece_start = node_start + left
            piece_end = node_start + right
            selected = (
                piece_start >= selection_start
                and piece_end <= selection_end
                and piece_start < piece_end
            )
            cloned_text = deepcopy(text_node)
            cloned_text.text = value[left:right]
            cloned_text.tail = text_node.tail if right == len(value) else None
            _set_xml_space(cloned_text)
            if segments and segments[-1][0] == selected:
                segments[-1][1].append(cloned_text)
            else:
                segments.append((selected, [cloned_text]))
    if not segments:
        raise NativePackageError("Could not split the selected native DOCX text run.")
    try:
        parent = parent_map[run]
    except KeyError as error:
        raise NativePackageError("Could not locate the native run parent.") from error
    index = list(parent).index(run)
    parent.remove(run)
    selected_runs = 0
    for offset, (selected, text_children) in enumerate(segments):
        clone = _clone_run_segment(
            run,
            text_children,
            preserve_tail=offset == len(segments) - 1,
        )
        if selected:
            patch_text_style(clone, style, fields)
            selected_runs += 1
        parent.insert(index + offset, clone)
    return selected_runs


def _format_text_range(
    paragraph: ET.Element,
    *,
    start: int,
    end: int,
    style: TextStyle,
    fields: set[str],
) -> int:
    runs = list(paragraph.iter(_q(W, "r")))
    parent_map = {child: parent for parent in paragraph.iter() for child in list(parent)}
    cursor = 0
    selected_runs = 0
    for run in runs:
        value = _run_text(run)
        run_start = cursor
        run_end = cursor + len(value)
        cursor = run_end
        if not value or run_end <= start or run_start >= end:
            continue
        if start <= run_start and run_end <= end:
            patch_text_style(run, style, fields)
            selected_runs += 1
        else:
            selected_runs += _split_and_format_run(
                run,
                run_start=run_start,
                selection_start=start,
                selection_end=end,
                style=style,
                fields=fields,
                parent_map=parent_map,
            )
    if selected_runs == 0:
        raise NativePackageError("The selected text range mapped to no editable DOCX runs.")
    return selected_runs


def apply_docx_operations(
    package: NativePackage,
    spec: AiOfficeDocumentSpec,
    result_spec: AiOfficeDocumentSpec,
    operations: Sequence[Mapping[str, Any]],
) -> tuple[NativePackage, FidelityReport, dict[str, NativeRef]]:
    supported = {
        "text.replace",
        "paragraph.format",
        "text.format",
        "node.remove",
        "style.apply",
        "style.define",
        "style.format",
    }
    unsupported = sorted({str(operation.get("op")) for operation in operations} - supported)
    if unsupported:
        raise NativePackageError(
            "Imported DOCX V0.2 currently supports native lowering for "
            "text.replace, paragraph.format, text.format, node.remove, "
            "style.apply, style.define, and style.format; "
            f"unsupported: {', '.join(unsupported)}."
        )

    updated = package.clone()
    root = parse_xml(updated.get_part("/word/document.xml"))
    body = root.find(_q(W, "body"))
    if body is None:
        raise NativePackageError("DOCX main document part has no w:body.")
    style_operations = {
        "style.apply",
        "style.define",
        "style.format",
    }.intersection(str(operation.get("op")) for operation in operations)
    styles_root: ET.Element | None = None
    styles_changed = False
    document_changed = False
    if style_operations:
        if not updated.has_part("/word/styles.xml"):
            raise NativePackageError(
                "Native DOCX has no /word/styles.xml part for named-style editing."
            )
        styles_root = parse_xml(updated.get_part("/word/styles.xml"))
    original_elements = list(body)
    source_elements: dict[str, list[ET.Element]] = {}
    for node in spec.content:
        source_ref = node.source_ref
        if not isinstance(source_ref, NativeRef):
            continue
        indices = _source_indices(source_ref)
        if (
            source_ref.part_uri == "/word/document.xml"
            and indices
            and all(index < len(original_elements) for index in indices)
        ):
            source_elements[node.id] = [original_elements[index] for index in indices]

    for operation in operations:
        operation_name = operation.get("op")
        if operation_name == "style.define":
            assert styles_root is not None
            try:
                named_style = NamedStyle.model_validate(operation.get("style"))
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower style.define values: {error}"
                ) from error
            if find_named_style(styles_root, named_style.id) is not None:
                raise NativePackageError(
                    f"Native DOCX paragraph style {named_style.id!r} already exists."
                )
            upsert_named_style(styles_root, named_style, custom_style=True)
            styles_changed = True
            continue
        if operation_name == "style.format":
            assert styles_root is not None
            paragraph_scope = operation.get("paragraph", {})
            text_scope = operation.get("text", {})
            paragraph_set = paragraph_scope.get("set", {})
            paragraph_fields = set(paragraph_set) | set(
                paragraph_scope.get("clear", [])
            )
            text_set = text_scope.get("set", {})
            text_fields = set(text_set) | set(text_scope.get("clear", []))
            try:
                paragraph_style = ParagraphStyle.model_validate(paragraph_set)
                text_style = TextStyle.model_validate(text_set)
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower style.format values: {error}"
                ) from error
            raw_target = operation.get("target")
            if not isinstance(raw_target, str) or not raw_target:
                raise NativePackageError("style.format requires a named style ID.")
            style_id = raw_target
            if (
                find_named_style(styles_root, style_id) is None
                and raw_target.startswith("@")
            ):
                style_id = raw_target[1:]
            format_named_style(
                styles_root,
                style_id,
                paragraph_style=paragraph_style,
                paragraph_fields=paragraph_fields,
                text_style=text_style,
                text_fields=text_fields,
            )
            styles_changed = True
            continue

        source_ref = _find_source_ref(spec, operation.get("target"))
        indices = _source_indices(source_ref)
        if any(index >= len(original_elements) for index in indices):
            raise NativePackageError("DOCX source reference points outside w:body.")
        elements = [original_elements[index] for index in indices]
        if operation_name == "style.apply":
            if len(elements) != 1 or elements[0].tag != _q(W, "p"):
                raise NativePackageError(
                    "style.apply requires a native reference to one w:p element."
                )
            style_ref = operation.get("style_ref")
            native_style_ref = style_ref
            if native_style_ref is None:
                target_id = _target_id(operation.get("target"))
                result_node = next(
                    (node for node in result_spec.content if node.id == target_id),
                    None,
                )
                if isinstance(result_node, Heading):
                    native_style_ref = f"Heading{result_node.level}"
            if native_style_ref is not None:
                if not isinstance(native_style_ref, str) or not native_style_ref:
                    raise NativePackageError(
                        "style.apply style_ref must be a non-empty string or null."
                    )
                assert styles_root is not None
                if find_named_style(styles_root, native_style_ref) is None:
                    raise NativePackageError(
                        f"Native DOCX has no paragraph style {native_style_ref!r}."
                    )
            patch_paragraph_style_ref(elements[0], native_style_ref)
            document_changed = True
        elif operation_name == "text.replace":
            if len(elements) != 1:
                raise NativePackageError(
                    "text.replace requires a native reference to exactly one element."
                )
            search = operation.get("search")
            replacement = operation.get("replacement")
            if not isinstance(search, str) or not search or not isinstance(replacement, str):
                raise NativePackageError(
                    "text.replace requires a non-empty search and string replacement."
                )
            _replace_text_nodes(
                elements[0],
                search,
                replacement,
                replace_all=bool(operation.get("replace_all", False)),
            )
            document_changed = True
        elif operation_name == "paragraph.format":
            if len(elements) != 1 or elements[0].tag != _q(W, "p"):
                raise NativePackageError(
                    "paragraph.format requires a native reference to one w:p element."
                )
            fields = set(operation.get("set", {})) | set(operation.get("clear", []))
            try:
                style = ParagraphStyle.model_validate(operation.get("set", {}))
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower paragraph.format values: {error}"
                ) from error
            patch_paragraph_style(elements[0], style, fields)
            document_changed = True
        elif operation_name == "text.format":
            if len(elements) != 1 or elements[0].tag != _q(W, "p"):
                raise NativePackageError(
                    "text.format requires a native reference to one w:p element."
                )
            fields = set(operation.get("set", {})) | set(operation.get("clear", []))
            try:
                style = TextStyle.model_validate(operation.get("set", {}))
            except ValidationError as error:
                raise NativePackageError(f"Could not lower text.format values: {error}") from error
            text = "".join(node.text or "" for node in elements[0].iter(_q(W, "t")))
            try:
                selection = resolve_text_selection(
                    text,
                    range_value=operation.get("range"),
                    match_value=operation.get("match"),
                )
            except (ValidationError, ValueError) as error:
                raise NativePackageError(
                    f"Could not resolve native text.format selection: {error}"
                ) from error
            if selection is None:
                runs = list(elements[0].iter(_q(W, "r")))
                patch_paragraph_mark_text_style(elements[0], style, fields)
                for run in runs:
                    patch_text_style(run, style, fields)
            else:
                _format_text_range(
                    elements[0],
                    start=selection.start,
                    end=selection.end,
                    style=style,
                    fields=fields,
                )
            document_changed = True
        elif operation_name == "node.remove":
            for element in elements:
                if element not in list(body):
                    raise NativePackageError("DOCX node has already been removed by this patch.")
                body.remove(element)
            document_changed = True

    current_indices = {id(element): index for index, element in enumerate(list(body))}
    identity_updates: dict[str, NativeRef] = {}
    for node_id, elements in source_elements.items():
        indices = [current_indices.get(id(element)) for element in elements]
        if any(index is None for index in indices):
            continue
        normalized_indices = [index for index in indices if index is not None]
        original_ref = next(
            node.source_ref
            for node in spec.content
            if node.id == node_id and isinstance(node.source_ref, NativeRef)
        )
        assert isinstance(original_ref, NativeRef)
        identity_updates[node_id] = native_ref_for_elements(
            elements,
            normalized_indices,
            native_kind=original_ref.native_kind,
            native_id=original_ref.native_id,
        )

    if document_changed:
        updated.set_part("/word/document.xml", serialize_xml(root))
    if styles_changed:
        assert styles_root is not None
        updated.set_part("/word/styles.xml", serialize_xml(styles_root))
    if updated.has_part(MANIFEST_PART_URI):
        manifest_spec = result_spec.model_copy(deep=True)
        for node in manifest_spec.content:
            if node.id in identity_updates:
                node.source_ref = identity_updates[node.id]
        updated.set_part(
            MANIFEST_PART_URI,
            serialize_identity_manifest(build_identity_manifest(manifest_spec)),
        )
    return updated, updated.fidelity_report(), identity_updates


__all__ = ["apply_docx_operations"]
