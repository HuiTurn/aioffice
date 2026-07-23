"""Structured WordprocessingML field parsing, generation, and patching."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal
from xml.etree import ElementTree as ET

from aioffice.formats.docx_style import apply_text_style
from aioffice.native.identity import fingerprint_elements
from aioffice.spec.models import DocumentField, NativeRef, TextStyle

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
XML = "http://www.w3.org/XML/1998/namespace"

FieldForm = Literal["simple", "complex"]
FieldKind = Literal[
    "page_number",
    "page_count",
    "section_number",
    "section_page_count",
    "native",
]
FieldNumberFormat = Literal[
    "decimal",
    "upper_roman",
    "lower_roman",
    "upper_letter",
    "lower_letter",
]

_KIND_FROM_CODE: dict[str, FieldKind] = {
    "PAGE": "page_number",
    "NUMPAGES": "page_count",
    "SECTION": "section_number",
    "SECTIONPAGES": "section_page_count",
}
_CODE_FROM_KIND = {value: key for key, value in _KIND_FROM_CODE.items()}
_FORMAT_FROM_NATIVE: dict[str, FieldNumberFormat] = {
    "ARABIC": "decimal",
    "ROMAN": "upper_roman",
    "roman": "lower_roman",
    "ALPHABETIC": "upper_letter",
    "alphabetic": "lower_letter",
}
_FORMAT_TO_NATIVE = {value: key for key, value in _FORMAT_FROM_NATIVE.items()}
_FORMAT_SWITCH = re.compile(r"\\\*\s+(\"[^\"]*\"|\S+)")


class FieldStructureError(ValueError):
    """Raised when field markup cannot be isolated without guessing."""


@dataclass(slots=True)
class FieldMatch:
    ordinal: int
    form: FieldForm
    start_index: int
    end_index: int
    elements: list[ET.Element]
    instruction: str
    instruction_nodes: list[ET.Element]
    result_elements: list[ET.Element]
    begin_char: ET.Element | None
    simple_element: ET.Element | None
    nested: bool = False

    @property
    def cached_result(self) -> str | None:
        text = "".join(
            node.text or ""
            for element in self.result_elements
            for node in element.iter(_q("t"))
        )
        return text or None

    @property
    def dirty(self) -> bool:
        element = (
            self.simple_element
            if self.simple_element is not None
            else self.begin_char
        )
        if element is None:
            return False
        value = element.get(_q("dirty"))
        return value is not None and value.casefold() not in {
            "0",
            "false",
            "off",
            "no",
        }


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _field_chars(element: ET.Element) -> list[ET.Element]:
    return list(element.iter(_q("fldChar")))


def _field_char_type(element: ET.Element) -> str | None:
    return element.get(_q("fldCharType"))


def _contains_field_markup(element: ET.Element) -> bool:
    return (
        element.tag == _q("fldSimple")
        or element.find(f".//{_q('fldChar')}") is not None
        or element.find(f".//{_q('instrText')}") is not None
    )


def parse_paragraph_fields(paragraph: ET.Element) -> list[FieldMatch]:
    """Return direct, isolated fields in document order.

    Complex fields are represented by their complete begin/instruction/result/end
    child range. Field markup nested in another inline container is rejected
    because changing it safely would require reconstructing unknown containment.
    """

    children = list(paragraph)
    matches: list[FieldMatch] = []
    index = 0
    while index < len(children):
        child = children[index]
        if child.tag == _q("fldSimple"):
            instruction = child.get(_q("instr"), "")
            matches.append(
                FieldMatch(
                    ordinal=len(matches),
                    form="simple",
                    start_index=index,
                    end_index=index,
                    elements=[child],
                    instruction=instruction,
                    instruction_nodes=[],
                    result_elements=[child],
                    begin_char=None,
                    simple_element=child,
                )
            )
            index += 1
            continue

        chars = _field_chars(child)
        if not chars:
            if _contains_field_markup(child):
                raise FieldStructureError(
                    "Field instruction markup is not bounded by a direct complex field."
                )
            index += 1
            continue
        if (
            len(chars) != 1
            or _field_char_type(chars[0]) != "begin"
            or child.tag != _q("r")
        ):
            raise FieldStructureError(
                "Complex field does not begin in one direct paragraph run."
            )

        depth = 0
        nested = False
        separator_index: int | None = None
        end_index: int | None = None
        begin_char = chars[0]
        for candidate_index in range(index, len(children)):
            candidate = children[candidate_index]
            for field_char in _field_chars(candidate):
                field_type = _field_char_type(field_char)
                if field_type == "begin":
                    depth += 1
                    if depth > 1:
                        nested = True
                elif field_type == "separate" and depth == 1:
                    if separator_index is not None:
                        nested = True
                    else:
                        separator_index = candidate_index
                elif field_type == "end":
                    depth -= 1
                    if depth < 0:
                        raise FieldStructureError("Complex field has an unmatched end.")
                    if depth == 0:
                        end_index = candidate_index
                        break
            if end_index is not None:
                break
        if end_index is None or depth != 0:
            raise FieldStructureError("Complex field has no matching end.")

        field_elements = children[index : end_index + 1]
        instruction_limit = (
            separator_index if separator_index is not None else end_index
        )
        instruction_nodes = [
            node
            for element in children[index : instruction_limit + 1]
            for node in element.iter(_q("instrText"))
        ]
        instruction = "".join(node.text or "" for node in instruction_nodes)
        result_elements = (
            children[separator_index + 1 : end_index]
            if separator_index is not None
            else []
        )
        matches.append(
            FieldMatch(
                ordinal=len(matches),
                form="complex",
                start_index=index,
                end_index=end_index,
                elements=field_elements,
                instruction=instruction,
                instruction_nodes=instruction_nodes,
                result_elements=result_elements,
                begin_char=begin_char,
                simple_element=None,
                nested=nested,
            )
        )
        index = end_index + 1

    return matches


def normalized_field_instruction(
    instruction: str,
    *,
    nested: bool = False,
) -> tuple[
    FieldKind,
    str | None,
]:
    """Parse the safe subset of Word field instructions."""

    match = re.match(r"^\s*([A-Za-z]+)\b(.*)$", instruction, re.DOTALL)
    if match is None or nested:
        return "native", None
    kind = _KIND_FROM_CODE.get(match.group(1).upper())
    if kind is None:
        return "native", None
    remainder = match.group(2)
    formats = [value.strip('"') for value in _FORMAT_SWITCH.findall(remainder)]
    residual = _FORMAT_SWITCH.sub(" ", remainder).strip()
    if residual:
        return "native", None
    normalized_formats = [
        (
            "decimal"
            if value.casefold() == "arabic"
            else _FORMAT_FROM_NATIVE[value]
        )
        for value in formats
        if value.casefold() == "arabic"
        or value in _FORMAT_FROM_NATIVE
    ]
    unknown_formats = [
        value
        for value in formats
        if value.casefold() not in {"arabic", "mergeformat"}
        and value not in _FORMAT_FROM_NATIVE
    ]
    if unknown_formats or len(normalized_formats) > 1:
        return "native", None
    return kind, normalized_formats[0] if normalized_formats else None


def field_payload(
    paragraph: ET.Element,
    paragraph_index: int,
    match: FieldMatch,
    *,
    field_id: str,
    part_uri: str,
    root_path: str,
    style: TextStyle | None = None,
) -> dict[str, object]:
    """Project one isolated native field into the normalized inline model."""

    kind, number_format = normalized_field_instruction(
        match.instruction,
        nested=match.nested,
    )
    source_ref = native_ref_for_field(
        paragraph,
        paragraph_index,
        match,
        part_uri=part_uri,
        root_path=root_path,
    )
    payload: dict[str, object] = {
        "id": field_id,
        "type": "field",
        "kind": kind,
        "cached_result": match.cached_result,
        "editable": kind != "native",
        "source_ref": source_ref.model_dump(mode="json", exclude_none=True),
        "metadata": {
            "native_form": match.form,
            "native_instruction": match.instruction,
            "dirty": match.dirty,
        },
    }
    if kind == "native":
        payload["instruction"] = match.instruction or "UNKNOWN"
    elif number_format is not None:
        payload["number_format"] = number_format
    if style is not None:
        payload["style"] = style.model_dump(mode="json", exclude_none=True)
    return payload


def native_ref_for_field(
    paragraph: ET.Element,
    paragraph_index: int,
    match: FieldMatch,
    *,
    part_uri: str,
    root_path: str,
) -> NativeRef:
    para_id = paragraph.get(f"{{{W14}}}paraId")
    native_id = (
        f"{para_id}:field:{match.ordinal}"
        if para_id is not None
        else f"index:{paragraph_index}:field:{match.ordinal}"
    )
    native_kind = "w:fldSimple" if match.form == "simple" else "w:complex-field"
    return NativeRef(
        format="docx",
        part_uri=part_uri,
        native_kind=native_kind,
        element_index=paragraph_index,
        element_indices=[paragraph_index],
        sub_index=match.ordinal,
        path_hint=(
            f"{root_path}/*[{paragraph_index + 1}]"
            f"/w:field[{match.ordinal + 1}]"
        ),
        native_id=native_id,
        fingerprint=fingerprint_elements(match.elements),
    )


def canonical_field_instruction(field: DocumentField) -> str:
    if field.kind == "native":
        raise ValueError("A native field instruction is not editable.")
    code = _CODE_FROM_KIND[field.kind]
    fragments = [code]
    if field.number_format is not None:
        fragments.extend(["\\*", _FORMAT_TO_NATIVE[field.number_format]])
    fragments.extend(["\\*", "MERGEFORMAT"])
    return " " + " ".join(fragments) + " "


def append_complex_field(
    parent: ET.Element,
    field: DocumentField,
    *,
    effective_style: TextStyle | None,
) -> list[ET.Element]:
    """Append a conventional complex field and return its direct run elements."""

    if field.kind == "native":
        raise ValueError("Semantic DOCX generation cannot compile a native-only field.")

    def add_run() -> ET.Element:
        run = ET.SubElement(parent, _q("r"))
        apply_text_style(run, effective_style)
        return run

    begin_run = add_run()
    ET.SubElement(
        begin_run,
        _q("fldChar"),
        {
            _q("fldCharType"): "begin",
            _q("dirty"): "1",
        },
    )
    instruction_run = add_run()
    instruction = ET.SubElement(
        instruction_run,
        _q("instrText"),
        {f"{{{XML}}}space": "preserve"},
    )
    instruction.text = canonical_field_instruction(field)
    separator_run = add_run()
    ET.SubElement(
        separator_run,
        _q("fldChar"),
        {_q("fldCharType"): "separate"},
    )
    result_run = add_run()
    result = ET.SubElement(result_run, _q("t"))
    result.text = field.cached_result if field.cached_result is not None else "1"
    if (
        result.text[:1].isspace()
        or result.text[-1:].isspace()
        or "  " in result.text
    ):
        result.set(f"{{{XML}}}space", "preserve")
    end_run = add_run()
    ET.SubElement(
        end_run,
        _q("fldChar"),
        {_q("fldCharType"): "end"},
    )
    return [
        begin_run,
        instruction_run,
        separator_run,
        result_run,
        end_run,
    ]


def patch_field_instruction(match: FieldMatch, field: DocumentField) -> None:
    """Change only the instruction payload and dirty flag of one native field."""

    instruction = canonical_field_instruction(field)
    if match.form == "simple":
        assert match.simple_element is not None
        match.simple_element.set(_q("instr"), instruction)
        match.simple_element.set(_q("dirty"), "1")
        return
    if not match.instruction_nodes or match.begin_char is None:
        raise FieldStructureError(
            "Complex field has no editable instruction text or begin character."
        )
    match.instruction_nodes[0].text = instruction
    match.instruction_nodes[0].set(f"{{{XML}}}space", "preserve")
    for node in match.instruction_nodes[1:]:
        node.text = ""
    match.begin_char.set(_q("dirty"), "1")


def field_match_at(paragraph: ET.Element, sub_index: int) -> FieldMatch:
    matches = parse_paragraph_fields(paragraph)
    if sub_index >= len(matches):
        raise FieldStructureError(
            f"Field sub-index {sub_index} points outside its native paragraph."
        )
    return matches[sub_index]


__all__ = [
    "FieldMatch",
    "FieldStructureError",
    "append_complex_field",
    "canonical_field_instruction",
    "field_match_at",
    "field_payload",
    "native_ref_for_field",
    "normalized_field_instruction",
    "parse_paragraph_fields",
    "patch_field_instruction",
]
