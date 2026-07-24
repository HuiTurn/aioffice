from __future__ import annotations

import copy
import hashlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Workspace
from aioffice.core.errors import WorkspaceError
from aioffice.documents import Document, DocumentBuilder
from aioffice.native import (
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
    build_identity_manifest,
)

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _rewrite_package(
    source: bytes,
    *,
    replacements: dict[str, bytes] | None = None,
    removals: set[str] | None = None,
) -> bytes:
    replacements = replacements or {}
    removals = removals or set()
    output = io.BytesIO()
    with (
        ZipFile(io.BytesIO(source)) as input_archive,
        ZipFile(
            output,
            "w",
            compression=ZIP_DEFLATED,
        ) as output_archive,
    ):
        for info in input_archive.infolist():
            if info.filename in removals:
                continue
            payload = replacements.get(
                info.filename,
                input_archive.read(info.filename),
            )
            output_archive.writestr(copy.copy(info), payload)
    return output.getvalue()


def _without_embedded_identity(source: bytes, *, strip_native_ids: bool = False) -> bytes:
    with ZipFile(io.BytesIO(source)) as archive:
        relationships = ET.fromstring(archive.read("_rels/.rels"))
        document = ET.fromstring(archive.read("word/document.xml"))
    for relationship in list(relationships):
        if relationship.attrib.get("Type") == MANIFEST_RELATIONSHIP_TYPE:
            relationships.remove(relationship)
    if strip_native_ids:
        for element in document.iter():
            element.attrib.pop(f"{{{W14}}}paraId", None)
    return _rewrite_package(
        source,
        replacements={
            "_rels/.rels": ET.tostring(
                relationships,
                encoding="utf-8",
                xml_declaration=True,
            ),
            "word/document.xml": ET.tostring(
                document,
                encoding="utf-8",
                xml_declaration=True,
            ),
        },
        removals={MANIFEST_PART_URI.lstrip("/")},
    )


class WorkspaceTests(unittest.TestCase):
    def _source(self) -> bytes:
        return (
            DocumentBuilder(title="Workspace")
            .heading("Draft", id="status")
            .bullet_list(["One", "Two"], id="items")
            .build()
            .to_bytes("docx")
        )

    def test_workspace_persists_revisions_and_idempotent_patches(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "third-party.docx"
            source_bytes = _without_embedded_identity(self._source())
            source.write_bytes(source_bytes)

            workspace = Workspace.init(root / "project")
            imported = workspace.import_document(source)
            artifact_id = imported.id
            original_ids = [node["id"] for node in imported.to_spec()["content"]]
            self.assertEqual(imported.to_bytes("docx"), source_bytes)

            operations = [
                {
                    "op": "text.replace",
                    "target": f"#{original_ids[0]}",
                    "search": "Draft",
                    "replacement": "Approved",
                }
            ]
            result = workspace.apply(
                artifact_id,
                operations,
                idempotency_key="approve-status",
            )
            self.assertTrue(result.success)
            self.assertEqual(result.result_revision, 2)
            self.assertIsNotNone(result.diff)

            reopened_workspace = Workspace.open(root / "project")
            revision_one = reopened_workspace.checkout(artifact_id, revision=1)
            revision_two = reopened_workspace.open_document(artifact_id)
            self.assertEqual(revision_one.revision, 1)
            self.assertEqual(revision_two.revision, 2)
            self.assertEqual(
                [node["id"] for node in revision_two.to_spec()["content"]],
                original_ids,
            )
            self.assertIn("Approved", revision_two.to_json())

            replay = reopened_workspace.apply(
                artifact_id,
                operations,
                idempotency_key="approve-status",
            )
            self.assertEqual(replay.result_revision, 2)
            self.assertEqual(replay.diff, result.diff)
            self.assertEqual(
                reopened_workspace.list_artifacts()[0]["latest_revision"],
                2,
            )
            with self.assertRaises(WorkspaceError):
                reopened_workspace.apply(
                    artifact_id,
                    [{"op": "node.remove", "target": f"#{original_ids[1]}"}],
                    idempotency_key="approve-status",
                )

            external = root / "external-edit.docx"
            external_bytes = revision_two.to_bytes("docx")
            with ZipFile(io.BytesIO(external_bytes)) as archive:
                external_xml = archive.read("word/document.xml").replace(
                    b"Approved",
                    b"Externally approved",
                )
            external.write_bytes(
                _rewrite_package(
                    external_bytes,
                    replacements={"word/document.xml": external_xml},
                )
            )
            preview = reopened_workspace.reconcile_document(
                artifact_id,
                external,
            )
            self.assertEqual(preview.revision, 3)
            self.assertEqual(preview.import_diagnostics, [])
            self.assertEqual(
                reopened_workspace.list_artifacts()[0]["latest_revision"],
                2,
            )
            reconciled = reopened_workspace.reconcile_document(
                artifact_id,
                external,
                commit=True,
            )
            self.assertEqual(reconciled.revision, 3)
            self.assertIn("Externally approved", reconciled.to_json())
            self.assertEqual(
                [node["id"] for node in reconciled.to_spec()["content"]],
                original_ids,
            )

            exported = root / "approved.docx"
            reopened_workspace.export_document(artifact_id, exported)
            self.assertTrue(exported.exists())
            with self.assertRaises(WorkspaceError):
                reopened_workspace.export_document(artifact_id, exported)

            state = root / "project" / ".aioffice" / "artifacts" / artifact_id
            self.assertTrue((state / "revisions" / "00000001.docx").exists())
            self.assertTrue((state / "revisions" / "00000002.docx").exists())
            self.assertTrue((state / "revisions" / "00000003.docx").exists())
            self.assertTrue((state / "snapshots" / "00000003.json").exists())
            self.assertTrue((state / "patches" / "00000003.json").exists())

    def test_identity_mismatch_is_reported_instead_of_silently_reused(self) -> None:
        source = _without_embedded_identity(
            DocumentBuilder().paragraph("Draft").build().to_bytes("docx"),
            strip_native_ids=True,
        )
        document = Document.from_docx(source)
        manifest = build_identity_manifest(
            document.spec,
            package_sha256=hashlib.sha256(source).hexdigest(),
        )

        with ZipFile(io.BytesIO(source)) as archive:
            xml = archive.read("word/document.xml").replace(b"Draft", b"Changed")
        changed = _rewrite_package(
            source,
            replacements={"word/document.xml": xml},
        )
        reopened = Document.from_docx(changed, identity_manifest=manifest)
        self.assertTrue(
            any(
                diagnostic.code == "IDENTITY_AMBIGUOUS"
                for diagnostic in reopened.import_diagnostics
            )
        )
        self.assertNotEqual(
            reopened.to_spec()["content"][0]["id"],
            document.to_spec()["content"][0]["id"],
        )

    def test_workspace_refuses_ambiguous_external_edit(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source_bytes = _without_embedded_identity(
                DocumentBuilder().paragraph("Draft").build().to_bytes("docx"),
                strip_native_ids=True,
            )
            source = root / "source.docx"
            source.write_bytes(source_bytes)
            workspace = Workspace.init(root / "project")
            imported = workspace.import_document(source)

            with ZipFile(io.BytesIO(source_bytes)) as archive:
                changed_xml = archive.read("word/document.xml").replace(
                    b"Draft",
                    b"Changed",
                )
            changed = root / "changed.docx"
            changed.write_bytes(
                _rewrite_package(
                    source_bytes,
                    replacements={"word/document.xml": changed_xml},
                )
            )
            preview = workspace.reconcile_document(imported.id, changed)
            self.assertTrue(preview.import_diagnostics)
            with self.assertRaises(WorkspaceError):
                workspace.reconcile_document(imported.id, changed, commit=True)
            self.assertEqual(
                workspace.list_artifacts()[0]["latest_revision"],
                1,
            )

    def test_external_insertion_keeps_bound_ids_unique(self) -> None:
        source = _without_embedded_identity(
            (DocumentBuilder().paragraph("First").paragraph("Second").build().to_bytes("docx")),
            strip_native_ids=True,
        )
        document = Document.from_docx(source)
        original_ids = [node["id"] for node in document.to_spec()["content"]]
        manifest = build_identity_manifest(
            document.spec,
            package_sha256=hashlib.sha256(source).hexdigest(),
        )

        with ZipFile(io.BytesIO(source)) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        body = root.find(f"{{{W}}}body")
        assert body is not None
        paragraph = ET.Element(f"{{{W}}}p")
        run = ET.SubElement(paragraph, f"{{{W}}}r")
        ET.SubElement(run, f"{{{W}}}t").text = "Inserted"
        body.insert(0, paragraph)
        changed = _rewrite_package(
            source,
            replacements={
                "word/document.xml": ET.tostring(
                    root,
                    encoding="utf-8",
                    xml_declaration=True,
                )
            },
        )
        reopened = Document.from_docx(changed, identity_manifest=manifest)
        result_ids = [node["id"] for node in reopened.to_spec()["content"]]
        self.assertEqual(len(result_ids), len(set(result_ids)))
        self.assertTrue(set(original_ids).issubset(result_ids))
        self.assertNotIn(result_ids[0], original_ids)

    def test_external_deletion_is_proven_by_neighboring_identities(self) -> None:
        source = _without_embedded_identity(
            (
                DocumentBuilder()
                .paragraph("First")
                .paragraph("Remove me")
                .paragraph("Third")
                .build()
                .to_bytes("docx")
            ),
            strip_native_ids=True,
        )
        document = Document.from_docx(source)
        original_ids = [node["id"] for node in document.to_spec()["content"]]
        manifest = build_identity_manifest(
            document.spec,
            package_sha256=hashlib.sha256(source).hexdigest(),
        )

        with ZipFile(io.BytesIO(source)) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        body = root.find(f"{{{W}}}body")
        assert body is not None
        body.remove(list(body)[1])
        changed = _rewrite_package(
            source,
            replacements={
                "word/document.xml": ET.tostring(
                    root,
                    encoding="utf-8",
                    xml_declaration=True,
                )
            },
        )
        reopened = Document.from_docx(changed, identity_manifest=manifest)
        result_ids = [node["id"] for node in reopened.to_spec()["content"]]
        self.assertEqual(reopened.import_diagnostics, [])
        self.assertEqual(result_ids, [original_ids[0], original_ids[2]])

    def test_workspace_index_rejects_unknown_fields(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workspace = Workspace.init(directory)
            index_path = workspace.state_dir / "workspace.json"
            payload = json.loads(index_path.read_text(encoding="utf-8"))
            payload["unexpected"] = True
            index_path.write_text(json.dumps(payload), encoding="utf-8")
            with self.assertRaises(WorkspaceError):
                Workspace.open(directory)

    def test_workspace_persists_native_node_move_patch(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.docx"
            (
                DocumentBuilder()
                .paragraph("A", id="a")
                .paragraph("B", id="b")
                .paragraph("C", id="c")
                .build()
                .export(source)
            )
            workspace = Workspace.init(root / "project")
            document = workspace.import_document(source)
            operation = {
                "op": "node.move_before",
                "target": "#c",
                "before": "#a",
            }
            result = workspace.apply(
                document.id,
                [operation],
                base_revision=document.revision,
            )
            self.assertTrue(result.success, result.model_dump())
            reopened = workspace.open_document(document.id)
            self.assertEqual(
                [
                    node["id"]
                    for node in reopened.to_spec()["content"]
                ],
                ["c", "a", "b"],
            )
            self.assertIn(
                "node.insert_after",
                workspace.capabilities(document.id)[
                    "patch_operations"
                ],
            )
            self.assertIn(
                "node.insert_before",
                workspace.capabilities(document.id)[
                    "patch_operations"
                ],
            )
            self.assertIn(
                "node.move_before",
                workspace.capabilities(document.id)[
                    "patch_operations"
                ],
            )
            patch_path = (
                root
                / "project"
                / ".aioffice"
                / "artifacts"
                / document.id
                / "patches"
                / f"{result.result_revision:08d}.json"
            )
            patch = json.loads(
                patch_path.read_text(encoding="utf-8")
            )
            self.assertEqual(patch["operations"], [operation])
            self.assertEqual(
                patch["changes"][0]["moved_nodes"],
                ["c"],
            )
            removed = workspace.apply(
                document.id,
                [{"op": "node.remove", "target": "#b"}],
                base_revision=result.result_revision,
            )
            self.assertTrue(removed.success, removed.model_dump())
            after_remove = workspace.open_document(document.id)
            self.assertEqual(
                [
                    node["id"]
                    for node in after_remove.to_spec()["content"]
                ],
                ["c", "a"],
            )
            self.assertEqual(
                removed.changes[0]["removed_nodes"],
                ["b"],
            )
            inserted = workspace.apply(
                document.id,
                [
                    {
                        "op": "node.insert_before",
                        "target": "#c",
                        "content": {
                            "id": "inserted",
                            "type": "paragraph",
                            "text": "Inserted",
                        },
                    }
                ],
                base_revision=removed.result_revision,
            )
            self.assertTrue(inserted.success, inserted.model_dump())
            after_insert = workspace.open_document(document.id)
            self.assertEqual(
                [
                    node["id"]
                    for node in after_insert.to_spec()["content"]
                ],
                ["inserted", "c", "a"],
            )


if __name__ == "__main__":
    unittest.main()
