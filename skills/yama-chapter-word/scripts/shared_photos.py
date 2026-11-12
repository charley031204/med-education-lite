"""Source-bound shared-photo adjudication, grouping, and DOCX transport.

Only byte/pixel equality is automatic. Perceptual distance orders the complete
review queue; it never determines identity or excludes a pair from review.
"""
from __future__ import annotations

import argparse
from collections.abc import Mapping
import copy
import hashlib
import io
import itertools
import json
from pathlib import Path
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw, ImageOps


SCHEMA = "yama-shared-photos/1"
DOCX_PART = "customXml/yama-shared-photos.json"
DOCX_REL_TYPE = "urn:med-education:yama-shared-photos"
DOCX_ALIAS_RELS = "customXml/_rels/yama-shared-photos.json.rels"
_SHA = re.compile(r"^[0-9a-f]{64}$")


def canonical_json(value):
    """Return deterministic UTF-8 JSON bytes; reject non-JSON values and NaN."""
    return json.dumps(value, ensure_ascii=False, sort_keys=True,
                      separators=(",", ":"), allow_nan=False).encode("utf-8")


def digest_json(value):
    return hashlib.sha256(canonical_json(value)).hexdigest()


def source_digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def inventory_digest(manifest):
    return digest_json({"schema": manifest["schema"],
                        "source_digest": manifest["source_digest"],
                        "inventory": manifest["inventory"]})


def manifest_digest(manifest):
    return digest_json(manifest)


def _image_entry(entry):
    if isinstance(entry, str):
        return {"path": entry}
    if not isinstance(entry, dict) or not isinstance(entry.get("path"), str):
        raise ValueError("Each image must be a path or an object with a path.")
    return copy.deepcopy(entry)


def _normalize_questions(questions):
    questions = json.loads(canonical_json(questions))
    if not isinstance(questions, list):
        raise ValueError("questions must be a list")
    seen, occurrences = set(), set()
    for question in questions:
        if not isinstance(question, dict):
            raise ValueError("Each question must be an object")
        qid, section = question.get("id"), question.get("section")
        if not isinstance(qid, str) or not qid or qid in seen:
            raise ValueError("Question ids must be unique nonempty strings")
        if not isinstance(section, str) or not section:
            raise ValueError("Every question needs its exact section identifier")
        seen.add(qid)
        if not isinstance(question.get("images"), list):
            raise ValueError("Every question needs an images list, including []")
        question["images"] = [_image_entry(entry) for entry in question["images"]]
        for index, entry in enumerate(question["images"]):
            if not entry["path"]:
                raise ValueError("Empty image path")
            oid = entry.setdefault("occurrence_id", f"{qid}:image:{index + 1}")
            if not isinstance(oid, str) or not oid or oid in occurrences:
                raise ValueError("Image occurrence ids must be unique nonempty strings")
            occurrences.add(oid)
    return questions


def _file_record(path, image_root, fid):
    resolved = _resolve_image(path, image_root)
    data = resolved.read_bytes()
    with Image.open(io.BytesIO(data)) as image:
        # Multi-frame sources cannot be reduced to their first frame silently.
        if getattr(image, "n_frames", 1) != 1:
            raise ValueError(f"Multi-frame image requires explicit extraction: {path}")
        rgba = ImageOps.exif_transpose(image).convert("RGBA")
        profile = image.info.get("icc_profile", b"") or b""
        # Different display orientation/color profiles must not become exact
        # merely because their encoded RGB sample arrays happen to match.
        pixel = hashlib.sha256(canonical_json(list(rgba.size))
                               + hashlib.sha256(profile).digest() + rgba.tobytes()).hexdigest()
        thumb = ImageOps.contain(rgba.convert("RGB"), (16, 16))
        normalized = Image.new("RGB", (16, 16), "white")
        normalized.paste(thumb, ((16 - thumb.width) // 2, (16 - thumb.height) // 2))
        return {"id": fid, "path": path, "sha256": hashlib.sha256(data).hexdigest(),
                "pixel_sha256": pixel, "width": image.width, "height": image.height,
                "format": image.format, "normalized_rgb": list(normalized.tobytes())}


def _resolve_image(path, image_root):
    """An alias map allows original paths to remain immutable after extraction."""
    if isinstance(image_root, Mapping):
        if path not in image_root:
            raise ValueError(f"Missing original-image alias: {path}")
        return Path(image_root[path])
    return Path(image_root) / path


def _exact_groups(files):
    groups, by_pixels = [], {}
    for file in files:
        key = file["pixel_sha256"]
        if key not in by_pixels:
            group = {"id": f"e{len(groups) + 1:04}", "file_ids": [],
                     "representative_file_id": file["id"]}
            groups.append(group)
            by_pixels[key] = group
        by_pixels[key]["file_ids"].append(file["id"])
    return groups


def _review_pairs(groups, files, evidence_digest):
    by_id = {file["id"]: file for file in files}
    pairs = []
    for left, right in itertools.combinations(groups, 2):
        a = by_id[left["representative_file_id"]]["normalized_rgb"]
        b = by_id[right["representative_file_id"]]["normalized_rgb"]
        distance = round(sum(abs(x - y) for x, y in zip(a, b)) / (len(a) * 255), 8)
        pairs.append({"id": f"{left['id']}:{right['id']}", "left": left["id"],
                      "right": right["id"], "distance_hint": distance,
                      "verdict": "pending", "rationale": "", "reviewer": "",
                      "independent_reviewer": "", "independent_verdict": "pending",
                      "independent_rationale": "", "evidence_digest": evidence_digest})
    return sorted(pairs, key=lambda pair: (pair["distance_hint"], pair["id"]))


def _byte_representatives(files):
    seen, result = set(), []
    for file in files:
        if file["sha256"] not in seen:
            seen.add(file["sha256"])
            result.append(file["id"])
    return result


def prepare_manifest(questions, source_digest, image_root):
    """Inventory every appearance; return a review draft, never an approval."""
    if not isinstance(source_digest, str) or not _SHA.fullmatch(source_digest):
        raise ValueError("source_digest must be a lowercase SHA-256 digest")
    normalized = _normalize_questions(questions)
    files, appearances, by_path = [], [], {}
    for question in normalized:
        for index, entry in enumerate(question["images"]):
            path = entry["path"]
            if path not in by_path:
                file = _file_record(path, image_root, f"f{len(files) + 1:04}")
                files.append(file)
                by_path[path] = file
            appearances.append({"id": entry["occurrence_id"], "question_id": question["id"],
                                "section": question["section"], "image_index": index,
                                "file_id": by_path[path]["id"], "source_image": copy.deepcopy(entry)})
    manifest = {"schema": SCHEMA, "source_digest": source_digest,
                "inventory": {"questions": normalized, "files": files, "appearances": appearances},
                "exact_groups": _exact_groups(files), "representatives": {}}
    bound = inventory_digest(manifest)
    manifest["inventory_digest"] = bound
    manifest["pairs"] = _review_pairs(manifest["exact_groups"], files, bound)
    manifest["discovery_review"] = {"evidence_digest": bound, "viewed_file_ids": [],
                                    "independent_viewed_file_ids": [], "reviewer": "",
                                    "independent_reviewer": "", "rationale": "",
                                    "independent_rationale": ""}
    return manifest


class _Union:
    def __init__(self, ids):
        self.parent = {value: value for value in ids}

    def find(self, value):
        while self.parent[value] != value:
            value = self.parent[value]
        return value

    def join(self, left, right):
        left, right = self.find(left), self.find(right)
        self.parent[max(left, right)] = min(left, right)


def _photo_components(manifest):
    exact = manifest["exact_groups"]
    union = _Union(group["id"] for group in exact)
    for pair in manifest["pairs"]:
        if pair["verdict"] == "same":
            union.join(pair["left"], pair["right"])
    components = {}
    for group in exact:
        components.setdefault(union.find(group["id"]), []).extend(group["file_ids"])
    return [{"id": f"photo:{key}", "file_ids": ids} for key, ids in components.items()]


def _signed_review(review, errors, prefix, bound):
    if review.get("evidence_digest") != bound:
        errors.append(f"{prefix}: stale evidence digest")
    for name in ("reviewer", "independent_reviewer", "rationale", "independent_rationale"):
        if not isinstance(review.get(name), str) or not review[name].strip():
            errors.append(f"{prefix}: missing {name}")
    if review.get("reviewer", "").strip().casefold() == review.get("independent_reviewer", "").strip().casefold():
        errors.append(f"{prefix}: reviewers must be independent")


def validate_manifest(manifest, questions=None, source_digest=None, image_root=None, *, require_complete=True):
    """Return errors. Optional current inputs also detect stale/changed sources.

    Without image_root this is a pure structural/evidence validator. Builders
    must supply all three current inputs before they render a new document.
    """
    errors = []
    try:
        if manifest.get("schema") != SCHEMA:
            return ["Unsupported shared-photo manifest schema"]
        if not _SHA.fullmatch(manifest.get("source_digest", "")):
            errors.append("Invalid source digest")
        if source_digest is not None and manifest["source_digest"] != source_digest:
            errors.append("Source digest changed")
        inv = manifest["inventory"]
        normalized = _normalize_questions(inv["questions"])
        if normalized != inv["questions"]:
            errors.append("Inventory questions are not normalized")
        if questions is not None and _normalize_questions(questions) != normalized:
            errors.append("Question/image mapping or original payload changed")
        bound = inventory_digest(manifest)
        if manifest["inventory_digest"] != bound:
            errors.append("Inventory digest changed")
        files = inv["files"]
        by_id = {file["id"]: file for file in files}
        by_path = {file["path"]: file for file in files}
        if len(by_id) != len(files) or len(by_path) != len(files):
            errors.append("Duplicate inventory file id/path")
        for index, file in enumerate(files):
            if file["id"] != f"f{index + 1:04}":
                errors.append("Inventory file order/id changed")
            if (not _SHA.fullmatch(file["sha256"]) or not _SHA.fullmatch(file["pixel_sha256"])
                    or not isinstance(file["width"], int) or not isinstance(file["height"], int)
                    or file["width"] <= 0 or file["height"] <= 0
                    or len(file["normalized_rgb"]) != 768
                    or any(type(v) is not int or not 0 <= v <= 255 for v in file["normalized_rgb"])):
                errors.append(f"Invalid image metadata: {file['id']}")
            if image_root is not None:
                try:
                    current = _file_record(file["path"], image_root, file["id"])
                    if current != file:
                        errors.append(f"Image file changed: {file['path']}")
                except (OSError, ValueError) as exc:
                    errors.append(f"Image file unavailable: {file['path']}: {exc}")
        appearances, paths = [], []
        for question in normalized:
            for index, entry in enumerate(question["images"]):
                if entry["path"] not in paths:
                    paths.append(entry["path"])
                appearances.append({"id": entry["occurrence_id"], "question_id": question["id"],
                                    "section": question["section"], "image_index": index,
                                    "file_id": by_path[entry["path"]]["id"], "source_image": entry})
        if inv["appearances"] != appearances or [file["path"] for file in files] != paths:
            errors.append("Inventory omitted or altered an image appearance/file")
        expected_groups = _exact_groups(files)
        if manifest["exact_groups"] != expected_groups:
            errors.append("Exact identity groups changed")
            return errors
        expected_pairs = _review_pairs(expected_groups, files, bound)
        expected_by_id = {pair["id"]: pair for pair in expected_pairs}
        pairs = manifest["pairs"]
        if len(pairs) != len(expected_pairs) or {pair["id"] for pair in pairs} != set(expected_by_id):
            errors.append("Incomplete or duplicate pair coverage; every exact-group pair is required")
            return errors
        union = _Union(group["id"] for group in expected_groups)
        for pair in pairs:
            expected = expected_by_id[pair["id"]]
            if any(pair.get(key) != expected[key] for key in ("left", "right", "distance_hint", "evidence_digest")):
                errors.append(f"{pair['id']}: pair inventory/evidence changed")
            if pair.get("verdict") not in ("pending", "same", "different", "uncertain"):
                errors.append(f"{pair['id']}: invalid verdict")
            if require_complete:
                _signed_review(pair, errors, pair["id"], bound)
                if pair.get("verdict") not in ("same", "different"):
                    errors.append(f"{pair['id']}: unresolved pair")
                if pair.get("independent_verdict") != pair.get("verdict"):
                    errors.append(f"{pair['id']}: independent review disagrees")
                # A diagnostic label or similarity score does not establish a field.
                if pair.get("verdict") == "same" and pair.get("identity_basis") != "same_visual_field":
                    errors.append(f"{pair['id']}: same requires identity_basis=same_visual_field")
                if pair.get("verdict") == "same" and (pair.get("source_visual_information_preserved") is not True
                        or not isinstance(pair.get("preservation_rationale"), str)
                        or not pair["preservation_rationale"].strip()):
                    errors.append(f"{pair['id']}: same requires visual-information preservation evidence")
            if pair.get("verdict") == "same":
                union.join(pair["left"], pair["right"])
        for pair in pairs:
            if pair.get("verdict") == "different" and union.find(pair["left"]) == union.find(pair["right"]):
                errors.append(f"{pair['id']}: different conflicts with transitive same identity")
        if require_complete and files:
            discovery = manifest["discovery_review"]
            _signed_review(discovery, errors, "discovery", bound)
            required = set(_byte_representatives(files))
            for key in ("viewed_file_ids", "independent_viewed_file_ids"):
                viewed = discovery.get(key, [])
                if len(viewed) != len(set(viewed)) or set(viewed) != required:
                    errors.append(f"discovery: {key} must cover every unique-byte representative")
        components = {item["id"]: item for item in _photo_components(manifest)}
        if require_complete:
            exact_by_file = {fid: group["id"] for group in expected_groups for fid in group["file_ids"]}
            for photo_id, component in components.items():
                if len({exact_by_file[fid] for fid in component["file_ids"]}) > 1 and photo_id not in manifest["representatives"]:
                    errors.append(f"{photo_id}: nonexact identity requires explicit representative review")
        for photo_id, selection in manifest["representatives"].items():
            if photo_id not in components or selection.get("file_id") not in components[photo_id]["file_ids"]:
                errors.append(f"{photo_id}: canonical representative is not a component member")
            notes = selection.get("display_notes", [])
            if not isinstance(notes, list) or any(not isinstance(note, str) or not note.strip() for note in notes):
                errors.append(f"{photo_id}: display_notes must contain nonempty strings")
            if require_complete:
                _signed_review(selection, errors, f"{photo_id}: representative", bound)
                if selection.get("suitable") is not True:
                    errors.append(f"{photo_id}: representative suitability not confirmed")
                if selection.get("source_visual_information_preserved") is not True:
                    errors.append(f"{photo_id}: representative would lose original visual information")
    except (KeyError, TypeError, ValueError, AttributeError) as exc:
        errors.append(f"Malformed manifest: {exc}")
    return errors


def grouping_plan(manifest):
    """Return deterministic section-bound parents with per-occurrence membership."""
    errors = validate_manifest(manifest)
    if errors:
        raise ValueError("Shared-photo review is incomplete: " + "; ".join(errors))
    inv = manifest["inventory"]
    files = {file["id"]: file for file in inv["files"]}
    photos, by_file = [], {}
    for component in _photo_components(manifest):
        ids = component["file_ids"]
        selected = manifest["representatives"].get(component["id"], {}).get("file_id")
        if selected is None:
            selected = min(ids, key=lambda fid: (-files[fid]["width"] * files[fid]["height"], fid))
        photo = {**component, "canonical_file_id": selected, "canonical_path": files[selected]["path"],
                 "display_notes": copy.deepcopy(manifest["representatives"].get(component["id"], {}).get("display_notes", [])),
                 "occurrence_ids": [a["id"] for a in inv["appearances"] if a["file_id"] in ids]}
        photos.append(photo)
        for fid in ids:
            by_file[fid] = photo
    questions = inv["questions"]
    union = _Union(question["id"] for question in questions)
    first_owner = {}
    for appearance in inv["appearances"]:
        key = (appearance["section"], by_file[appearance["file_id"]]["id"])
        if key in first_owner:
            union.join(first_owner[key], appearance["question_id"])
        else:
            first_owner[key] = appearance["question_id"]
    parent_groups = {}
    for question in questions:
        parent_groups.setdefault(union.find(question["id"]), []).append(question)
    parents = []
    for members in parent_groups.values():
        question_ids = [question["id"] for question in members]
        source_appearances = [a for a in inv["appearances"] if a["question_id"] in question_ids]
        image_map, memberships = {}, {qid: [] for qid in question_ids}
        for appearance in source_appearances:
            photo = by_file[appearance["file_id"]]
            pid = photo["id"]
            record = image_map.setdefault(pid, {"photo_id": pid, "path": photo["canonical_path"],
                                               "file_id": photo["canonical_file_id"],
                                               "display_notes": copy.deepcopy(photo["display_notes"]),
                                               "occurrence_ids": [], "question_ids": []})
            record["occurrence_ids"].append(appearance["id"])
            if appearance["question_id"] not in record["question_ids"]:
                record["question_ids"].append(appearance["question_id"])
            memberships[appearance["question_id"]].append({**copy.deepcopy(appearance["source_image"]),
                "occurrence_id": appearance["id"], "photo_id": pid, "path": photo["canonical_path"],
                "original_path": appearance["source_image"]["path"]})
        for record in image_map.values():
            record["shared"] = len(record["question_ids"]) > 1
        parents.append({"id": f"parent:{len(parents) + 1:04}", "section": members[0]["section"],
                        "question_ids": question_ids, "photo_ids": list(image_map),
                        "images": list(image_map.values()),
                        "questions": [{"id": qid, "image_memberships": memberships[qid]} for qid in question_ids]})
    return {"schema": SCHEMA, "inventory_digest": manifest["inventory_digest"],
            "manifest_digest": manifest_digest(manifest), "photos": photos, "parents": parents}


def export_manifest(manifest, path):
    """Write a stable file without adding a timestamp or an environment path."""
    Path(path).write_bytes(canonical_json(manifest) + b"\n")


def _document_binding(parts):
    required = {"word/document.xml", "[Content_Types].xml", "_rels/.rels", DOCX_ALIAS_RELS}
    missing = required - set(parts)
    if missing:
        raise ValueError("DOCX binding part is missing: " + ", ".join(sorted(missing)))
    names = sorted(name for name in parts if name in required
                   or (name.startswith("word/") and "/_rels/" in name and name.endswith(".rels"))
                   or name.startswith("word/media/"))
    return {name: hashlib.sha256(parts[name]).hexdigest() for name in names}


def _validate_alias_relationships(parts, aliases):
    """Validate actual OPC reachability, beyond hashes of the envelope itself."""
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    root_rels = ET.fromstring(parts["_rels/.rels"])
    alias_rels = ET.fromstring(parts[DOCX_ALIAS_RELS])
    types = ET.fromstring(parts["[Content_Types].xml"])
    if (root_rels.tag != f"{{{rel_ns}}}Relationships" or alias_rels.tag != f"{{{rel_ns}}}Relationships"
            or types.tag != f"{{{types_ns}}}Types"):
        raise ValueError("Invalid shared-photo OPC relationship/content-types structure")
    manifest_refs = [item for item in root_rels if item.get("Type") == DOCX_REL_TYPE]
    if (len(manifest_refs) != 1 or manifest_refs[0].tag != f"{{{rel_ns}}}Relationship"
            or manifest_refs[0].get("Target") != DOCX_PART
            or manifest_refs[0].get("TargetMode", "Internal") != "Internal"
            or not manifest_refs[0].get("Id")
            or sum(item.get("Id") == manifest_refs[0].get("Id") for item in root_rels) != 1):
        raise ValueError("Shared-photo manifest relationship changed")
    expected_rels = {f"rId{alias['file_id']}": alias["part"].removeprefix("customXml/") for alias in aliases}
    if len(alias_rels) != len(expected_rels) or len(expected_rels) != len(aliases):
        raise ValueError("Shared-photo original-image relationship coverage changed")
    seen_ids = set()
    for item in alias_rels:
        rid = item.get("Id")
        if (item.tag != f"{{{rel_ns}}}Relationship" or rid in seen_ids or rid not in expected_rels
                or item.get("Type") != DOCX_REL_TYPE + "/source-image"
                or item.get("Target") != expected_rels[rid]
                or item.get("TargetMode", "Internal") != "Internal"):
            raise ValueError("Shared-photo original-image relationship changed")
        seen_ids.add(rid)
    expected_types = {"/" + DOCX_PART: "application/json",
                      **{"/" + alias["part"]: "application/octet-stream" for alias in aliases}}
    for name, content_type in expected_types.items():
        entries = [item for item in types if item.get("PartName") == name]
        if (len(entries) != 1 or entries[0].tag != f"{{{types_ns}}}Override"
                or entries[0].get("ContentType") != content_type):
            raise ValueError(f"Shared-photo content type changed: {name}")


def embed_docx_manifest(docx_path, manifest, payload, *, image_root=None):
    """Embed the validated decision and arbitrary source payload in an OPC part.

    The binding excludes this part and includes document XML, Word relationship
    parts, media, package relationships, content types and alias relationships.
    External layout QC owns the full DOCX digest.
    """
    errors = validate_manifest(manifest)
    if errors:
        raise ValueError("Cannot embed invalid manifest: " + "; ".join(errors))
    path = Path(docx_path)
    with zipfile.ZipFile(path) as package:
        infos = package.infolist()
        if len({info.filename for info in infos}) != len(infos):
            raise ValueError("DOCX contains duplicate ZIP members")
        parts = {info.filename: package.read(info.filename) for info in infos}
    envelope = {"schema": SCHEMA, "manifest": manifest, "manifest_digest": manifest_digest(manifest),
                "payload": payload, "payload_digest": digest_json(payload)}
    aliases = []
    if manifest["inventory"]["files"] and image_root is None:
        raise ValueError("image_root is required to preserve every original-image alias")
    for file in manifest["inventory"]["files"]:
        data = _resolve_image(file["path"], image_root).read_bytes()
        if hashlib.sha256(data).hexdigest() != file["sha256"]:
            raise ValueError(f"Original-image alias changed: {file['path']}")
        suffix = Path(file["path"]).suffix.lower()
        if not re.fullmatch(r"\.[a-z0-9]{1,8}", suffix):
            suffix = ".bin"
        part = f"customXml/yama-shared-photos-files/{file['id']}{suffix}"
        parts[part] = data
        aliases.append({"file_id": file["id"], "path": file["path"], "part": part, "sha256": file["sha256"]})
    envelope["aliases"] = aliases
    types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    types = ET.fromstring(parts["[Content_Types].xml"])
    if not any(item.get("PartName") == "/" + DOCX_PART for item in types):
        ET.SubElement(types, f"{{{types_ns}}}Override", PartName="/" + DOCX_PART,
                      ContentType="application/json")
    for alias in aliases:
        if not any(item.get("PartName") == "/" + alias["part"] for item in types):
            ET.SubElement(types, f"{{{types_ns}}}Override", PartName="/" + alias["part"],
                          ContentType="application/octet-stream")
    rels = ET.fromstring(parts["_rels/.rels"])
    alias_rels = ET.Element(f"{{{rel_ns}}}Relationships")
    for alias in aliases:
        ET.SubElement(alias_rels, f"{{{rel_ns}}}Relationship", Id=f"rId{alias['file_id']}",
                      Type=DOCX_REL_TYPE + "/source-image", Target=alias["part"].removeprefix("customXml/"))
    parts[DOCX_ALIAS_RELS] = ET.tostring(alias_rels, encoding="utf-8", xml_declaration=True)
    existing = [item for item in rels if item.get("Type") == DOCX_REL_TYPE]
    if existing:
        for item in existing:
            item.set("Target", DOCX_PART)
    else:
        ids = {item.get("Id") for item in rels}
        number = 1
        while f"rIdSharedPhotos{number}" in ids:
            number += 1
        ET.SubElement(rels, f"{{{rel_ns}}}Relationship", Id=f"rIdSharedPhotos{number}",
                      Type=DOCX_REL_TYPE, Target=DOCX_PART)
    parts["[Content_Types].xml"] = ET.tostring(types, encoding="utf-8", xml_declaration=True)
    parts["_rels/.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
    _validate_alias_relationships(parts, aliases)
    envelope["binding"] = _document_binding(parts)
    parts[DOCX_PART] = canonical_json(envelope)
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=".docx", delete=False) as temporary:
        temporary_path = Path(temporary.name)
    try:
        with zipfile.ZipFile(temporary_path, "w", compression=zipfile.ZIP_DEFLATED) as package:
            for name, data in parts.items():
                package.writestr(name, data)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)


def _read_docx_manifest(docx_path):
    """Read and verify the manifest, payload and current DOCX binding; else fail."""
    with zipfile.ZipFile(docx_path) as package:
        names = package.namelist()
        if len(set(names)) != len(names):
            raise ValueError("DOCX contains duplicate ZIP members")
        parts = {name: package.read(name) for name in names}
        if DOCX_PART not in names:
            if (DOCX_REL_TYPE.encode() in parts.get("_rels/.rels", b"")
                    or DOCX_PART.encode() in parts.get("[Content_Types].xml", b"")
                    or any(name.startswith("customXml/yama-shared-photos-files/") for name in names)):
                raise ValueError("Shared-photo manifest part is missing")
            return None
    envelope = json.loads(parts[DOCX_PART])
    if envelope.get("schema") != SCHEMA or envelope.get("binding") != _document_binding(parts):
        raise ValueError("Shared-photo DOCX binding changed")
    if envelope.get("manifest_digest") != manifest_digest(envelope["manifest"]):
        raise ValueError("Embedded shared-photo manifest changed")
    if envelope.get("payload_digest") != digest_json(envelope["payload"]):
        raise ValueError("Embedded shared-photo payload changed")
    errors = validate_manifest(envelope["manifest"])
    if errors:
        raise ValueError("Embedded shared-photo review invalid: " + "; ".join(errors))
    expected_files = envelope["manifest"]["inventory"]["files"]
    aliases = envelope.get("aliases", [])
    if len(aliases) != len(expected_files):
        raise ValueError("Embedded original-image alias coverage changed")
    for file, alias in zip(expected_files, aliases):
        if (alias.get("file_id") != file["id"] or alias.get("path") != file["path"]
                or alias.get("sha256") != file["sha256"]
                or not alias.get("part", "").startswith("customXml/yama-shared-photos-files/")
                or alias["part"] not in parts
                or hashlib.sha256(parts[alias["part"]]).hexdigest() != file["sha256"]):
            raise ValueError(f"Embedded original-image alias changed: {file['id']}")
    _validate_alias_relationships(parts, aliases)
    return envelope


def read_docx_manifest(docx_path):
    """Return a fully verified envelope, None for legacy DOCX, or ValueError."""
    try:
        return _read_docx_manifest(docx_path)
    except (KeyError, TypeError, AttributeError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise ValueError(f"Malformed shared-photo DOCX: {exc}") from exc


def extract_docx_images(docx_path, output_dir):
    """Extract verified immutable aliases to a separate folder and return a map.

    No archive path or original path is used as a destination. Existing files
    with different bytes are refused instead of silently overwritten.
    """
    envelope = read_docx_manifest(docx_path)
    if envelope is None:
        raise ValueError("DOCX has no shared-photo manifest")
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    result = {}
    with zipfile.ZipFile(docx_path) as package:
        for alias in envelope["aliases"]:
            suffix = Path(alias["part"]).suffix
            if not re.fullmatch(r"\.[a-z0-9]{1,8}", suffix):
                suffix = ".bin"
            target = output / f"{alias['file_id']}{suffix}"
            data = package.read(alias["part"])
            if target.exists() and target.read_bytes() != data:
                raise ValueError(f"Alias extraction would overwrite a different file: {target}")
            target.write_bytes(data)
            result[alias["path"]] = str(target.resolve())
    return result


def contact_sheets(manifest, image_root, output_dir, *, pairs=False, pair_ids=None):
    """Create complete inventory sheets or selected/full pair sheets and an index.

    These files are review aids. Creating them does not attest that anyone viewed
    them. Reviewers must record actual viewing and compare plausible pairs at
    native resolution before they decide identity.
    """
    errors = validate_manifest(manifest, image_root=image_root, require_complete=False)
    if errors:
        raise ValueError("Cannot render stale inventory: " + "; ".join(errors))
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    files = {file["id"]: file for file in manifest["inventory"]["files"]}
    groups = {group["id"]: group for group in manifest["exact_groups"]}
    rows = []
    if pairs:
        selected = set(pair_ids) if pair_ids is not None else None
        if selected is not None and selected - {pair["id"] for pair in manifest["pairs"]}:
            raise ValueError("Unknown review pair id")
        for pair in manifest["pairs"]:
            if selected is None or pair["id"] in selected:
                rows.append((pair["id"].replace(":", "-"), [groups[pair[key]]["representative_file_id"]
                                                          for key in ("left", "right")]))
    else:
        ids = _byte_representatives(list(files.values()))
        rows = [(f"inventory-{index // 12 + 1:03}", ids[index:index + 12]) for index in range(0, len(ids), 12)]
    result = []
    for label, ids in rows:
        columns, cell_width, cell_height = (2 if pairs else 3), (800 if pairs else 400), (650 if pairs else 340)
        canvas = Image.new("RGB", (columns * cell_width, ((len(ids) + columns - 1) // columns) * cell_height), "white")
        draw = ImageDraw.Draw(canvas)
        for index, fid in enumerate(ids):
            file = files[fid]
            x, y = index % columns * cell_width, index // columns * cell_height
            with Image.open(_resolve_image(file["path"], image_root)) as original:
                thumbnail = ImageOps.contain(ImageOps.exif_transpose(original).convert("RGB"),
                                            (cell_width - 16, cell_height - 42))
                canvas.paste(thumbnail, (x + (cell_width - thumbnail.width) // 2, y + 32))
            # ASCII identifiers keep sheets portable without a font dependency.
            draw.text((x + 8, y + 7), f"{fid}  {file['width']}x{file['height']}", fill="black")
        path = output / f"{label}.jpg"
        if path.resolve() in {_resolve_image(file["path"], image_root).resolve() for file in files.values()}:
            raise ValueError("Contact sheet destination would overwrite a source image")
        canvas.save(path, quality=95)
        result.append({"sheet": path.name, "file_ids": ids})
    index = {"inventory_digest": manifest["inventory_digest"], "sheets": result,
             "files": [{key: file[key] for key in ("id", "path", "width", "height", "sha256")} for file in files.values()]}
    export_manifest(index, output / ("pair-index.json" if pairs else "inventory-index.json"))
    return index


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    inventory = sub.add_parser("inventory")
    inventory.add_argument("--questions", required=True, type=Path)
    inventory.add_argument("--source", required=True, type=Path)
    inventory.add_argument("--image-root", required=True, type=Path)
    inventory.add_argument("--out", required=True, type=Path)
    validate = sub.add_parser("validate")
    validate.add_argument("--manifest", required=True, type=Path)
    validate.add_argument("--questions", required=True, type=Path)
    validate.add_argument("--source", required=True, type=Path)
    validate.add_argument("--image-root", required=True, type=Path)
    sheets = sub.add_parser("contact-sheets")
    sheets.add_argument("--manifest", required=True, type=Path)
    sheets.add_argument("--image-root", required=True, type=Path)
    sheets.add_argument("--out", required=True, type=Path)
    sheets.add_argument("--pairs", action="store_true")
    sheets.add_argument("--pair-id", action="append")
    args = parser.parse_args(argv)
    try:
        if args.command == "inventory":
            questions = json.loads(args.questions.read_text(encoding="utf-8"))
            manifest = prepare_manifest(questions, source_digest(args.source), args.image_root)
            export_manifest(manifest, args.out)
            print(json.dumps({"status": "REVIEW_REQUIRED", "files": len(manifest["inventory"]["files"]),
                              "pairs": len(manifest["pairs"]), "inventory_digest": manifest["inventory_digest"]}))
        elif args.command == "validate":
            manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
            errors = validate_manifest(manifest, json.loads(args.questions.read_text(encoding="utf-8")),
                                       source_digest(args.source), args.image_root)
            print(json.dumps({"status": "CHECK" if errors else "OK", "errors": errors}, ensure_ascii=False))
            return 1 if errors else 0
        else:
            manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
            index = contact_sheets(manifest, args.image_root, args.out, pairs=args.pairs, pair_ids=args.pair_id)
            print(json.dumps({"sheets": len(index["sheets"]), "status": "REVIEW_REQUIRED"}))
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as exc:
        print(json.dumps({"status": "CHECK", "errors": [str(exc)]}, ensure_ascii=False))
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
