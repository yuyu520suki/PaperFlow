"""Run import, template formatting and native integrations with an evidence log."""
from __future__ import annotations
from collections import Counter
import hashlib
import json
from pathlib import Path
import shutil

from .ingest import import_document
from .formatting import normalize_to_template
from .package import (audit_docx, expand_markers, marked_zotero_keys, navigation_index_kind, read_package,
                      validate_chapter_caption_options)
from .zotero import ZoteroClient, preference_style_id, read_document_preferences
from .word import private_word, build_from_template, refresh_fields, export_pdf, heading_list_labels
from .journal import CNS_STYLE, audit_journal, audit_rendered_object_pages

CNS = "computers-and-structures"


def _write_report(path, manifest):
    path.write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding="utf-8")


def _check_metadata(metadata):
    if metadata is not None and not isinstance(metadata, dict):
        raise ValueError("Submission metadata must be a JSON object")


def _citation_style_choice(csl_style, journal):
    if csl_style is not None:
        return csl_style, "explicit"
    if journal == CNS:
        return CNS_STYLE, "journal_default"
    return None, "document"


def _finish_word(source, output, app, *, csl_style=None, style_selection="document",
                 require_mathtype=True, chapter_migrations=None, legacy_migration=None, record):
    """Refresh a working copy, including Word revisions made since the last build."""
    from .citations import CitationError, prepare_citation_style, refresh_zotero, verify_zotero_fields_preserved

    current=source
    choice = None
    if csl_style is not None:
        current, choice = prepare_citation_style(current, output/"05-citation-style.docx", csl_style)
        record("citation_style", {"selection": style_selection, **choice})
    if require_mathtype:
        from .equations import convert_to_mathtype
        equations=output/"06-mathtype.docx"
        math_report = convert_to_mathtype(current, equations, app=app)
        if legacy_migration:
            expected = Counter(legacy_migration["mtef_sha256"])
            actual = Counter(item["mtef_sha256"] for item in math_report["math_audit"]["objects"])
            if expected - actual:
                raise RuntimeError("Word changed a migrated Equation 3 payload while inserting the source")
            count = legacy_migration["converted_legacy_equation_objects"]
            math_report["converted_legacy_equation_objects"] += count
            math_report["preserved_native_count"] -= count
            math_report["legacy_mtef_preserved"] = True
        record("mathtype", math_report)
        current=equations
    ordinary=output/"07-fields-refreshed.docx"
    record("word_fields",refresh_fields(current,ordinary,app=app))
    zotero_output=output/"07-zotero-refreshed.docx"
    zotero_report=refresh_zotero(ordinary,zotero_output,app=app)
    record("zotero",zotero_report)
    master=output/"editable-master.docx"
    candidate = master
    if chapter_migrations:
        candidate = output/"07-chapter-candidate.docx" if zotero_report["macro_invoked"] else zotero_output
    if zotero_report["macro_invoked"]:
        # A longer bibliography or citation can shift headings to other pages.
        record("word_fields_after_zotero",refresh_fields(zotero_output,candidate,app=app))
        record("zotero_after_repagination",verify_zotero_fields_preserved(zotero_output,candidate))
    elif candidate != zotero_output:
        with zotero_output.open("rb") as source_stream, candidate.open("xb") as candidate_stream:
            shutil.copyfileobj(source_stream,candidate_stream)
    citation_fields = zotero_report["after"]["citation_fields"]
    bibliography_fields = zotero_report["after"]["bibliography_fields"]
    saved_style = None
    if choice or citation_fields or bibliography_fields:
        saved_style = preference_style_id(read_document_preferences(candidate))
    if choice and saved_style != choice["id"]:
        raise CitationError("The saved Word master does not contain the selected Zotero style")
    style = {
        "selection": style_selection,
        "requested_style_id": choice["id"] if choice else None,
        "previous_style_id": choice["previous_style_id"] if choice else saved_style,
        "saved_style_id": saved_style,
        "preferences_changed": choice["preferences_changed"] if choice else False,
        "native_refresh_verified": zotero_report["zotero_native_refresh_verified"],
        "citation_fields": citation_fields, "bibliography_fields": bibliography_fields,
    }
    if require_mathtype:
        from .mathtype import audit_docx_math
        before, after = (audit_docx_math(path) for path in (current, candidate))
        if (not after["all_mathtype_objects_valid"] or after["omml_equations"] or after["legacy_equation_objects"] or
                [o["mtef_sha256"] for o in before["objects"]] !=
                [o["mtef_sha256"] for o in after["objects"]]):
            raise RuntimeError("Word field refresh changed the MathType formula data")
        record("math_after_refresh", {"native_equations": after["native_mathtype_objects"],
                                     "mtef_payloads_preserved": True})
    if chapter_migrations:
        from .chapter_captions import validate_chapter_migration_values

        record("chapter_migration_values", validate_chapter_migration_values(candidate, chapter_migrations))
        with candidate.open("rb") as source_stream, master.open("xb") as master_stream:
            shutil.copyfileobj(source_stream,master_stream)
    from .table_native import audit_table_layout
    record("table_layout", audit_table_layout(master, app=app))
    record("pdf",export_pdf(master,output/"editable-master.pdf",app=app))
    return master, style


def _submission_copy(master, output, app, journal, record):
    from .citations import create_submission_copy
    base=output/"08-unlinked-submission.docx"
    record("submission_unlink",create_submission_copy(master,base))
    placement=None
    if journal == CNS:
        from .placement import append_end_objects
        arranged=output/"09-with-end-figures-tables.docx"
        placement=append_end_objects(base,arranged)
        record("journal_object_placement",placement)
        base=arranged
    submission=output/"submission-copy.docx"
    record("submission_word_fields",refresh_fields(base,submission,app=app))
    record("submission_pdf",export_pdf(submission,output/"submission-copy.pdf",app=app))
    return submission, placement


def _audit_outputs(manifest, master, submission):
    for key, path in (("structural_audit", master), ("submission_audit", submission)):
        if path is None:
            continue
        audit = audit_docx(path)
        manifest[key] = audit
        if not audit["structural_pass"]:
            raise RuntimeError(f"Word output has unresolved fields: {audit['errors'] + audit['pending_markers']}")


def _citation_items(document, csl_style):
    keys = marked_zotero_keys(document)
    if not keys:
        return {}
    if not csl_style:
        raise ValueError("New Zotero citation markers need --csl-style or a journal profile")
    client = ZoteroClient()
    return {key: client.get_item(key) for key in sorted(keys)}


def _expand_template_fields(source, output, items, normalization, chapter_level, chapter_separator):
    if chapter_level is not None and chapter_level not in normalization["numbered_heading_levels"]:
        raise ValueError(f"Chapter caption numbering requires the template's numbered Word Heading {chapter_level} style")
    styles = normalization["role_style_ids"]
    return expand_markers(source, output, items, caption_style=styles.get("caption", "Caption"),
                          equation_style=styles.get("equation", "Equation"),
                          caption_chapter_level=chapter_level, caption_chapter_separator=chapter_separator)


def _assemble_layout(source, template, assignments, output, app, *, csl_style, table_policy, table_style, role_style_map,
                     chapter_level, chapter_separator, record):
    from .layout import split_source

    staging = output / "layout-parts"
    chunks = split_source(source, assignments, staging)
    items = _citation_items(source, csl_style)
    current = template
    built = output / "04-template-built.docx"
    review_required = False
    placements = []
    chapter_bookmarks = []
    for index, (assignment, chunk) in enumerate(zip(assignments, chunks, strict=True), 1):
        slot = assignment["slot"]
        normalized = staging / f"part-{index:03d}-styles.docx"
        roles = {**(role_style_map or {}), **assignment.get("role_style_map", {})}
        normalization = normalize_to_template(chunk, template, normalized, slot=slot, table_policy=table_policy,
                                              table_style=table_style, role_style_map=roles)
        record("layout_normalize", {"slot": slot, **normalization})
        review_required |= bool(normalization["unresolved"])
        expanded = staging / f"part-{index:03d}-fields.docx"
        expansion = _expand_template_fields(normalized, expanded, items, normalization, chapter_level, chapter_separator)
        record("layout_semantic_fields", {"slot": slot, **expansion})
        chapter_bookmarks.extend(expansion["chapter_bookmarks"])
        destination = built if index == len(assignments) else staging / f"template-through-{index:03d}.docx"
        insertion = build_from_template(current, expanded, destination, slot=slot, app=app, preserve_section_stories=True)
        matched = insertion["before"]["sections"] == insertion["after"]["sections"]
        placement = {"slot": slot, "source_blocks": [assignment["start"], assignment["end"]],
                     "section_geometry_matches_template": matched, "output": str(destination),
                     "section_stories_match_template": insertion["section_stories_match_template"],
                     "inserted_range": insertion["inserted_range"]}
        record("layout_insert", placement)
        if not matched:
            raise RuntimeError(f"Word changed the template's section geometry while filling {slot!r}")
        placements.append(placement)
        current = destination
    record("template_insert", {"slots": placements, "section_geometry_matches_template": True, "output": str(built)})
    return built, review_required, chapter_bookmarks


def run_pipeline(source, template, output_dir, *, csl_style=None, metadata=None, require_mathtype=True,
                 slot="{{PAPERFLOW_BODY}}", journal=None, submission_copy=False, table_policy="preserve",
                 ocr_mode="auto", ocr_language=None, role_style_map=None, draft_plan=None, layout_plan=None,
                 caption_chapter_level=None, caption_chapter_separator=None, citekey_plan=None, table_style=None):
    _check_metadata(metadata)
    chapter_separator = "-" if caption_chapter_separator is None else caption_chapter_separator
    if caption_chapter_level is not None or draft_plan is None:
        validate_chapter_caption_options(caption_chapter_level, chapter_separator)
    source,template=Path(source).resolve(strict=True),Path(template).resolve(strict=True)
    assignments = None
    if layout_plan is not None:
        if slot != "{{PAPERFLOW_BODY}}":
            raise ValueError("A layout plan defines the insertion positions. Do not also set a single slot.")
        from .layout import validate_layout
        assignments = validate_layout(source, template, layout_plan)
    output=Path(output_dir).resolve(); output.mkdir(parents=True,exist_ok=False)
    if journal not in {None,CNS}: raise ValueError(f"Unknown journal profile: {journal}")
    csl_style, style_selection = _citation_style_choice(csl_style, journal)
    manifest={"source":str(source),"template":str(template),"journal":journal,"source_sha256":hashlib.sha256(source.read_bytes()).hexdigest(),
              "template_sha256":hashlib.sha256(template.read_bytes()).hexdigest(),"steps":[],"native_build_finished":False}
    def record(stage,evidence):
        manifest["steps"].append({"stage":stage,"evidence":evidence})
        if stage == "table_layout":
            manifest["layout_needs_review"] = bool(evidence["unresolved"])
        _write_report(output/"workflow.json",manifest)
    try:
        prepared_ast = None
        if citekey_plan is not None:
            if draft_plan is not None or layout_plan is not None:
                raise ValueError("Apply the citekey plan to the original Markdown or LaTeX source before any DOCX layout plan")
            if not csl_style:
                raise ValueError("A citekey plan needs --csl-style or a journal profile")
            from .citekeys import prepare_citekey_import, verify_imported_citekey_fields

            prepared_ast, citation_ids, citekey_evidence = prepare_citekey_import(source, citekey_plan)
            if citekey_evidence["source_sha256"] != manifest["source_sha256"]:
                raise ValueError("The source changed while the citekey plan was being applied")
            record("citekey_plan", citekey_evidence)
        else:
            from .citekeys import require_plan_for_structured_citations

            require_plan_for_structured_citations(source)
        with private_word() as app:
            imported=output/"01-imported.docx"
            imported_report=import_document(source,imported,template=template,app=app,ocr_mode=ocr_mode,
                                            ocr_language=ocr_language,pandoc_ast=prepared_ast,reuse_template_indexes=True,
                                            role_style_map=role_style_map)
            if source.suffix.lower() == ".tex" and hashlib.sha256(source.read_bytes()).hexdigest() != manifest["source_sha256"]:
                raise ValueError("The LaTeX source changed during import")
            if prepared_ast is not None:
                imported_report["linked_citekey_fields"] = verify_imported_citekey_fields(imported, citation_ids)
            record("import",imported_report)
            latex_references = imported_report.get("latex_references")
            manifest["input_needs_semantic_review"]=bool(imported_report.get("semantic_review_required"))
            migrated_chapter_captions = []
            if draft_plan is not None:
                if source.suffix.lower() != ".docx":
                    raise ValueError("A draft plan must refer to a reviewed DOCX; import other formats first")
                from .semantics import apply_draft_plan
                planned=output/"01-draft-plan-applied.docx"
                plan_report=apply_draft_plan(imported,planned,draft_plan)
                record("draft_plan",plan_report)
                policy = plan_report["chapter_caption_policy"]
                if policy is not None:
                    if (caption_chapter_level not in (None, policy["level"])
                            or caption_chapter_separator not in (None, policy["separator"])):
                        raise ValueError("Run's chapter caption options differ from the reviewed draft plan")
                    caption_chapter_level, chapter_separator = policy["level"], policy["separator"]
                    migrated_chapter_captions = plan_report["chapter_captions"]
                manifest["input_needs_semantic_review"] |= bool(plan_report["issues"])
                imported=planned
            validate_chapter_caption_options(caption_chapter_level, chapter_separator)
            legacy_migration = None
            if require_mathtype:
                from .mathtype import audit_docx_math

                legacy = audit_docx_math(imported)["legacy_objects"]
                if legacy:
                    from .legacy_equations import rewrap_legacy_equations

                    migrated = output / "01-migrated-equation3.docx"
                    legacy_migration, legacy_anchors = rewrap_legacy_equations(
                        read_package(imported), legacy, migrated, anchor_for_insertion=True,
                    )
                    record("legacy_equations_before_insertion", legacy_migration)
                    imported = migrated
            if assignments is not None:
                if (manifest["source_sha256"] != layout_plan["source_sha256"]
                        or manifest["template_sha256"] != layout_plan["template_sha256"]):
                    raise ValueError("The original source or template changed after layout validation")
                record("layout_plan", {"assignments": assignments, "source_and_template_hashes_verified": True})
                built, needs_review, chapter_bookmarks = _assemble_layout(
                    imported, template, assignments, output, app, csl_style=csl_style,
                    table_policy=table_policy, table_style=table_style, role_style_map=role_style_map,
                    chapter_level=caption_chapter_level,
                    chapter_separator=chapter_separator, record=record,
                )
                manifest["input_needs_semantic_review"] |= needs_review
            else:
                normalized=output/"02-template-styles.docx"
                normalization=normalize_to_template(imported,template,normalized,slot=slot,table_policy=table_policy,
                                                    table_style=table_style,role_style_map=role_style_map)
                record("normalize",normalization)
                manifest["input_needs_semantic_review"] |= bool(normalization["unresolved"])
                expanded=output/"03-dynamic-fields.docx"
                expansion = _expand_template_fields(normalized, expanded, _citation_items(normalized, csl_style),
                                                     normalization, caption_chapter_level, chapter_separator)
                record("semantic_fields", expansion)
                chapter_bookmarks = expansion["chapter_bookmarks"]
                built=output/"04-template-built.docx"
                insertion=build_from_template(template,expanded,built,slot=slot,app=app,preserve_section_stories=True)
                insertion["section_geometry_matches_template"]=insertion["before"]["sections"]==insertion["after"]["sections"]
                record("template_insert",insertion)
                manifest["input_needs_semantic_review"] |= not insertion["section_geometry_matches_template"]
            if legacy_migration:
                from .legacy_equations import restore_legacy_payloads

                restored = output / "04-legacy-payloads-restored.docx"
                record("legacy_equations_after_insertion", restore_legacy_payloads(built, restored, legacy_anchors))
                built = restored
            chapter_bookmarks.extend(entry["bookmark"] for entry in migrated_chapter_captions)
            if chapter_bookmarks:
                from .chapter_captions import validate_chapter_captions

                record("chapter_caption_headings", validate_chapter_captions(built, caption_chapter_level, chapter_bookmarks))
            master, style = _finish_word(built,output,app,csl_style=csl_style,style_selection=style_selection,
                                        require_mathtype=require_mathtype,chapter_migrations=migrated_chapter_captions,
                                        legacy_migration=legacy_migration,record=record)
            manifest["citation_style"] = style
            if latex_references and latex_references["targets"]:
                from .chapter_captions import verify_generated_reference_values

                verification = verify_generated_reference_values(
                    master, latex_references["targets"], latex_references["references"],
                )
                numeric = {ref["bookmark"] for ref in latex_references["references"] if ref["command"] != "pageref"}
                headings = [target["bookmark"] for target in latex_references["targets"]
                            if target["kind"] == "Heading" and target["bookmark"] in numeric]
                verification["word_heading_list_labels"] = heading_list_labels(master, headings, app=app)
                record("latex_references_after_word", verification)
            if table_style is not None:
                from .table_style import verify_template_table_style

                applications = [step["evidence"]["table_style"] for step in manifest["steps"]
                                if step["stage"] in {"normalize", "layout_normalize"}]
                record("table_style_after_word", verify_template_table_style(template, master, applications))
            submission, placement=_submission_copy(master,output,app,journal,record) if submission_copy else (None,None)
        _audit_outputs(manifest,master,submission)
        if latex_references and latex_references["indexes"]:
            observed = {}
            for key in ("structural_audit", "submission_audit"):
                if key not in manifest:
                    continue
                kinds = Counter(navigation_index_kind(field["instruction"]) for field in manifest[key]["fields"]
                                if field["part"] == "word/document.xml")
                for entry in latex_references["indexes"]:
                    count = kinds[entry["kind"]]
                    if count < 1 or (entry["mode"] == "source" and count != 1):
                        raise ValueError(f"Word saved {count} {entry['kind']} indexes in {key} after LaTeX import")
                observed[key] = {entry["kind"]: kinds[entry["kind"]] for entry in latex_references["indexes"]}
            record("latex_indexes_after_word", observed)
        from .template import section_story_signature
        expected_stories = section_story_signature(template)
        if any(section_story_signature(path) != expected_stories for path in (master, submission) if path is not None):
            raise RuntimeError("The final Word file's headers, footers, or page numbering differ from the template")
        record("template_furniture_after_refresh", {"section_stories_match_template": True})
        artwork=None
        if placement is not None:
            pages=audit_rendered_object_pages(output/"editable-master.pdf",output/"submission-copy.pdf",placement)
            record("journal_object_pages",pages)
            from .artwork import stage_artwork
            artwork=stage_artwork(placement,metadata,source.parent,output/"artwork")
            record("journal_artwork",artwork)
        else: pages=None
        if journal == CNS:
            manifest["journal_audit"]=audit_journal(submission or master,metadata,placement=pages,artwork=artwork,
                                                    citation_style=style)
        manifest["source_unchanged"]=manifest["source_sha256"]==hashlib.sha256(source.read_bytes()).hexdigest()
        manifest["template_unchanged"]=manifest["template_sha256"]==hashlib.sha256(template.read_bytes()).hexdigest()
        # A successful native build is not proof of journal compliance, visual
        # fidelity, semantic input coverage, or readiness for every user template.
        manifest["native_build_finished"]=True
        manifest["visual_reviewed"]=False
        record("finished",{"editable_master":str(master)})
        return manifest
    except Exception as exc:
        manifest["failure"]={"type":type(exc).__name__,"message":str(exc)}
        record("failed",manifest["failure"])
        raise


def revise_master(source, output_dir, *, csl_style=None, journal=None, submission_copy=False,
                  metadata=None, artwork_root=None):
    """Finalize a Word-edited master directly; do not reconstruct it from an earlier source."""
    _check_metadata(metadata)
    source=Path(source).resolve(strict=True)
    if source.suffix.lower() != ".docx": raise ValueError("Revision requires an editable DOCX master")
    if journal not in {None,CNS}: raise ValueError(f"Unknown journal profile: {journal}")
    csl_style, style_selection = _citation_style_choice(csl_style, journal)
    output=Path(output_dir).resolve(); output.mkdir(parents=True,exist_ok=False)
    original_hash=hashlib.sha256(source.read_bytes()).hexdigest()
    manifest={"source":str(source),"source_sha256":original_hash,"journal":journal,"mode":"revision","steps":[]}
    def record(stage,evidence):
        manifest["steps"].append({"stage":stage,"evidence":evidence})
        if stage == "table_layout":
            manifest["layout_needs_review"] = bool(evidence["unresolved"])
        _write_report(output/"workflow.json",manifest)
    try:
        with private_word() as app:
            master, style = _finish_word(source,output,app,csl_style=csl_style,style_selection=style_selection,record=record)
            manifest["citation_style"] = style
            submission, placement=_submission_copy(master,output,app,journal,record) if submission_copy else (None,None)
        _audit_outputs(manifest,master,submission)
        pages=artwork=None
        if placement is not None:
            pages=audit_rendered_object_pages(output/"editable-master.pdf",output/"submission-copy.pdf",placement)
            record("journal_object_pages",pages)
            from .artwork import stage_artwork
            artwork=stage_artwork(placement,metadata,artwork_root or source.parent,output/"artwork")
            record("journal_artwork",artwork)
        if journal == CNS:
            manifest["journal_audit"]=audit_journal(submission or master,metadata,placement=pages,artwork=artwork,
                                                    citation_style=style)
        manifest["source_unchanged"]=hashlib.sha256(source.read_bytes()).hexdigest()==original_hash
        manifest["native_build_finished"]=True
        record("finished",{"editable_master":str(master),"source_docx_edited_in_place":False})
        return manifest
    except Exception as exc:
        manifest["failure"]={"type":type(exc).__name__,"message":str(exc)}
        record("failed",manifest["failure"])
        raise
