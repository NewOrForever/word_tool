# Header Footer Alignment And Loading Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add independent left, center, and right alignment controls for headers and footers, safely remove obsolete header/footer image data, and rebuild verified Windows packages

**Architecture:** Keep the existing direct DOCX package editor and extend its public processing call with explicit alignment values. Parse only the relationship files owned by replaced headers and footers, remove their stale links, and delete an old internal image only when no remaining package relationship references it. Pass the two independent alignment selections from the Tkinter interface into the existing background worker

**Tech Stack:** Python 3.14, standard-library zipfile and ElementTree, Tkinter/ttk, unittest, PyInstaller, Microsoft Word automation for final opening checks

## Global Constraints

- Header and footer alignment are selected independently from left, center, and right
- Header defaults to left and footer defaults to right to preserve current behavior
- Body content and body images must remain unchanged
- Existing drag-and-drop, recursive folder processing, and in-place replacement remain available
- A failed file must not replace the original with a damaged output
- Do not create a Git commit

---

### Task 1: Core Alignment Support

**Files:**
- Modify: `test_footer_update.py`
- Modify: `remove_word_headers_footers.py`

**Interfaces:**
- Produces: `build_header_footer_xml(kind: str, text_value: str, alignment: str) -> bytes`
- Produces: `process_docx(..., header_alignment: str = "left", footer_alignment: str = "right") -> ProcessResult`

- [ ] **Step 1: Write failing alignment tests**

Add tests that process fresh sample documents with header and footer alignments `left`, `center`, and `right`, parse the resulting XML, and assert the exact `w:jc/@w:val` value. Add one test using `header_alignment="center"` and `footer_alignment="left"` in the same call

```python
result = process_docx(
    source,
    "set_both",
    None,
    True,
    header_text="页眉",
    footer_text="页脚",
    header_alignment="center",
    footer_alignment="left",
)
self.assert_part_alignment(source, "word/header1.xml", "center")
self.assert_part_alignment(source, "word/footer1.xml", "left")
```

- [ ] **Step 2: Run tests and verify the expected failure**

Run: `py -m unittest test_footer_update.py`

Expected: FAIL because `process_docx` does not accept the two alignment arguments and header XML has no explicit alignment

- [ ] **Step 3: Implement the alignment parameters**

Add an allowed-value set and validate alignment before modifying a file. Always write an explicit paragraph alignment for both headers and footers

```python
ALIGNMENT_VALUES = {"left", "center", "right"}

def build_header_footer_xml(kind: str, text_value: str, alignment: str) -> bytes:
    if alignment not in ALIGNMENT_VALUES:
        raise ValueError(f"不支持的对齐方式：{alignment}")
    wrapper = ET.Element(f"{{{WORD_NS}}}{HEADER_FOOTER_CONFIGS[kind]['xml_tag']}")
    paragraph = ET.SubElement(wrapper, f"{{{WORD_NS}}}p")
    paragraph_properties = ET.SubElement(paragraph, f"{{{WORD_NS}}}pPr")
    ET.SubElement(paragraph_properties, f"{{{WORD_NS}}}jc", {f"{{{WORD_NS}}}val": alignment})
```

Pass the selected value through `prepare_header_footer_update` and use `left` and `right` as the public defaults

- [ ] **Step 4: Run the focused and full tests**

Run: `py -m unittest test_footer_update.py`

Expected: all alignment and existing update tests pass

Run: `py -m unittest discover`

Expected: all tests pass

### Task 2: Safe Cleanup Of Replaced Images And Links

**Files:**
- Modify: `test_footer_update.py`
- Modify: `remove_word_headers_footers.py`

**Interfaces:**
- Produces: `resolve_relationship_target(relationship_path: str, target: str) -> str`
- Produces: `collect_obsolete_part_files(source_zip: zipfile.ZipFile, updated_parts: set[str]) -> set[str]`

- [ ] **Step 1: Write failing cleanup tests**

Extend the sample package with `word/media/old-footer.png` referenced only by `word/_rels/footer1.xml.rels`. Assert that replacing the footer removes both the relationship file and the old image. Add a second package where `word/_rels/document.xml.rels` also references the same image and assert that the shared image remains

```python
with zipfile.ZipFile(source) as zf:
    self.assertNotIn("word/_rels/footer1.xml.rels", zf.namelist())
    self.assertNotIn("word/media/old-footer.png", zf.namelist())
```

- [ ] **Step 2: Run tests and verify the expected failure**

Run: `py -m unittest test_footer_update.py`

Expected: FAIL because the obsolete internal image is still copied into the output package

- [ ] **Step 3: Implement relationship-aware cleanup**

For each updated header/footer part, parse its own relationship file and collect internal image targets. Parse all relationship files that will remain, resolve their targets relative to their owning part, and retain every candidate image still referenced elsewhere. Return only the old relationship files and unreferenced image parts for removal

```python
def collect_obsolete_part_files(source_zip: zipfile.ZipFile, updated_parts: set[str]) -> set[str]:
    removed_relationships = {build_part_relationship_path(part) for part in updated_parts}
    candidate_images = collect_internal_image_targets(source_zip, removed_relationships)
    remaining_targets = collect_remaining_internal_targets(source_zip, removed_relationships)
    return removed_relationships | (candidate_images - remaining_targets)
```

Malformed optional relationship files must be ignored for image deletion so the cleaner never guesses that an image is unused

- [ ] **Step 4: Run focused and full tests**

Run: `py -m unittest test_footer_update.py`

Expected: old footer-only images are removed and shared images remain

Run: `py -m unittest discover`

Expected: all tests pass

### Task 3: Independent GUI Controls And Packaged Self-Test

**Files:**
- Modify: `word_header_footer_gui.py`

**Interfaces:**
- Consumes: `process_docx(..., header_alignment, footer_alignment)` from Task 1
- Produces: `header_alignment_var` and `footer_alignment_var` selections mapped from Chinese labels to `left`, `center`, and `right`

- [ ] **Step 1: Strengthen the self-test before changing the interface**

Update `run_self_test` to process a sample with a centered header and left-aligned footer. Assert both values and assert the old relationship and old footer-only image are gone

```python
update_result = process_docx(
    update_source,
    "set_both",
    None,
    True,
    header_text="测试修改页眉",
    footer_text="测试修改页脚",
    header_alignment="center",
    footer_alignment="left",
)
```

- [ ] **Step 2: Run source self-test and verify the expected failure**

Run: `py word_header_footer_gui.py --self-test`

Expected: FAIL until the new assertions and sample cleanup are supported

- [ ] **Step 3: Add two read-only dropdowns**

Import `Combobox`, add the two annotated `StringVar` properties, initialize them to `左边` and `右边`, and place one fixed-width selector after each text input. Convert the selected labels through one shared mapping before starting the worker

```python
ALIGNMENT_LABELS = {"左边": "left", "居中": "center", "右边": "right"}
self.header_alignment_var = StringVar(value="左边")
self.footer_alignment_var = StringVar(value="右边")
```

Pass both values through `_run_processing`, include them in the processing log, and forward them to every `process_docx` call

- [ ] **Step 4: Run source checks**

Run: `py word_header_footer_gui.py --self-test`

Expected: `SELF TEST PASS`

Run: `py -m unittest discover`

Expected: all tests pass

### Task 4: Rebuild And Real-Document Verification

**Files:**
- Modify: `WordHeaderFooterCleaner.exe`
- Modify: `WordHeaderFooterCleanerSetup.exe`

**Interfaces:**
- Consumes: verified source code from Tasks 1-3
- Produces: two runnable Windows executables in the workspace root

- [ ] **Step 1: Build the main executable**

Run PyInstaller with `--hidden-import platform`, exclude the installed `tkinterdnd2` module, and bundle the known working `tkinterdnd2` package data. Copy the successful output to `D:\videcode\word\WordHeaderFooterCleaner.exe`

- [ ] **Step 2: Build the setup executable**

Bundle the rebuilt main executable into `WordHeaderFooterCleanerSetup.exe` and copy the successful output to the workspace root

- [ ] **Step 3: Verify both packaged self-tests**

Run each windowed executable with `Start-Process -Wait -PassThru -NoNewWindow`, redirect standard output, and assert exit code `0`

Expected: `SELF TEST PASS` and `SETUP SELF TEST PASS`

- [ ] **Step 4: Process all 128 attached documents on copies**

Copy every `.docx` under `C:\Users\Admin\Desktop\aaa - 副本\北师大版\合集\2026test` to a temporary verification directory. Process all copies with different header/footer alignments and verify ZIP integrity, selected text and alignment, absence of external `file:///` links, and absence of unreferenced old footer images

Expected: `BATCH_OK 128`

- [ ] **Step 5: Open representative processed documents in Word**

Open small, medium, large, first, middle, and last processed files through Word automation, close each cleanly, record opening time, and confirm no repair or link prompt appears

- [ ] **Step 6: Final clean verification and cleanup**

Run: `py -m unittest discover`

Expected: all tests pass with zero failures

Remove temporary build, verification, and Word automation files. Confirm no `WINWORD.EXE` automation process remains and report only the intended modified source, tests, design documents, and two packages. Do not commit
