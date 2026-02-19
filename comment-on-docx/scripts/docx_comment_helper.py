"""
Helper functions for adding comments to Word documents.
Provides a unified interface that handles run splitting automatically.
"""
from docx import Document
from docx.text.run import Run
from docx.oxml import OxmlElement
from lxml import etree
from datetime import datetime
from pathlib import Path
from typing import Union, List, Optional
import copy as python_copy
from docx.text.paragraph import Paragraph as ParagraphCls
from docx.table import Table as TableCls

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def _iter_all_runs(para):
    """
    Yield (run_element, is_hyperlink) for every <w:r> in the paragraph,
    in document order, including runs nested inside <w:hyperlink> and <w:ins>.
    Skips runs inside <w:del> (proposed deletions / track changes).
    This shows the "all suggestions accepted" version of the document.

    Note: the reader's iter_all_runs yields a 3rd element (hyperlink_url).
    The commenter doesn't need URLs, so this version yields 2-tuples.
    Both must iterate runs in the same order to keep global IDs in sync.
    """
    def _yield_runs(container):
        for child in container:
            tag = child.tag.split('}')[-1]
            if tag == 'r':
                yield child, False
            elif tag == 'hyperlink':
                for inner in child.findall(f'{W}r'):
                    yield inner, True
            elif tag == 'ins':
                # Proposed insertions — recurse to pick up runs and hyperlinks
                yield from _yield_runs(child)
            # 'del' is implicitly skipped (proposed deletions)

    yield from _yield_runs(para._element)


def _iter_document_paragraphs(doc):
    """
    Yield (Paragraph, table_info) for every paragraph in the document body,
    in document order, including paragraphs inside table cells.

    table_info is None for body paragraphs, or a dict with row/col/dimensions
    for table cell paragraphs.
    """
    body = doc.element.body
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            yield ParagraphCls(child, body), None
        elif tag == 'tbl':
            tbl = TableCls(child, body)
            num_rows = len(tbl.rows)
            num_cols = len(tbl.columns)
            for row_idx, row in enumerate(tbl.rows):
                seen_tc = set()
                for col_idx, cell in enumerate(row.cells):
                    # Skip duplicate cells from merged cells
                    tc_id = id(cell._tc)
                    if tc_id in seen_tc:
                        continue
                    seen_tc.add(tc_id)
                    for para in cell.paragraphs:
                        yield para, {
                            'row': row_idx,
                            'col': col_idx,
                            'num_rows': num_rows,
                            'num_cols': num_cols,
                        }


def find_run_by_global_id(doc: Document, global_run_id: int):
    """
    Find a run by its global ID (sequential numbering across all paragraphs,
    including runs inside hyperlinks and table cells).

    Returns:
        (paragraph, run_element, is_hyperlink) or (None, None, None) if not found
    """
    run_counter = 0
    for para, _table_info in _iter_document_paragraphs(doc):
        for run_elem, is_hyp in _iter_all_runs(para):
            if run_counter == global_run_id:
                return para, run_elem, is_hyp
            run_counter += 1
    return None, None, None


def split_run_at_text(run_element, subset_text: str):
    """
    Split a run element to isolate specific text. Works for both regular
    and hyperlink runs by operating on the XML element's parent directly.

    Args:
        run_element: The lxml <w:r> element to split
        subset_text: The text to isolate

    Returns:
        The new <w:r> element containing the isolated subset_text, or None if not found
    """
    text = run_element.findtext(f'{W}t', default='')

    # Find the subset text
    start_idx = text.lower().find(subset_text.lower())
    if start_idx == -1:
        print(f"❌ Text '{subset_text}' not found in run")
        return None

    end_idx = start_idx + len(subset_text)

    before_text = text[:start_idx]
    target_text = text[start_idx:end_idx]
    after_text = text[end_idx:]

    # The parent could be <w:p> (regular run) or <w:hyperlink>
    parent_element = run_element.getparent()

    # Find position of this run in the parent
    run_position = list(parent_element).index(run_element)

    # Create new run elements for target and after text
    # Use OxmlElement so they get proper CT_R proxy class (needed by doc.add_comment)
    target_run_elem = OxmlElement('w:r')
    after_run_elem = OxmlElement('w:r')

    # Copy formatting from original run
    original_rPr = run_element.find(f'{W}rPr')
    if original_rPr is not None:
        target_run_elem.append(python_copy.deepcopy(original_rPr))
        after_run_elem.append(python_copy.deepcopy(original_rPr))

    # Create text elements
    target_t = etree.SubElement(target_run_elem, f'{W}t')
    target_t.text = target_text
    target_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    after_t = etree.SubElement(after_run_elem, f'{W}t')
    after_t.text = after_text
    after_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    # Modify original run to contain only "before" text
    original_t = run_element.find(f'{W}t')
    if original_t is not None:
        original_t.text = before_text

    # Insert new runs after the original in the parent
    parent_element.insert(run_position + 1, target_run_elem)
    parent_element.insert(run_position + 2, after_run_elem)

    return target_run_elem


def add_comment(
    doc: Document,
    run_ids: Union[int, List[int]],
    text: str,
    subset_text: Optional[str] = None,
    author: str = "Claude",
    initials: str = "CS"
) -> bool:
    """
    Add a comment to specified runs, with optional text subsetting.
    
    Args:
        doc: Document object
        run_ids: Single global run ID or list of consecutive global run IDs
        text: Comment text
        subset_text: Optional - if provided, splits the run to comment on just this text
        author: Comment author name
        initials: Comment author initials
    
    Returns:
        True if successful, False otherwise
    
    Examples:
        # Comment on entire run 42
        add_comment(doc, run_ids=42, text="Good point!")
        
        # Comment on runs 10-15 (entire range)
        add_comment(doc, run_ids=[10, 11, 12, 13, 14, 15], text="This section is unclear")
        
        # Comment on just the word "strategizes" within run 42
        add_comment(doc, run_ids=42, subset_text="strategizes", text="Word choice issue")
    """
    # Normalize to list
    if isinstance(run_ids, int):
        run_ids = [run_ids]
    
    if not run_ids:
        print("❌ No run IDs provided")
        return False
    
    # If subset_text is provided, we can only work with a single run
    if subset_text and len(run_ids) > 1:
        print("❌ subset_text can only be used with a single run ID")
        return False
    
    # Find all the runs
    target_run_objs = []
    for run_id in run_ids:
        para, run_elem, is_hyp = find_run_by_global_id(doc, run_id)
        if para is None:
            print(f"❌ Run {run_id} not found in document")
            return False

        # If subset_text is provided, split the run first
        if subset_text:
            target_elem = split_run_at_text(run_elem, subset_text)
            if target_elem is None:
                return False
            run_elem = target_elem
            print(f"✅ Split run {run_id} to isolate '{subset_text}'")

        # Wrap raw element in a python-docx Run object so doc.add_comment works
        run_obj = Run(run_elem, para)
        target_run_objs.append(run_obj)

    # Filter out empty runs
    def _has_text(run_obj):
        """Check for text via Run.text or direct XML, handling raw etree elements."""
        t = run_obj.text
        if t is None:
            t = run_obj._r.findtext(f'{W}t', default='')
        return bool(t.strip())

    non_empty_runs = [r for r in target_run_objs if _has_text(r)]
    if not non_empty_runs:
        print("❌ All target runs are empty - cannot add comment")
        return False

    # Add the comment
    try:
        doc.add_comment(
            runs=non_empty_runs if len(non_empty_runs) > 1 else non_empty_runs[0],
            text=text,
            author=author,
            initials=initials
        )

        run_desc = f"run {run_ids[0]}" if len(run_ids) == 1 else f"runs {run_ids[0]}-{run_ids[-1]}"
        subset_desc = f" (on '{subset_text}')" if subset_text else ""
        print(f"✅ Added comment to {run_desc}{subset_desc}")
        return True

    except Exception as e:
        print(f"❌ Failed to add comment: {e}")
        return False


def add_comments_batch(
    doc: Document,
    comments: List[dict],
    author: str = "Claude",
    initials: str = "CS"
) -> tuple[int, int]:
    """
    Add multiple comments in a single batch, automatically handling run ID shifts
    from subset_text splits by processing comments in reverse run-ID order.

    Args:
        doc: Document object
        comments: List of comment dicts, each with keys:
            - "run_ids": int or list of ints (global run IDs from read script)
            - "text": str (comment text)
            - "subset_text": optional str (phrase to isolate within a single run)
        author: Comment author name (applied to all comments)
        initials: Comment author initials (applied to all comments)

    Returns:
        (successes, failures) tuple

    Example:
        add_comments_batch(doc, [
            {"run_ids": 28, "subset_text": "some phrase", "text": "Comment..."},
            {"run_ids": [82, 83, 84], "text": "Comment on section..."},
            {"run_ids": 155, "text": "Comment on whole run..."},
        ])
    """
    if not comments:
        print("❌ No comments provided")
        return (0, 0)

    def sort_key(c):
        ids = c["run_ids"]
        max_id = max(ids) if isinstance(ids, list) else ids
        # Process subset_text comments before non-subset at the same run ID,
        # since the split only affects runs *after* the split point
        has_subset = 1 if c.get("subset_text") else 0
        return (-max_id, -has_subset)

    sorted_comments = sorted(comments, key=sort_key)

    successes = 0
    failures = 0
    for c in sorted_comments:
        ok = add_comment(
            doc,
            run_ids=c["run_ids"],
            text=c["text"],
            subset_text=c.get("subset_text"),
            author=author,
            initials=initials,
        )
        if ok:
            successes += 1
        else:
            failures += 1

    print(f"\n📊 Batch result: {successes} succeeded, {failures} failed (out of {len(comments)})")
    return (successes, failures)


def save_with_suffix(doc: Document, original_path: str, suffix: str = "claude commented") -> str:
    """
    Save document with a suffix added to the filename.
    
    Args:
        doc: Document object to save
        original_path: Original file path
        suffix: Suffix to add (default: "claude commented")
    
    Returns:
        The output file path
    """
    original_file = Path(original_path)
    stem = original_file.stem
    output_filename = f"{stem} - {suffix}.docx"
    output_path = original_file.parent / output_filename
    
    doc.save(str(output_path))
    print(f"💾 Saved to: {output_path}")
    
    return str(output_path)


def verify_comments(doc: Document, expected_author: str = "Claude") -> int:
    """
    Verify comments were added successfully.
    
    Args:
        doc: Document to verify
        expected_author: Author name to check for
    
    Returns:
        Number of comments from expected author
    """
    total_comments = 0
    author_comments = 0

    for comment in doc.comments:
        total_comments += 1
        try:
            if comment.author == expected_author:
                author_comments += 1
        except Exception:
            pass  # skip malformed pre-existing comments (e.g. missing w:author)
    
    print(f"📊 Verification:")
    print(f"   Total comments: {total_comments}")
    print(f"   From {expected_author}: {author_comments}")
    
    return author_comments
