"""
Helper script to read a Word document as numbered runs.
Displays all runs with their indices, making it easy to reference specific text for commenting.
"""
from docx import Document
from lxml import etree
import sys
from typing import Optional

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def iter_all_runs(para):
    """
    Yield (run_element, is_hyperlink) for every <w:r> in the paragraph,
    in document order, including runs nested inside <w:hyperlink>.
    """
    for child in para._element:
        tag = child.tag.split('}')[-1]
        if tag == 'r':
            yield child, False
        elif tag == 'hyperlink':
            for inner in child.findall(f'{W}r'):
                yield inner, True


def read_document_runs(docx_path: str) -> dict:
    """
    Read document and return all runs numbered for easy reference.

    Returns:
        dict with keys:
            - 'runs': list of dicts with run info (para_idx, run_idx, text, bold, italic, is_hyperlink)
            - 'comments': list of existing comments with their anchored runs
            - 'total_runs': total number of runs
            - 'total_chars': total character count
    """
    doc = Document(docx_path)

    all_runs = []
    run_counter = 0
    total_chars = 0

    # Collect all runs from all paragraphs, including hyperlink runs
    for para_idx, para in enumerate(doc.paragraphs):
        for run_elem, is_hyperlink in iter_all_runs(para):
            rPr = run_elem.find(f'{W}rPr')
            text = run_elem.findtext(f'{W}t', default='')
            bold = rPr is not None and rPr.find(f'{W}b') is not None if rPr is not None else False
            italic = rPr is not None and rPr.find(f'{W}i') is not None if rPr is not None else False
            run_info = {
                'global_run_id': run_counter,
                'para_idx': para_idx,
                'text': text,
                'bold': bold,
                'italic': italic,
                'is_hyperlink': is_hyperlink,
            }
            all_runs.append(run_info)
            total_chars += len(text)

            run_counter += 1
    
    # Read existing comments and find which paragraphs they're in
    existing_comments = []
    if hasattr(doc, 'comments') and doc.comments is not None:
        try:
            comment_locations = {}  # comment_id -> para_idx
            
            for para_idx, para in enumerate(doc.paragraphs):
                para_elem = para._element
                # Look for commentRangeStart or commentReference in this paragraph
                for elem in para_elem.iter():
                    if elem.tag == f'{W}commentRangeStart' or elem.tag == f'{W}commentReference':
                        comment_id = elem.get(f'{W}id')
                        if comment_id is not None:
                            comment_locations[comment_id] = para_idx
            
            for i, comment in enumerate(doc.comments):
                comment_id_str = str(getattr(comment, 'id', i))
                para_idx = comment_locations.get(comment_id_str)
                
                comment_info = {
                    'id': getattr(comment, 'id', i),
                    'author': getattr(comment, 'author', 'Unknown'),
                    'text': getattr(comment, 'text', ''),
                    'para_idx': para_idx,
                }
                existing_comments.append(comment_info)
        except Exception as e:
            print(f"⚠️  Warning: Could not fully read all existing comments: {e}")
    
    return {
        'runs': all_runs,
        'comments': existing_comments,
        'total_runs': len(all_runs),
        'total_chars': total_chars,
    }


def display_document_runs(docx_path: str) -> None:
    """Display document runs in a format easy for Claude to read."""
    
    print("=" * 80)
    print(f"READING: {docx_path}")
    print("=" * 80)
    
    result = read_document_runs(docx_path)
    
    print(f"\n📊 DOCUMENT STATISTICS:")
    print(f"   Total runs: {result['total_runs']}")
    print(f"   Total characters: {result['total_chars']:,}")
    print(f"   Existing comments: {len(result['comments'])}")
    
    if result['comments']:
        print(f"\n💬 EXISTING COMMENTS:")
        for comment in result['comments']:
            author = comment['author']
            text = comment['text']
            para_info = f" (Paragraph {comment['para_idx']})" if comment.get('para_idx') is not None else ""
            print(f"   [{author}]{para_info} {text}")
    
    print(f"\n📖 ALL RUNS (numbered for easy reference):")
    print("=" * 80)
    
    current_para = -1
    for run_info in result['runs']:
        # Print paragraph separator when we move to a new paragraph
        if run_info['para_idx'] != current_para:
            current_para = run_info['para_idx']
            print(f"\n--- Paragraph {current_para} ---")
        
        # Format the run display
        run_id = run_info['global_run_id']
        text = run_info['text']
        
        # Show formatting indicators
        formatting = []
        if run_info.get('is_hyperlink'):
            formatting.append('LINK')
        if run_info['bold']:
            formatting.append('BOLD')
        if run_info['italic']:
            formatting.append('ITALIC')
        format_str = f" [{', '.join(formatting)}]" if formatting else ""
        
        # Display the run
        # Show empty runs as [EMPTY]
        if not text:
            print(f"[Run {run_id}] [EMPTY]{format_str}")
        else:
            # Show full run text - do NOT truncate, as commenting requires reading every word
            display_text = text
            print(f"[Run {run_id}] {display_text}{format_str}")
    
    print("=" * 80)
    print(f"\n✅ Document read complete. Total runs: {result['total_runs']}")
    print(f"   (Make sure you see all runs from [Run 0] to [Run {result['total_runs'] - 1}])")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python read_document_runs.py <path_to_docx>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    display_document_runs(docx_path)
