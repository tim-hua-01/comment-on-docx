---
name: comment-on-docx
description: Adds comments to a word doc following commenting guidelines. Use when the user asks for comments, reviews, or feedback on Microsoft Word (.docx) documents. Provides helper scripts and instructions for how to read documents, draft thoughtful comments, and add them programmatically.
---
# Add Comments to Word Documents

## Overview

Add thoughtful stylistic and content-wise comments to Microsoft Word (.docx) documents using helper scripts that abstract the complexity of the python-docx library.

## Prerequisites

- `python-docx>=1.2.0`
- Helper scripts: `scripts/read_document_runs.py` and `scripts/docx_comment_helper.py` (bundled with this skill)
- Access to the target `.docx` file

## Custom Environment Information

<!-- Users can add information about their Python environment here -->

If you're in the conda evals environment, python-docx is installed there.



/Users/timhua/anaconda3/bin/conda run -n evals

## Workflow

### Step 1: Install python-docx

Verify `python-docx>=1.2.0` is installed

### Step 2: Read the Complete Document

**CRITICAL**: You MUST read the ENTIRE document before adding comments.

Run the helper script to see all runs numbered:

```bash
python .claude/skills/comment-on-docx/scripts/read_document_runs.py "document.docx"
```

**Output format:**
```
📊 DOCUMENT STATISTICS:
   Total runs: 245
   Total characters: 25,431
   Existing comments: 3

💬 EXISTING COMMENTS:
   [Author Name] (Paragraph 5) Comment text...
   [Other Author] (Paragraph 12) Another comment...

📖 ALL RUNS (numbered for easy reference):
================================================================================

--- Paragraph 0 ---
[Run 0] This is the title text

--- Paragraph 1 ---
[Run 1] First sentence of paragraph 1. [ITALIC]
[Run 2] [EMPTY]
[Run 3] Second sentence with some bolded text.
[Run 4] more bolded text [BOLD]
[Run 5] and back to normal.

...

--- Paragraph N ---
[Run 244] Final run text
================================================================================

✅ Document read complete. Total runs: 245
   (Make sure you see all runs from [Run 0] to [Run 244])
```

**What you're looking for:**
- The total number of runs (e.g., 245)
- All runs numbered from `[Run 0]` to `[Run N-1]`
- Formatting indicators: `[BOLD]`, `[ITALIC]`, `[EMPTY]`
- Existing comments to avoid duplicating feedback

**Verification**: Confirm you see the final run matching `Total runs - 1`. If the output seems truncated or you don't see all runs, **STOP and report the issue**.

> **HARD STOP RULE**: If ANY run text ends with "..." or appears cut off, the document has NOT been fully read. You MUST stop immediately and fix the truncation issue before proceeding. Do NOT draft comments, do NOT write code, do NOT skip ahead. Commenting on text you haven't fully read produces low-quality, superficial feedback and wastes the user's time. There are no exceptions to this rule. Fix the read script or read the document another way first.

### Step 3: Draft Comments

Take time to formulate thoughtful, constructive comments. When you get to step three, open up references/commenting.md (bundled with this skill) for additional instructions. Read it in full.

### Step 4: Add Comments Using `add_comments_batch`

**Always use `add_comments_batch`** to add comments. It processes comments in reverse run-ID order so that `subset_text` splits (which insert new runs) don't shift the IDs of comments that haven't been processed yet. This means you can use the run IDs straight from the read script output without worrying about offsets.

Create a Python script using the helper functions:

```python
import sys
sys.path.insert(0, '.claude/skills/comment-on-docx')
from docx import Document
from scripts.docx_comment_helper import add_comments_batch, save_with_suffix, verify_comments

# Load document
doc = Document('your_document.docx')

# Define all comments as a list of dicts
comments = [
    {"run_ids": 42, "text": "Your comment here"},
    {"run_ids": [10, 11, 12, 13], "text": "Comment on this whole section"},
    {"run_ids": 42, "subset_text": "specific phrase", "text": "Comment on just this phrase"},
]

# Add all comments (order doesn't matter — batch handles it)
successes, failures = add_comments_batch(doc, comments)
```

#### Comment dict format

Each comment dict has these keys:

| Key | Required | Description |
|-----|----------|-------------|
| `run_ids` | Yes | Single int or list of ints (global run IDs from read script) |
| `text` | Yes | The comment text |
| `subset_text` | No | Phrase to isolate within a single run. Splits the run automatically. |

#### Comment on Entire Runs (Paragraphs or Sections)

```python
# Single run
{"run_ids": 42, "text": "Your comment here"}

# Multiple consecutive runs
{"run_ids": [10, 11, 12, 13], "text": "Comment on this whole section"}
```

#### Comment on Specific Text Within a Run

```python
# Comment on just a phrase — the batch function handles the split safely
{"run_ids": 42, "subset_text": "specific phrase", "text": "Comment on just this phrase"}
```

**Important Notes:**
- Use run IDs from the read script output (`[Run 0]`, `[Run 1]`, etc.) — no manual offset adjustment needed
- Single run: `run_ids=42`
- Multiple runs: `run_ids=[10, 11, 12, 13]`
- `subset_text` only works with single runs
- Empty runs cannot be commented on (will be skipped)
- Bolded/italicized text is often already its own run
- `subset_text` matching is case-insensitive, but beware of smart quotes/special characters in Word docs — when in doubt, comment on the whole run instead

**⚠️ Multiple Comments on Same Run:**
- If you need to add **two or more `subset_text` comments to the same run**, you must use **two passes**
- First pass: add all comments in a batch, note any failures
- Second pass: create a new script for failed comments, searching for them in the updated document
- This is because `subset_text` splits create new runs, shifting the IDs of later splits on the same run

**⚠️ Uniqueness Requirement:**
- Each `subset_text` should be **unique within the document** (or unique enough to avoid ambiguity)
- If the same phrase appears multiple times (e.g., "strategizes" in runs 10 and 50), the batch function may comment on the wrong occurrence after run ID shifts
- **Best practice:** Use longer, more specific phrases (e.g., "The AI strategizes to avoid" instead of just "strategizes")
- If you must comment on non-unique text, comment on the whole run instead of using `subset_text`

**Understanding Runs:**

A "run" is a text fragment with consistent formatting. Runs are created when:
- Formatting changes (normal → bold → normal)
- Inline styles are applied
- The document structure requires it

Example paragraph breakdown:
```
"This is normal. This is bold. This is normal again."
  ↓
[Run 0] "This is normal. "
[Run 1] "This is bold." [BOLD]
[Run 2] " This is normal again."
```

### Step 5: Save with Standard Suffix

```python
# Save to new file
output_path = save_with_suffix(
    doc,
    'your_document.docx',
    suffix="claude commented"  # Default value
)
```

This creates: `your_document - claude commented.docx` Name your comment script [short_title]_comments.py, where [short_title] is a shortened version of the doc filename.

**Never overwrite the original file.**

### Step 6: Verify Success

```python
# Verify comments were added
doc_verify = Document(output_path)
count = verify_comments(doc_verify, expected_author="Claude")

print(f"Added {count} comments successfully")
```

The verification shows:
- Total comments in document
- Number of comments from your author name
- Confirms successful addition

## Complete Example

```python
import sys
sys.path.insert(0, '.claude/skills/comment-on-docx')
from docx import Document
from scripts.docx_comment_helper import add_comments_batch, save_with_suffix, verify_comments

# Load document
doc = Document('research_post.docx')

# Define all comments — order doesn't matter, batch handles it
comments = [
    {
        "run_ids": 0,
        "text": "Style: The title is unclear. Consider: 'How Models Generalize Reward-Seeking Goals'",
    },
    {
        "run_ids": 37,
        "text": "Emphasis: Is the bolding necessary here? It may distract from the main point.",
    },
    {
        "run_ids": 42,
        "subset_text": "strategizes",
        "text": "Word choice: 'strategizes' assumes sophisticated reasoning. Justify or soften this claim.",
    },
    {
        "run_ids": [15, 16, 17, 18],
        "text": "Structure: This paragraph is dense. Consider splitting into two: one for the problem, one for the solution.",
    },
]

# Add all comments in one batch
add_comments_batch(doc, comments)

# Save with standard suffix
output = save_with_suffix(doc, 'research_post.docx')

# Verify
doc_verify = Document(output)
count = verify_comments(doc_verify, expected_author="Claude")
print(f"✅ Added {count} comments to {output}")
```

## Common Issues

### Issue 1: Cannot read entire document

**Symptom**: Output is truncated, doesn't show all runs

**Solution**: The document may be too large. Report to user that the document is too long to process safely.

### Issue 2: Run not found

**Symptom**: `❌ Run X not found in document`

**Cause**: Run ID doesn't exist (document has fewer runs than expected)

**Solution**: Re-run the read script to get current run structure

### Issue 3: Subset text not found

**Symptom**: `❌ Text 'phrase' not found in run`

**Cause**: The text doesn't exist in that run, or there's a typo

**Solution**: Check the read script output to verify which run contains the text

### Issue 4: Empty run error

**Symptom**: `❌ All target runs are empty - cannot add comment`

**Cause**: Trying to comment on runs with no text content

**Solution**: Skip empty runs or comment on adjacent runs with text

### Issue 5: Multiple subset_text on same run fails

**Symptom**: Second `subset_text` comment on the same run fails with "Text not found in run"

**Cause**: When you split a run with `subset_text`, it creates new runs and shifts IDs. If you try to add a second `subset_text` comment to the original run ID, it won't find the text because the run structure has changed.

**Example of the problem:**
```python
# Both comments target run 100, but with different subset_text
{"run_ids": 100, "subset_text": "phrase A", "text": "Comment 1"},
{"run_ids": 100, "subset_text": "phrase B", "text": "Comment 2"},  # This will fail!
```

**Solution**: Use a two-pass approach:
1. **First pass**: Add all comments in a batch, note which ones fail
2. **Second pass**: For failed comments, either:
   - Comment on a range of runs instead: `{"run_ids": [100, 101, 102], "text": "..."}`
   - Or create a second script that searches for the text and adds the comment

**Alternative**: Use longer unique `subset_text` that's less likely to need multiple comments on the same run

## Best Practices

1. **Always read the complete document first** - Don't add comments without full context
2. **Draft comments before coding** - Think through feedback before writing code
3. **Use descriptive prefixes** - Start comments with "Style:", "Content:", "Clarity:", etc.
4. **Be specific and constructive** - Explain the issue and suggest improvements
5. **Never overwrite originals** - Always use `save_with_suffix()`
6. **Verify after adding** - Check that comments were successfully added
7. **Comment at appropriate granularity**:
   - Whole paragraphs for structural/flow issues
   - Specific sentences for clarity problems
   - Individual words for terminology/word choice issues

## Limitations

- Very large documents (>100,000 words) may be slow or fail to read completely
- Comments cannot be added to headers, footers, or within existing comments
- Track changes are NOT supported (requires different approach)
- Nested or overlapping comments are not supported

## Helper Script Details

### scripts/read_document_runs.py
- Reads document and numbers all runs sequentially
- Shows formatting (bold, italic, empty)
- Lists existing comments with their paragraph locations
- Provides verification that document was read completely

### scripts/docx_comment_helper.py
- `add_comments_batch()`: **Primary interface** — takes a list of comment dicts, processes them in reverse run-ID order to avoid ID shift issues from `subset_text` splits
- `add_comment()`: Low-level interface for adding a single comment (use `add_comments_batch` instead to avoid run ID shift bugs)
- `save_with_suffix()`: Save with standard naming convention
- `verify_comments()`: Confirm comments were added successfully
- Handles run splitting automatically when using `subset_text`
- Preserves formatting when splitting runs
