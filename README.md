# comment-on-docx

A Claude Code skill that adds thoughtful, constructive comments to Microsoft Word (.docx) documents. It reads the document structure, drafts comments following editorial guidelines, and writes them back as native Word comments.

## What it does

When you ask Claude Code to review or comment on a `.docx` file, this skill:

1. Reads the document and numbers every text run for precise targeting
2. Drafts comments following guidelines for style, content, structure, and clarity
3. Adds comments programmatically using `python-docx`, targeting specific words, sentences, or paragraphs
4. Saves a new copy (never overwrites the original)

## Prerequisites

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) installed
- Python 3.10+
- `python-docx >= 1.2.0` (`pip install python-docx`)

## Installation

### Option A: Install from `.skill` file

1. Download `comment-on-docx.skill` from this repo
2. In Claude Code, run:
   ```
   /install-skill comment-on-docx.skill
   ```

### Option B: Manual install (project-level)

Copy the `comment-on-docx` directory into your project's `.claude/skills/` folder:

```bash
# From your project root
mkdir -p .claude/skills
cp -r /path/to/this/repo/comment-on-docx .claude/skills/
```

The skill will be automatically discovered next time you start Claude Code in that project.

### Option C: Manual install (global, all projects)

Copy the `comment-on-docx` directory into your personal skills folder:

```bash
mkdir -p ~/.claude/skills
cp -r /path/to/this/repo/comment-on-docx ~/.claude/skills/
```

This makes the skill available in all your Claude Code projects.

## Configuration

After installing, open `.claude/skills/comment-on-docx/SKILL.md` and edit the **Custom Environment Information** section to match your setup. For example, if you use conda:

```
/path/to/conda run -n your_env
```

Or if python-docx is in your system Python, you can clear this section.

## Usage

Once installed, just ask Claude Code to comment on a Word document:

```
comment on research_paper.docx
```

```
review the writing in "My Document.docx"
```

```
add feedback to draft.docx
```

Claude will automatically use this skill, read the full document, draft comments, and produce a new file called `<original> - claude commented.docx`.

## Customizing comment guidelines

The file `references/commenting.md` contains the editorial guidelines Claude follows when drafting comments (what to look for, what makes a good comment, etc.). Edit this file to change the kinds of feedback Claude provides.

## File structure

```
comment-on-docx/
├── SKILL.md                          # Main skill instructions for Claude
├── scripts/
│   ├── read_document_runs.py         # Reads and numbers document text runs
│   └── docx_comment_helper.py        # Adds comments via python-docx
└── references/
    └── commenting.md                 # Editorial guidelines for comment quality
```
