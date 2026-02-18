# comment-on-docx

A Claude Code skill that add comments to Microsoft Word (.docx) documents. It reads the document structure, drafts comments following editorial guidelines, and writes them back as native Word comments.

For Google docs, you can download them as a Word doc, then have Claude comment it up. You could also then re-upload it to Google Drive (and optionally convert it back to a google doc file.)

## How to use

### On Claude.ai

Upload the `comment-on-docx.skill` file to [claude.ai/settings/capabilities](https://claude.ai/settings/capabilities). Claude will now use this skill every time you upload a Word doc and ask for comments on it. 

### Locally with Claude Code

This skill requires a python environment with the `python-docx >= 1.2.0` package. You might want to tell Claude which local python environment to use. Currently, `SKILL.md` has this quote about my (Tim's) local environment, which you should swap out.

> If you're in the conda evals environment, python-docx is installed there. Use: /Users/timhua/anaconda3/bin/conda run -n evals

Then, you can ask Claude to install the skill for you and use it directly from Claude code.

## Remainder of the ReadME file is Claude written:

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

On my personal computer, the command is:

```bash
rm -rf ~/.claude/skills/comment-on-docx && cp -r /Users/timhua/Documents/aisafety_githubs/comment-on-docx/comment-on-docx ~/.claude/skills/comment-on-docx && echo "Done" && ls ~/.claude/skills/comment-on-docx/
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

## Development

### Auto-packaging Hook (Optional)

This repo includes a git pre-commit hook that automatically regenerates `comment-on-docx.skill` whenever you commit changes to the skill directory. Git hooks are local-only (not pushed to GitHub for security), so you'll need to install it manually after cloning.

**To install the hook:**

```bash
cp scripts/hooks/pre-commit .git/hooks/pre-commit
chmod +x .git/hooks/pre-commit
```

**Or manually package the skill:**

```bash
zip -r comment-on-docx.skill comment-on-docx/
```

The hook ensures the `.skill` file stays in sync with the source code automatically.
