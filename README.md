# comment-on-docx

<small>February 2026</small>

<small>[Tim Hua](https://timhua.me/)</small>

This is Claude Code skill that lets Claude add comments to Microsoft Word documents (.docx). It reads the document (including figures), drafts comments following editorial guidelines in `references/commenting.md`, and adds them as native Word comments. It doesn't modify the original file and instead outputs a "claude commented" version.

To use this skill together with Google docs, you'll need to download the GDoc as a Word document first.[^1] You could also then re-upload it to Google Drive (and optionally convert it back to a google doc file.) It might be more convenient to just have Claude's word comments open on the side while one incorporates the changes in the main google doc though.

[^1]: There technically is a Google Doc API, but everyone who I've talked to has had a bad time getting it to work, so we're sticking with this workaround for now.

## How to use

### On Claude.ai

Upload the `comment-on-docx.skill` file to [claude.ai/settings/capabilities](https://claude.ai/settings/capabilities). Claude will now use this skill every time you upload a Word doc and ask for comments on it. 

### Locally with Claude Code

This skill requires a python environment with the `python-docx >= 1.2.0` package. You might want to first tell Claude which local python environment to use by modifying `SKILL.md`. Currently, `SKILL.md` has this quote about my (Tim's) local environment, which you should swap out.

> If you're in the conda evals environment, python-docx is installed there. Use: /Users/timhua/anaconda3/bin/conda run -n evals

Then, you can ask Claude to install the skill for you and use it directly from Claude code.

## Claude's workflow

Claude first reads the Word file using the `read_document_runs.py` script. The read script allows Claude to view images, footnotes, links, and tables. Then, Claude is supposed to re-read the `references/commenting.md` guidelines, and draft its comments. After successfully drafting the comments, Claude is told to reflect on them to check that they are indeed good comments. 

## Usage notes:

- Existing comment threads are squashed if you download a Google doc (or, at least it is squashed for me). Claude can see where the existing comments are, but it is currently not instructed to respond to them.

- If there are suggestions with track changes on, Claude will read the file as if all changes are accepted.

- By default, I noticed that Claude tends to draft around ten comments in total regardless of the lengths of the piece. You can just prompt it to write more or less comments, although if you're doing this it might not bring up all the grammatical errors.

- Claude tends to barrel ahead even when something goes wrong. If comments appear missing or the document seems cut off, it may continue drafting anyway rather than stopping to flag the issue. I believe the read issues that caused truncation are now fixed, and `SKILL.md` instructs Claude to stop and raise an error if it suspects any part of the document wasn't fully read. Still, this may be a source of error.


# Remainder of the ReadME file is Claude written:

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

Or you can use a symlink.

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
