---
name: fill-form-from-paper
description: Use when filling a structured form/template document from a source paper, when the user says "fill this form based on this paper", "populate the template from the research", "generate an application form from academic content", or when creating formal documents (grant applications, patent forms, award nominations, invention reports) from papers. Triggers: docx template + paper source.
---

# Fill Form from Paper

## Overview

Automate filling a structured Word template from a research paper. Extract content from the paper, inject into correct template paragraphs without destroying formatting, then iteratively refine through 20+ rounds of user feedback.

## Quick Start

```bash
python scripts/preflight.py template.docx
python scripts/fill_form.py
python scripts/fix_all.py output.docx
python scripts/insert_figure.py figure.png ./output/
python scripts/rollback.py ./ --list
```

## Workflow

PREFLIGHT -> INJECT -> FIX_ALL -> AUDIT

## Core Rules

### Template preservation
- NEVER modify original template. Always versioned output.
- Naming: {base}_v{N}.docx, global counter.
- fix_all.py auto-increments to max(existing_v) + 1.

### Formatting standards
| Element | Font | Size | Align | Line Spacing | Bold |
|---------|------|------|-------|-------------|------|
| Body text | SimSun | 12pt | JUSTIFY | 22pt | No |
| H1 | SimHei | 14pt | CENTER | 28pt | Yes |
| H2 | SimSun | 12pt | JUSTIFY | 28pt | Yes |
| H3 | SimSun | 12pt | JUSTIFY | 22pt | First sentence |
| Figure caption | SimSun | 12pt | JUSTIFY | 22pt | Title only |
| All paragraphs | space_before=0 | space_after=0 | first_line_indent=0 | | |
| Margins | 2.9cm both sides | | | | |

### Bolding discipline
- H1/H2: FULL line bold
- H3: ONLY first sentence bold
- Body text: NEVER bold entire paragraphs
- RED FLAG: bold run > 1 sentence = BUG

## Common Mistakes
- Template modified in-place -> Always versioned output
- Empty paragraphs between headers -> fix_all.py removes all
- Entire paragraphs bolded -> Split runs, first sentence only

## Iteration Log
21 versions from real session. Three parallel subagent reviews caught 19 content discrepancies and 7 formatting issues.
