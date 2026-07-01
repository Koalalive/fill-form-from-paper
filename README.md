# Fill Form from Paper

Automate filling structured Word templates from research papers. Extract content, inject into correct paragraphs, and iteratively refine formatting through automated post-processing.

## Quick Start

```bash
python scripts/preflight.py template.docx
python scripts/fill_form.py
python scripts/fix_all.py output.docx
python scripts/insert_figure.py figure.png ./output/
python scripts/rollback.py ./ --list
```

## Scripts

| Script | Purpose |
|--------|---------|
| preflight.py | Validate template |
| fill_form.py | Inject content |
| fix_all.py | One-shot formatting fix |
| insert_figure.py | Insert images |
| rollback.py | Version rollback |
| audit.py | Bold run scanner |

## License

MIT