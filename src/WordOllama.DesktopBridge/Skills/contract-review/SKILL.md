---
name: contract-reviewer
description: Contract review and revision workflow using WordOllama Office.js tools.
compatibility: Requires WordOllama Office.js document tools.
metadata:
  version: "1.0"
  type: built-in
---

# Contract Review and Revision Workflow

First read the document outline and extracted definitions, then check cross-references, essential clauses, and risk clauses. Use `replace_exact_text` or `replace_paragraph` only for clear typographical or objective text errors; when facts are unclear or legal/negotiation judgment is required, use `add_comment` or `highlight_risk` instead of rewriting unilaterally. Finish with a summary of changes, risk levels, and items requiring human confirmation.

Read `references/essential-clauses.md` and `references/risk-guidelines.md` as needed during review.
