# Repriori

A SaMD-programme workflow tool: projects (DD | CRCO | Innovation) move through
JSON-defined stage gates, documents move through a Pre-draft → Draft → Review
→ Approval lifecycle, and every workflow is an independently-shippable
package wired together over HTTP, never by compile-time reference.

This repository is being built **phase by phase**, in the same style as the
`CalcSaMD` project: small steps, one living design document, and a
file-by-file walkthrough for every phase. If you are picking this up from
scratch, start with `docs/Repriori_Plan_and_Design.docx` — Part 6 is the full
phase roadmap, and Part 7 onward is a complete, in-order build log with the
exact commands to run and what you should see at each step.

## Solution file: `Repriori.slnx`

Same as CalcSaMD, this uses Visual Studio 2026's newer XML-based `.slnx`
format rather than a legacy `.sln` file.

## Folder layout (grows as phases land)

```
Repriori/
  Repriori.slnx
  README.md
  docs/
    Repriori_Plan_and_Design.docx
  src/                    (each subfolder is its own independently-deployable project — see Part 5 of the plan)
  tests/                  (one test project per package in src/)
```

`src/` and `tests/` are intentionally empty right now — Phase 1 only lays the
solution skeleton down. The first real project (`Repriori.DocProfile`, the
Word document format-comparer) is added in Phase 3.
