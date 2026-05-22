# Glossary application — progress tracker

Canonical Greek renderings live in `glossary.md` (sections Α/Β/Γ/Δ, with the user's edits).
Rule: **first use** of a term in the BODY (Ch.1 onward) gets the gloss/English-in-parens; later uses are plain.
The YAML abstract/keywords/prologue/acknowledgements are kept clean (translate common terms, no parens).

## Position — COMPLETE
- Done: front matter + **Ch.1–Ch.12 (all)**. Ch.13 = bibliography (English citations, left as-is).
- Final build clean: 143 pages, 0 errors, 0 missing chars.
- Remaining English in body = intentional: (a) English-in-parens glosses «Greek (english)»,
  (b) kept proper/product names + standard tech terms (Elixir module, projections, Use Cases/
  Gateway layers, pod, book-and-claim, email), (c) code identifiers, (d) bibliography titles.
- Ch.4: added Greek renderings after the 5 English mission block-quotes (*Ελληνική απόδοση:* «…»).
- Ch.3: glossed upstream/downstream, financial/operational control, fallback, cost/activity-based,
  LNG, SBTi-aligned; made all "mass balance" → «ισοζύγιο μάζας» + Greek-first headings.
- Added glosses in Ch2: net zero, methanol-ready, container ship, premium, allowance,
  e-fuels, status quo, ESG reporting→αναφορά ESG, ad hoc, audit-ready, Excel-based workflow.
  (Ch2 already had many Greek-first glosses by the author: self-service, audit trail,
  scalability, data freshness, data tracking, error rates, EUA, CII, ESG, etc.)

## First-use gloss already placed (term → where)
GHG, container carriers, shipping company, integrated logistics provider (§1.1);
event-driven, event sourcing, real-time, Extreme Programming (§1.2);
trade factors, mass balance, book-and-claim, XP, production team, pair programming,
test-driven development, continuous integration, vertical ownership, GIA (§1.3);
baseline, POS, solver, observability, monorepo, NRG/:nrg, components, pipeline,
web layer, deployment, TDD, feedback loops, onboarding, analytics (§1.4 roadmap).

## Notes / decisions
- Re-glossing a major term again at its dedicated chapter's first substantive use is OK
  (reader-friendly for a long thesis). Rule that matters: leave NO raw untranslated English;
  minor terms glossed once.
- English-in-parens uses *italics*; when nested inside a Greek parenthetical use [brackets].
- Abstract & front matter: translate common terms to Greek, keep product names + acronyms, no parentheticals.
- event sourcing → «προέλευση από συμβάντα (event sourcing)» on first body use.
- Extreme Programming / XP: methodology name, keep English, gloss «(ακραίος προγραμματισμός)» on first body use.
- net zero (generic) → «καθαρές μηδενικές εκπομπές»; "Net Zero 2040" goal name kept.
