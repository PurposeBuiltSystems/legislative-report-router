# Parser rules

Deterministic, staged, line-aware. No AI.

## Pipeline

```
Email HTML (or DOCX text)
  → HTML normalization (line-semantic tags → \n, entities decoded, <a href> captured)
  → line normalization (CRLF/LF, nbsp, trim)
  → header removal (everything before the first boundary line)
  → entry boundary detection
  → field extraction
  → division normalization
  → confidence scoring → human review
```

## Boundary rule (the important one)

A new legislative entry starts **only** at a line that consists entirely of
one identifier + number:

```regex
^(HF|SF|HSB|SSB|HJR|SJR|HCR|SCR|HR|SR)\s*-?\s*(\d+[A-Z]?)$
```

Because the whole line must match, references *inside* a brief — "Successor
to HSB171, as amended." — can never split an entry. They are instead
collected into `referencedBills`.

Identifier prefixes are configurable (`opts.identifiers`); longest-first
matching prevents "SF" from shadowing "SSB".

## Field extraction per entry

1. First non-blank line after the bill number: if it looks like a division
   designation → `distributedTo`.
2. Second such line → `commentRequestedFrom`; if absent, mirrors
   `distributedTo` with a warning (−0.10 confidence).
3. Everything after the division lines → `brief` (first line ≤120 chars also
   becomes `title`). Blank-line runs collapse.
4. No division line at all → warning, −0.35 confidence.
5. No brief → warning, −0.25 confidence.

"Looks like a division designation": ≤60 chars, no sentence-ending
punctuation, and every token (split on `/ ; , and &`) is either a known
division (from the routing rules — codes, names, aliases), a 2–6 letter
all-caps code, or "<Word> Office/Bureau/Division". This is what stops the
first line of a lowercase brief ("open records") from being eaten as a code.

## Division normalization

`MVD/TDD`, `MVD; SOD`, `ELT and MVD`, `MVD, TDD, and SOD` → `[MVD, TDD, …]`.
Multi-word divisions ("AG Office") survive because splitting only happens on
separators, never on spaces.

## What the parser never does

- Never discards unrecognized text silently — it stays in the brief and the
  per-item `sourceBlock`, and warnings surface in the UI.
- Never auto-publishes: low-confidence items are flagged, and everything
  passes through the Review screen.
- Never splits on identifiers inside sentences.
