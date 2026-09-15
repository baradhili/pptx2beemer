# Changelog — docx2hub fork, `front-page-layout` vs `master`

Goal of this fork: surface the Word features that the docx2tex pipeline on the
`front-page-layout` branch needs to reproduce a styled Word report (reference
document `HSS_REP.docx`, a Health Support Services template) — page geometry,
image crops, header/footer parts, section boundaries, section page numbering and
per-section page margins. 9 commits ahead of master (front-page-layout: 7, further-fixes: 2).

- **Surface page geometry, header/footer part rels and section page breaks** (`305578a`):
  the first section's `pgSz`/`pgMar` are emitted as `css:*` attributes on the hub root;
  image relationships in header/footer parts resolve against the part's own rels
  (`@xml:base`-based) instead of the main document rels; the docx2hub:header/footer divs
  keep their `@xml:base` and are ordered by section-reference position (first section
  first); the paragraph before a removed `sectPr` pseudo-paragraph is marked
  `css:page-break-after` so Word section breaks survive into the hub.
- **Keep embedded pictures out of the SVG renderer** (`b2cd89b`): AlternateContent drawings
  that contain embedded pictures are no longer converted to SVG, which used to swallow the
  logo pictures.
- **Fix lowercase hex colours** (`ef5606a`): `docx2hub:color()` matched hash-less hex only
  in uppercase (`[0-9A-F]{6}`), but Word writes colours lowercase — `ff0000`, `1b2546`,
  `3d3935`, `a6a6a6` matched no branch and were silently dropped (digit-only colours like
  `152147` worked, masking the bug). Now accepts both cases and normalises to uppercase.
- **Surface Word image crops** (`29662fa`): `a:srcRect` values ≤ 100000 are
  ST_Percentage in 1000ths of a percent (ISO 29500-1), not EMU — the old `css:clip`
  formula divided by 12700 and produced nonsense, and nothing consumed it. Crops are now
  emitted as `css:crop-top/right/bottom/left` percentages on `imagedata`, from both the
  DrawingML `a:srcRect` and the VML `crop*` (65536th-fraction) paths.
- **Accept capitalized built-in heading names when rewriting styleIds** (`86db888`):
  Word built-in styles are stored as lowercase `heading 1`, but some authoring tools
  write `Heading 1` (capitalized) with numeric styleIds (e.g. 1278). The case-sensitive
  regex never matched, so paragraphs kept `_1278`-style roles that no conf template
  recognizes and numbered headings fell through to enumerate lists. The name match and
  the `heading ` strip in the `Heading{N}` rewrite are now case-insensitive.
- **Surface section page-number format/restart and section boundaries** (`9d58a9d`): the
  add-props pass captures each `sectPr`'s `w:pgNumType` as `css:page-number-*` on the
  sectPr marker paras; wml-to-dbk flags the paragraph that ends a section
  (`docx2hub:section-end`) and copies the page-number attributes of the section that
  begins after the break onto it, so downstream converters can switch numbering at Word
  section boundaries.
- **Surface per-section page margins and header distance** (`21d7839`): each `sectPr`'s
  `w:pgMar` (top/bottom/left/right, twips → pt) and `w:header` distance are captured as
  `css:page-margin-*`/`css:page-header-distance` on the marker paras and copied onto the
  section-break paragraphs (final section via the body-level sectPr), enabling
  per-section `\newgeometry` downstream — e.g. a closing page with a tall custom top
  margin.
- **Surface the Normal style's default line spacing on the hub root** (`25e0fcb`):
  a "1.5 lines" Normal style (`w:line=360`, `lineRule=auto`) spaces the whole
  document, but the style never appears as a css:rule because role-less paragraphs
  do not reference it. The root now carries `css:default-line-height`
  (`w:line` div 240) when it differs from single spacing.
- **Always surface the Normal style's space-after on the hub root** (`1128ced`):
  `css:default-space-after` defaults to 0pt when the Normal style declares no
  `w:after` (Word's built-in default) instead of omitting the attribute — the
  downstream `\parskip` is fully driven by the document.
