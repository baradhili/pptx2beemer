# Changelog — xml2tex fork, `front-page-layout` vs `master`

Goal of this fork: the LaTeX-side table rendering features that the docx2tex pipeline on
the `front-page-layout` branch needs to reproduce Word table styling. 1 commit ahead of
master.

- **Row shading and border rule colours in `calstable2tabular`** (`13336a5`): CALS table
  row shading (`\rowcolor`) and per-rule border colours (`\arrayrulecolor`) are carried
  into the generated `tabular`, so the navy header row and grey grid of Word tables render
  as in the reference.
