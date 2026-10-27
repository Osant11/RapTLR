# Add news here

# RapTLR 0.0.4

## Bug Fixes

- `run_apdx()`: Output IDs in the List of Tables/Figures (TOC) are now always **bold**,
  including for multi-page outputs. Previously, the `fpar()` call omitted the
  `fp_text(bold = TRUE)` property on the ID text, so no output ID was ever bold.
  Additionally, the cross-reference field expression was being built before `idui` was
  assigned, which was a latent ordering bug in the loop. Both issues are fixed in
  `addTLF2TLR()`. (#addTLF2TLR)
