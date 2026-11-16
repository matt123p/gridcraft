# Main synchronization validation

The feature branch incorporates main at
`ef99f1e51f5f971a7c83da5acbc4dc7bcc62f2c5`. VBA documentation changes are confined
to `docs/vba/`; the main README and general documentation match main.

Passed checks:

- Native VBA host tests: 315 passed.
- Strict Clippy for the VBA host and VBA-enabled UI, including all targets.
- Workspace formatting and strict Clippy.
- Asset attribution, dependency layering and the WASM check.
- Workspace doctests, parity consistency and generated keytip consistency.
- Relative documentation links and pinned tiny-vba reference targets.

The complete `cargo xtask ci` run is **not green**. Its workspace test step stops
at seven `gridcraft-ui-egui` Arabic bidi tests because the current build cannot
find Arabic glyphs. A temporary fontconfig setup did not resolve the failure in
the current test executable. The Arabic test source is unchanged from main.
Later test binaries were not reached by that run. Historical milestone receipts
in this folder retain their original revision scope.
