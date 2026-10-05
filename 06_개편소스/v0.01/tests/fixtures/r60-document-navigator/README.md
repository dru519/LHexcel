# Native navigator fixtures

Created by the project's own `T_R60DocumentNavigator.CreateFixture` in candidate07 on Windows Excel, 2026-09-06. No third-party data or macros.

- `rich.xlsx`: Start, protected Target, Hidden, VeryHidden, ChartOnly; note, formatting, formula, conditional format and defined name.
- `simple.xlsx`: Start and Target with the same note, formula and formatting baseline.

The suite hashes and copies these files into its evidence directory before opening working copies. They replace repeated `Range.AddComment`, which began crashing the installed Excel build in ntdll before the navigator opened. Reuse preserves note/shape/format assertions rather than omitting them.
