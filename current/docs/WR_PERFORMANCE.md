# W/R Performance

## Scan W/R Performance

The `Scan W-R Performance` worksheet has one table per WT with a Scan four-quadrant Wafer Map. DC Pass Count is the number of distinct Raw X/Y coordinates minus the Chips indexed as gray and unavailable for that WT's Scan Map. The raw, gray and eligible counts can be reconciled within each table.

W, R and W/R count independently. W/R uses the Map's Write&Read Fail category. The W+R, W+W/R, R+W/R and W+R+W/R columns count exact combinations of those three modes. X does not enter those combinations. Each table shows the Fail Chips count and the numeric percentage to two decimal places. A zero DC Pass Count displays dashes for the percentages.

Import and Analyze and Refresh Maps rebuild both reports. When a separate Address or FBC import invalidates the Maps, both reports clear.

## W&R Performance

v13.2 adds a W&R Performance worksheet, rebuilt by Import and Analyze and Refresh Maps.

Each Pattern produces one table. The first header is the Pattern name and W/R Performance, followed by Wafer and Total Die. Imported WT1, WT2 and WT3 each have a merged Chips header and W, R, RW and X percentage columns. Each data row contains the first six LOT characters and the two-digit Wafer number.

The numerator is the number of distinct Chips with a positive count in the corresponding displayed four-quadrant Map category for the same WT and Pattern. A Chip counts once per category regardless of its captured Address count. A Chip can appear in several categories, so the percentages can total more than 100%.

The denominator uses the Raw Chip Locations coordinate set, deduplicated by X/Y, minus Chips whose WT1 Cat is C10, C11, C14, C15, C16 or C17. Those six categories are excluded from numerators too. Other recorded Cats and Chips without a recorded Cat are included in the denominator, following the requested explicit exclusion list. C20 is included. The WT1 denominator is shared across imported WTs.

The summary displays percentages to two decimal places. Missing WT/Pattern panels and percentages with a zero denominator display a dash.

Each available WT/Pattern panel has a detail table and chart below the summary. The first four rows count W, R, RW and X independently. The next 15 rows count exact combinations, including zero-count combinations. For example, W + R requires positive W and R counts and zero positive RW and X counts. A Chip with all four Fail modes contributes once to each independent row and once to the W + R + RW + X combination.

Fail Count is a Chip count. DC Pass Chips is the shared denominator defined above. The chart plots the counts and displays the denominator in its title. The detail table includes the share for each row. The 15 exact combination counts plus No recorded Fail (Chips) sum to DC Pass Chips. The latter row can include Chips without usable panel data.

Importing Address or FBC separately clears the summary, detail tables and charts until the full analysis is refreshed.

The report reads the generated Wafer Maps cell index. It uses the same Hard Fail, previous WT and previous Pattern eligibility already reflected in those Maps.

All table cells use Calibri 11, horizontal and vertical center alignment. Worksheet gridlines are hidden. W is yellow, R white, RW orange, and X grey.

Native validation is in current/tools/test_native_wr_performance.ps1. It covers three WTs, two Patterns, the six DC categories, repeated rows for the same Chip, zero counts, multi-category Chips, all 15 combinations, chart data, missing panels, 3-to-1 WT reruns, division by zero and stale report invalidation.
