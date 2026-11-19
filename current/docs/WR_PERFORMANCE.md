# W/R Performance

v13.2 adds a W&R Performance worksheet, rebuilt by Import and Analyze and Refresh Maps.

Each Pattern produces one table. The first header is the Pattern name and W/R Performance, followed by Wafer and Total Die. Imported WT1, WT2 and WT3 each have a merged Chips header and W, R, RW and X percentage columns. Each data row contains the first six LOT characters and the two-digit Wafer number.

The numerator is the number of distinct Chips with a positive count in the corresponding displayed four-quadrant Map category for the same WT and Pattern. A Chip counts once per category regardless of its captured Address count. A Chip can appear in several categories, so the percentages can total more than 100%.

The denominator uses the Raw Chip Locations coordinate set, deduplicated by X/Y, minus Chips whose WT1 Cat is C10, C11, C14, C15, C16 or C17. Those six categories are excluded from numerators too. Other recorded Cats and Chips without a recorded Cat are included in the denominator, following the requested explicit exclusion list. C20 is included. The WT1 denominator is shared across imported WTs.

Percentages use the numeric Excel format 0.0%, with no raw Chip or Fail Bit count in the four output columns. A zero denominator displays a dash. Missing WT panels display a dash. Only Patterns and WTs present in the four-quadrant panel index are output. Importing Address or FBC separately invalidates the old report until the full analysis is refreshed.

The report reads the generated Wafer Maps cell index. It uses the same Hard Fail, previous WT and previous Pattern eligibility already reflected in those Maps.

All table cells use Calibri 11, horizontal and vertical center alignment. Worksheet gridlines are hidden. W is yellow, R white, RW orange, and X grey.

Native validation is in current/tools/test_native_wr_performance.ps1. It covers three WTs, two Patterns, the six DC categories, repeated rows for the same Chip, zero counts, multi-category Chips, output formatting, 3-to-1 WT reruns, division by zero and stale report invalidation.
