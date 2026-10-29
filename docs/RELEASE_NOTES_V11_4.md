# GridForge 11.4.0 Release Notes

## Responsive chart rendering

- Worksheet chart SVGs now render in queued, one-chart-at-a-time tasks after the grid has painted. Scrolling and selection changes can be processed between chart renders.
- A newer worksheet render cancels stale queued chart work, preventing outdated chart images from being painted after edits or navigation.

## Chart editing

- The Format ribbon now includes direct editing for horizontal and vertical axis titles. The chart editor already supported these settings; the ribbon now exposes them beside the chart title, vertical scale, number format and size controls.

## Desktop startup

- The local WebView server binds to an available loopback port per launch. This removes collisions with another GridForge process holding a previous fixed port.
- Removed the circular reference from the exposed desktop API object back to its pywebview window. The packaged app was becoming unresponsive during startup; the rebuilt executable remained responsive and accepted worksheet input.
- The desktop launcher version metadata and troubleshooting guide now match the 11.4 release.

GridForge remains a standalone spreadsheet with its own worksheet UI and calculation engine. Microsoft Excel is not embedded or required.
