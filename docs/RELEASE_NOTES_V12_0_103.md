# GridForge v12.0.103

## Drawing access and chart precision

- Drawing objects receive accessible names and can be reached by keyboard while Select Objects is active. Enter or Space selects the focused object; Ctrl/Shift modifiers add or remove it from the selection.
- Ctrl+D and Delete/Backspace are suppressed while a text field, dropdown, or editable element has focus, preserving normal text editing.
- Linear least-squares fits now center values and use compensated summation, reducing cancellation for large x-coordinates such as date serials. Degenerate/non-finite fits still return a chart error rather than propagating invalid coordinates.
- Changing drawing color or width now affects only the corresponding style property; removing the last object clears its rendered SVG layer.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
