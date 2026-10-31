# GridForge 12.0.24

Draw and Page Layout now apply compact ribbon dimensions as soon as the tab is selected. The selection handler explicitly hides the standard command ribbon while either dedicated ribbon is active, and locks the active surface to 96 px with vertical overflow clipped. Switching to another tab removes these inline layout locks and restores the regular ribbon behavior.

Verification: TypeScript static check and Windows packaging with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.