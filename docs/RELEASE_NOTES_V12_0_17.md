# GridForge 12.0.17

The Draw and Page Layout toolbars are now explicitly constrained to a single 96 px ribbon row as direct children of the app shell. The matching normal ribbon row is forcibly hidden on those tabs, and the inactive specialized row is hidden. This protects the worksheet from ribbon flex growth that caused the tall blank panel shown in the reported screenshot.

This is a code-level layout correction. No live UI check or Windows UI test was run, so the visual result still needs confirmation on the user's setup.

Verification: static TypeScript check, Python syntax compilation, production build, and Windows packaging with `-SkipTests`.
