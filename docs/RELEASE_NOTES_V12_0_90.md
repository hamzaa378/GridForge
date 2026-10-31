# GridForge v12.0.90

## TEXTBEFORE and TEXTAFTER empty inputs

- Empty source text with a non-empty delimiter now follows the not-found behavior: `#N/A` or the supplied `if_not_found` value.
- Empty delimiter behavior and positive/negative instance handling remain unchanged.

## Verification

- TypeScript static validation and Windows release compilation use `-SkipTests`.
- No app launch, live UI, Windows runtime, or Excel GUI tests were run.
