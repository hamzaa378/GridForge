# GridForge 12.0.47

External JSON refresh now flattens nested object records into worksheet columns. Flattened paths are bounded to 32 levels and 512 characters, output remains limited to 104 columns and the existing cell cap, and ambiguous paths are rejected with a clear error. JSON arrays still need a selected JSON Pointer path or a separate expansion step.

Verification is limited to static TypeScript checking, Python syntax compilation, and production packaging with Windows tests skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
