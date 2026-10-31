# GridForge v12.0.55

- Formula tokenization now accepts newer Excel error literals including `#FIELD!`, `#BLOCKED!`, `#CONNECT!`, `#UNKNOWN!`, `#BUSY!`, and `#PYTHON!`, allowing formulas and error checks to parse them.
- Recognizing an error code does not supply its linked-data service or Python environment.
- Static TypeScript check and production build only so far; no test or live UI phases run.
