# v12.0.168 — ODS chart axis fidelity

Native OpenDocument chart import/export now retains the GridForge settings that ODF can represent for common charts:

- logarithmic value axes with base 10
- category-axis label rotation
- reversed category order, including horizontal bars

ODS represents logarithmic scaling as a boolean, so non-base-10 logarithmic settings are kept in GridForge's own chart metadata but are not exported as native ODF logarithmic axes. Axis number formats, category tick intervals, secondary-axis combo charts, advanced external chart formatting, and live source linkage remain open. No automated tests or LibreOffice UI round-trip were run.
