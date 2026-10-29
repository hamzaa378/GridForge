# GridForge v10.4 — Ten external Excel function names

Our frozen 475-name reference catalog previously lacked ten function names. The v10.4 parser now **recognizes** all ten, validates argument counts and reports **explicit, connection-dependent errors**. These are *not* working local implementations of external-data behavior.

| Function(s) | Excel service/model required | GridForge offline response | Real connection parity |
| --- | --- | --- | --- |
| `CUBEKPIMEMBER`, `CUBEMEMBER`, `CUBEMEMBERPROPERTY`, `CUBERANKEDMEMBER`, `CUBESET`, `CUBEVALUE` | Workbook OLAP cube / MDX connection | `#NAME?`, explains unavailable OLAP connection | **Not tested** |
| `GETPIVOTDATA` | Excel PivotTable metadata and values | `#REF!`, explains missing PivotTable engine | **Not tested** |
| `RTD` | Windows COM real-time data server | `#N/A` with explanation | **Not tested** |
| `STOCKHISTORY` | Microsoft 365 licensed market-data provider | `#CONNECT!` with explanation | **Not tested** |
| `WEBSERVICE` | Web endpoint and allowed outgoing request | `#CONNECT!` with explanation | **Not tested** |

No automatic HTTP calls are made for formulas; this is intentional. A browser cannot reproduce the Excel PivotTable and OLAP workbook object models simply by recognizing their function names. We do not reinterpret `GETPIVOTDATA` as an arbitrary range sum because that would be incorrect.

**Scope math:** 506 existing locally implemented names + 10 external recognized names = **516 recognized**. Our 475-name *bounded* comparison then has **zero unrecognized names**; it does not have 475 fully Excel-compatible functions, and it does not represent the entire Microsoft 365 function universe.

## Testing

- `tests/connected_v10_4.test.mjs` checks the ten names, categories, required arguments and non-fabrication of results.
- `npm test` preserves the original 506-function comparison fixture (does not silently reclassify these ten as executable).
- `python scripts/excel_connected_windows.py` checks **missing connection error behavior** for a subset in actual Excel, if run on Windows with desktop Excel + pywin32.
- RTD, STOCKHISTORY and WEBSERVICE are skipped in that runner by default, avoiding unapproved outbound or automation traffic.
- To establish real functional parity in future, the project would need separate PivotTable, cube query, data-vendor, RTD and network integrations with specific test fixtures and explicit user permissions.

## Microsoft reference pages

- https://support.microsoft.com/en-us/excel/cube-functions-reference
- https://support.microsoft.com/en-us/excel/functions/getpivotdata-function
- https://support.microsoft.com/en-us/excel/functions/stockhistory-function
- https://support.microsoft.com/en-us/excel/functions/webservice-function
