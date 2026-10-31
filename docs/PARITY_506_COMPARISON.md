# GridForge 506-function comparison with desktop Excel

> **Historical comparison note (v12.0.57):** The reported counts are from the earlier one-case-per-function run. Source updates since that run corrected the listed `FIELDVALUE` behavior, the `INFO("release")` / `INFO("numfile")` results, and the `FORECAST.ETS` argument handling. The Excel comparison was not rerun for those changes, so the 450/456 score remains historical rather than a v12.0.61 score.

Excel 16.0 (build 20430.0); **456 comparable cases**, **450 matches**, **6 differences**.

Comparable-case agreement: **98.68%**. This is one invocation per registered name, not proof of general Excel compatibility.

## Outcome summary

- **MATCH**: 450
- **DIFFERENCE**: 2
- **KNOWN_SEMANTIC_DIFFERENCE**: 4
- **EXCEL_UNAVAILABLE_OR_REJECTED**: 36
- **EXCEL_EXECUTION_FAILURE**: 6
- **ONLINE_FUNCTION_NOT_RUN**: 3
- **VOLATILE_NONDETERMINISTIC**: 5

## Differences and limitations

- **ACCRINT** [EXCEL_EXECUTION_FAILURE]: `=ACCRINT(45000,45300,45500,0.05)`  GridForge: `{'type': 'number', 'value': 68.33333333333333}`; Excel: `{'type': 'formula-assignment-error', 'value': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)", 'formula2Error': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)"}`
- **ACCRINTM** [EXCEL_EXECUTION_FAILURE]: `=ACCRINTM(45000,45300,0.05)`  GridForge: `{'type': 'number', 'value': 40.833333333333336}`; Excel: `{'type': 'formula-assignment-error', 'value': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)", 'formula2Error': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)"}`
- **BYCOL** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(BYCOL(M1:N3,"SUM"),1,1)`  GridForge: `{'type': 'number', 'value': 9}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **BYROW** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(BYROW(M1:N3,"SUM"),1,1)`  GridForge: `{'type': 'number', 'value': 3}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **CHOOSECOLS** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(CHOOSECOLS(M1:N3,1),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **CHOOSEROWS** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(CHOOSEROWS(M1:N3,1),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **COVARIANCEP** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=COVARIANCEP(A1:A5,B1:B5)`  GridForge: `{'type': 'number', 'value': 4}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **COVARIANCES** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=COVARIANCES(A1:A5,B1:B5)`  GridForge: `{'type': 'number', 'value': 5}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **CUBESETCOUNT** [KNOWN_SEMANTIC_DIFFERENCE]: `=CUBESETCOUNT("A,B,C")`  GridForge: `{'type': 'number', 'value': 3}`; Excel: `{'type': 'error', 'value': '#N/A'}`
- **DETECTLANGUAGE** [ONLINE_FUNCTION_NOT_RUN]: `=DETECTLANGUAGE("Hello")`  GridForge: `{'type': 'string', 'value': 'en'}`; Excel: `{'type': 'online-function-skipped', 'value': 'Opt in using --allow-online-functions; may contact online services.'}`
- **DROP** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(DROP(M1:N3,1),1,1)`  GridForge: `{'type': 'number', 'value': 3}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **EUROCONVERT** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=EUROCONVERT(100,"EUR","DEM")`  GridForge: `{'type': 'number', 'value': 195.583}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **EXPAND** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(EXPAND(M1:N3,5,3,0),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **FIELDVALUE** [KNOWN_SEMANTIC_DIFFERENCE]: `=FIELDVALUE("cost=500;region=North","cost")`  GridForge: `{'type': 'string', 'value': '500'}`; Excel: `{'type': 'error', 'value': '#FIELD!'}`
- **FLATTEN** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(FLATTEN(M1:N3),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **FORECAST.ETS.SEASONALITY** [KNOWN_SEMANTIC_DIFFERENCE]: `=FORECAST.ETS.SEASONALITY(5,P1:P4)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#N/A'}`
- **FORECAST.ETS.STAT** [EXCEL_EXECUTION_FAILURE]: `=FORECAST.ETS.STAT(5,P1:P4)`  GridForge: `{'type': 'number', 'value': 0}`; Excel: `{'type': 'formula-assignment-error', 'value': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)", 'formula2Error': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)"}`
- **GROUPBY** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(GROUPBY(M1:M3,P1:P3,"SUM"),1,1)`  GridForge: `{'type': 'string', 'value': 'A'}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **HSTACK** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(HSTACK(M1:N3,P1:Q3),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **IMAGE** [ONLINE_FUNCTION_NOT_RUN]: `=IMAGE("https://example.org/sample.png","Example")`  GridForge: `{'type': 'string', 'value': 'Example'}`; Excel: `{'type': 'online-function-skipped', 'value': 'Opt in using --allow-online-functions; may contact online services.'}`
- **IMPOWER** [DIFFERENCE]: `=IMPOWER("1+i","2")`  GridForge: `{'type': 'string', 'value': '1.22464679914735E-16+2i'}`; Excel: `{'type': 'string', 'value': '1.22514845490862E-16+2i'}`
- **INFO** [KNOWN_SEMANTIC_DIFFERENCE]: `=INFO("release")`  GridForge: `{'type': 'string', 'value': 'GridForge 10'}`; Excel: `{'type': 'string', 'value': '16.0'}`
- **ISBETWEEN** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=ISBETWEEN(5,1,10)`  GridForge: `{'type': 'boolean', 'value': True}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **ISOMITTED** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=ISOMITTED(L40)`  GridForge: `{'type': 'boolean', 'value': True}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **JIS** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=JIS("Hello")`  GridForge: `{'type': 'string', 'value': 'Ｈｅｌｌｏ'}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **LAMBDA** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=LAMBDA("x",2)`  GridForge: `{'type': 'number', 'value': 2}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **LET** [EXCEL_EXECUTION_FAILURE]: `=LET("x",2,2)`  GridForge: `{'type': 'number', 'value': 2}`; Excel: `{'type': 'formula-assignment-error', 'value': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)", 'formula2Error': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)"}`
- **LOGNORMDIST** [EXCEL_EXECUTION_FAILURE]: `=LOGNORMDIST(1,0,1,TRUE)`  GridForge: `{'type': 'number', 'value': 0.5}`; Excel: `{'type': 'formula-assignment-error', 'value': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)", 'formula2Error': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)"}`
- **MAKEARRAY** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(MAKEARRAY(3,3,"PRODUCT"),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **MAP** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(MAP(M1:N3,"ABS"),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **NETWORKDAYSINTL** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=NETWORKDAYSINTL(DATE(2026,10,19),DATE(2026,10,23),1)`  GridForge: `{'type': 'number', 'value': 5}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **NOW** [VOLATILE_NONDETERMINISTIC]: `=NOW()`  GridForge: `{'type': 'number', 'value': 46322.66599822917}`; Excel: `{'type': 'number', 'value': 46316.66505740741}`
- **PERCENTOF** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=PERCENTOF(M1:M2,P1:P3)`  GridForge: `{'type': 'number', 'value': 0.5}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **PHONETIC** [EXCEL_EXECUTION_FAILURE]: `=PHONETIC("Good morning")`  GridForge: `{'type': 'string', 'value': 'Good morning'}`; Excel: `{'type': 'formula-assignment-error', 'value': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)", 'formula2Error': "(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146827284), None)"}`
- **PIVOTBY** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(PIVOTBY(M1:M3,P1:P3,S1:S3,"SUM"),1,1)`  GridForge: `{'type': 'string', 'value': 'A'}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **RAND** [VOLATILE_NONDETERMINISTIC]: `=RAND()`  GridForge: `{'type': 'number', 'value': 0.2974123997042267}`; Excel: `{'type': 'number', 'value': 0.8448650233261406}`
- **RANDARRAY** [VOLATILE_NONDETERMINISTIC]: `=INDEX(RANDARRAY(2,3),1,1)`  GridForge: `{'type': 'number', 'value': 0.2919310922845496}`; Excel: `{'type': 'number', 'value': 0.6728049226244295}`
- **RANDBETWEEN** [VOLATILE_NONDETERMINISTIC]: `=RANDBETWEEN(1,6)`  GridForge: `{'type': 'number', 'value': 6}`; Excel: `{'type': 'number', 'value': 3.0}`
- **REDUCE** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=REDUCE(0,P1:P4,"SUM")`  GridForge: `{'type': 'number', 'value': 9}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **REGEXEXTRACT** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=REGEXEXTRACT("AB123","[0-9]+")`  GridForge: `{'type': 'string', 'value': '123'}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **REGEXREPLACE** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=REGEXREPLACE("ab12","[0-9]+","ZZ")`  GridForge: `{'type': 'string', 'value': 'abZZ'}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **REGEXTEST** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=REGEXTEST("ab12","[0-9]+")`  GridForge: `{'type': 'boolean', 'value': True}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **SCAN** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(SCAN(0,P1:P4,"SUM"),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **TAKE** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(TAKE(M1:N3,1),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **TEXTAFTER** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=TEXTAFTER("one-two-three","-",2)`  GridForge: `{'type': 'string', 'value': 'three'}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **TEXTBEFORE** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=TEXTBEFORE("one-two-three","-",2)`  GridForge: `{'type': 'string', 'value': 'one-two'}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **TEXTSPLIT** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(TEXTSPLIT("a,b,c",","),1,1)`  GridForge: `{'type': 'string', 'value': 'a'}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **TOCOL** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(TOCOL(M1:N3),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **TODAY** [VOLATILE_NONDETERMINISTIC]: `=TODAY()`  GridForge: `{'type': 'number', 'value': 46322}`; Excel: `{'type': 'number', 'value': 46316.0}`
- **TOROW** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(TOROW(M1:N3),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **TRANSLATE** [ONLINE_FUNCTION_NOT_RUN]: `=TRANSLATE("hello","en","es")`  GridForge: `{'type': 'string', 'value': 'hola'}`; Excel: `{'type': 'online-function-skipped', 'value': 'Opt in using --allow-online-functions; may contact online services.'}`
- **TRIMRANGE** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(TRIMRANGE(M1:N3),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **VSTACK** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(VSTACK(M1:N3,P1:Q3),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **WRAPCOLS** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(WRAPCOLS(M1:M4,2),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **WRAPROWS** [EXCEL_UNAVAILABLE_OR_REJECTED]: `=INDEX(WRAPROWS(M1:M4,2),1,1)`  GridForge: `{'type': 'number', 'value': 1}`; Excel: `{'type': 'error', 'value': '#NAME?'}`
- **XNPV** [DIFFERENCE]: `=XNPV(0,F1:F5,G1:G5)`  GridForge: `{'type': 'number', 'value': 400}`; Excel: `{'type': 'error', 'value': '#NUM!'}`

## Methodology and caveats

Each case invokes the named function in an isolated Excel worksheet state. Some matrix results are read with INDEX rather than testing automatic spill. Time-dependent functions are excluded from exact equality; Excel-version availability is reported separately. Other inputs, locales, error cases, workbooks, external links, macros, and dynamic-array behavior remain unverified by these 506 fixtures.
