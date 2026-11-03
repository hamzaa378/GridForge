# Frozen formula fixture: ODS export coverage

Standard OpenFormula expressions: **404/506 (79.84%)**.
Microsoft compatibility-namespace formulas: **84/506**.
LibreOffice-specific formulas: **3/506**.
GridForge-only recovery formulas: **15/506**; metadata recovery copies are present on 18 formulas.

One frozen desktop-Excel parity formula per case. Standard OpenFormula, Microsoft-namespace, LibreOffice-specific, and GridForge-only serialization are counted separately; none proves full calculation parity.

## LibreOffice-specific formulas with a GridForge recovery copy

- **REGEXEXTRACT**: `=REGEXEXTRACT("AB123","[0-9]+")`
- **REGEXREPLACE**: `=REGEXREPLACE("ab12","[0-9]+","ZZ")`
- **REGEXTEST**: `=REGEXTEST("ab12","[0-9]+")`

## Microsoft compatibility-namespace formulas

- **BETA.DIST**: `=BETA.DIST(0.5,2,2,TRUE)`
- **BETA.INV**: `=BETA.INV(0.5,2,2)`
- **BINOM.DIST**: `=BINOM.DIST(2,4,0.5,FALSE)`
- **BINOM.INV**: `=BINOM.INV(10,0.5,0.5)`
- **CEILING.MATH**: `=CEILING.MATH(7.2,2)`
- **CEILING.PRECISE**: `=CEILING.PRECISE(-4.1,2)`
- **CHISQ.DIST**: `=CHISQ.DIST(2,3,TRUE)`
- **CHISQ.DIST.RT**: `=CHISQ.DIST.RT(2,3)`
- **CHISQ.INV**: `=CHISQ.INV(0.5,3)`
- **CHISQ.INV.RT**: `=CHISQ.INV.RT(0.5,3)`
- **CHISQ.TEST**: `=CHISQ.TEST(M1:N1,P1:Q1)`
- **CHOOSECOLS**: `=INDEX(CHOOSECOLS(M1:N3,1),1,1)`
- **CHOOSEROWS**: `=INDEX(CHOOSEROWS(M1:N3,1),1,1)`
- **CONCAT**: `=CONCAT("A","B")`
- **CONFIDENCE.NORM**: `=CONFIDENCE.NORM(0.05,2,100)`
- **CONFIDENCE.T**: `=CONFIDENCE.T(0.05,3,20)`
- **DROP**: `=INDEX(DROP(M1:N3,1),1,1)`
- **EXPAND**: `=INDEX(EXPAND(M1:N3,5,3,0),1,1)`
- **EXPON.DIST**: `=EXPON.DIST(2,0.5,TRUE)`
- **F.DIST**: `=F.DIST(1,4,8,TRUE)`
- **F.DIST.RT**: `=F.DIST.RT(1,4,8)`
- **F.INV**: `=F.INV(0.6,4,8)`
- **F.INV.RT**: `=F.INV.RT(0.6,4,8)`
- **F.TEST**: `=F.TEST(M1:P1,P1:S1)`
- **FILTER**: `=INDEX(FILTER(M1:N3,P1:P3),1,1)`
- **FLOOR.MATH**: `=FLOOR.MATH(7.2,2)`
- **FLOOR.PRECISE**: `=FLOOR.PRECISE(-4.1,2)`
- **FORECAST.LINEAR**: `=FORECAST.LINEAR(6,B1:B5,A1:A5)`
- **GAMMA.DIST**: `=GAMMA.DIST(2,3,2,TRUE)`
- **GAMMA.INV**: `=GAMMA.INV(0.5,3,2)`
- **GAMMALN.PRECISE**: `=GAMMALN.PRECISE(5)`
- **HSTACK**: `=INDEX(HSTACK(M1:N3,P1:Q3),1,1)`
- **HYPGEOM.DIST**: `=HYPGEOM.DIST(1,3,4,12,FALSE)`
- **IFS**: `=IFS(FALSE,1/0,TRUE,42)`
- **ISO.CEILING**: `=ISO.CEILING(5.2,2)`
- **ISOMITTED**: `=ISOMITTED(L40)`
- **LAMBDA**: `=LAMBDA("x",2)`
- **LET**: `=LET("x",2,2)`
- **LOGNORM.DIST**: `=LOGNORM.DIST(1,0,1,TRUE)`
- **LOGNORM.INV**: `=LOGNORM.INV(0.5,0,1)`
- **MAXIFS**: `=MAXIFS(A1:A5,D1:D5,"East")`
- **MINIFS**: `=MINIFS(A1:A5,D1:D5,"East")`
- **MODE.SNGL**: `=MODE.SNGL(1,2,2,3)`
- **NORM.DIST**: `=NORM.DIST(0,0,1,TRUE)`
- **NORM.INV**: `=NORM.INV(0.5,0,1)`
- **NORM.S.DIST**: `=NORM.S.DIST(0,FALSE)`
- **NORM.S.INV**: `=NORM.S.INV(0.5)`
- **PERCENTILE.EXC**: `=PERCENTILE.EXC(A1:A5,0.5)`
- **PERCENTILE.INC**: `=PERCENTILE.INC(A1:A5,0.75)`
- **PERCENTRANK.EXC**: `=PERCENTRANK.EXC(A1:A5,3)`
- **PERCENTRANK.INC**: `=PERCENTRANK.INC(A1:A5,3)`
- **POISSON.DIST**: `=POISSON.DIST(0,2,FALSE)`
- **QUARTILE.EXC**: `=QUARTILE.EXC(A1:A5,1)`
- **QUARTILE.INC**: `=QUARTILE.INC(A1:A5,1)`
- **RANDARRAY**: `=INDEX(RANDARRAY(2,3),1,1)`
- **RANK.AVG**: `=RANK.AVG(2,A1:A5)`
- **RANK.EQ**: `=RANK.EQ(4,A1:A5)`
- **SEQUENCE**: `=INDEX(SEQUENCE(3,2),1,1)`
- **SORT**: `=INDEX(SORT(M1:N3,1,-1),1,1)`
- **SORTBY**: `=INDEX(SORTBY(M1:N3,P1:P3),1,1)`
- **STDEV.P**: `=STDEV.P(A1:A5)`
- **STDEV.S**: `=STDEV.S(A1:A5)`
- **SWITCH**: `=SWITCH(2,1,1/0,2,99,0)`
- **T.DIST**: `=T.DIST(1,7,TRUE)`
- **T.DIST.2T**: `=T.DIST.2T(1,7)`
- **T.DIST.RT**: `=T.DIST.RT(1,7)`
- **T.INV**: `=T.INV(0.5,7)`
- **T.INV.2T**: `=T.INV.2T(0.5,7)`
- **T.TEST**: `=T.TEST(M1:P1,P1:S1,2,2)`
- **TAKE**: `=INDEX(TAKE(M1:N3,1),1,1)`
- **TEXTJOIN**: `=TEXTJOIN("|",TRUE,D1:D3)`
- **TEXTSPLIT**: `=INDEX(TEXTSPLIT("a,b,c",","),1,1)`
- **TOCOL**: `=INDEX(TOCOL(M1:N3),1,1)`
- **TOROW**: `=INDEX(TOROW(M1:N3),1,1)`
- **UNIQUE**: `=INDEX(UNIQUE(M1:M4),1,1)`
- **VAR.P**: `=VAR.P(A1:A5)`
- **VAR.S**: `=VAR.S(A1:A5)`
- **VSTACK**: `=INDEX(VSTACK(M1:N3,P1:Q3),1,1)`
- **WEIBULL.DIST**: `=WEIBULL.DIST(2,1,2,TRUE)`
- **WRAPCOLS**: `=INDEX(WRAPCOLS(M1:M4,2),1,1)`
- **WRAPROWS**: `=INDEX(WRAPROWS(M1:M4,2),1,1)`
- **XLOOKUP**: `=XLOOKUP("k2",D1:D3,E1:E3,"missing")`
- **XMATCH**: `=XMATCH("East",D1:D5)`
- **Z.TEST**: `=Z.TEST(A1:A5,3)`

## Formulas stored only in GridForge recovery metadata

- **BYCOL**: `=INDEX(BYCOL(M1:N3,"SUM"),1,1)`
- **BYROW**: `=INDEX(BYROW(M1:N3,"SUM"),1,1)`
- **CUBESETCOUNT**: `=CUBESETCOUNT("A,B,C")`
- **DETECTLANGUAGE**: `=DETECTLANGUAGE("Hello")`
- **FIELDVALUE**: `=FIELDVALUE("cost=500;region=North","cost")`
- **GROUPBY**: `=INDEX(GROUPBY(M1:M3,P1:P3,"SUM"),1,1)`
- **IMAGE**: `=IMAGE("https://example.org/sample.png","Example")`
- **MAKEARRAY**: `=INDEX(MAKEARRAY(3,3,"PRODUCT"),1,1)`
- **MAP**: `=INDEX(MAP(M1:N3,"ABS"),1,1)`
- **PERCENTOF**: `=PERCENTOF(M1:M2,P1:P3)`
- **PIVOTBY**: `=INDEX(PIVOTBY(M1:M3,P1:P3,S1:S3,"SUM"),1,1)`
- **REDUCE**: `=REDUCE(0,P1:P4,"SUM")`
- **SCAN**: `=INDEX(SCAN(0,P1:P4,"SUM"),1,1)`
- **TRANSLATE**: `=TRANSLATE("hello","en","es")`
- **TRIMRANGE**: `=INDEX(TRIMRANGE(M1:N3),1,1)`
