# formulae | macros | code snippets

<br />

##### [How to extract the formula text in Excel (works with Google Sheets)](https://www.sageintelligence.com/tips-and-tricks/excel-tips-tricks/2017/02/extract-text-formula-excel/#:~:text=Select%20cell%20C16%20and%20enter,corrections%20or%20audit%20the%20formula.)
```
=FORMULATEXT(cell)
```

<br />

##### HYPERLINK Generator
Step 1: Create input & output sheets <br />
Step 2: Place the below formula in A1 cell of the output sheet & drag to right or down until it covers the entire range
```
=IF(LEFT('Sheet'!Cell,4)="http",CONCATENATE("=HYPERLINK(""",'Sheet'!Cell,""",""🔗""",")"),IF('Sheet'!Cell="🔗",FORMULATEXT('Sheet'!Cell),'Sheet'!Cell))
```
Step 3: Copy data from your data sheet & paste into cell A1 of input sheet <br />
Step 4: Copy result from output sheet & **paste as value** into your data sheet <br />
Step 5: Press CTRL+F on data sheet & click the overflow/kebab/3 vertical dots menu icon <br />
Step 6: Follow the below: <br />
  <ul><li> Find: http </li></ul>
  <ul><li> Replace with: (empty) </li></ul>
  <ul><li> Search: This Sheet </li></ul>
  <ul><li> All checkboxes unmarked </li></ul>
Step 7: Press Replace all button & notice the HYPERLINK formulae are now hyperlinks <br />

<br />

##### [Extract URLs in Google Sheets from Hyperlinks](https://infoinspired.com/google-docs/spreadsheet/extract-urls-in-google-sheets-without-script/)
###### Note: Only works with standard HYPERLINK formula <code> =HYPERLINK(cell or "URL","label") </code>
```
=REGEXEXTRACT(FORMULATEXT(cell),"""(.*)"",")
```

<br />

##### TEXTJOIN without duplicates if multiple column ranges match the value in another cell (A2)
```
=TEXTJOIN(" | ",TRUE,TRANSPOSE(SORT(UNIQUE(TRIM(TRANSPOSE(SPLIT(TEXTJOIN("|",TRUE,IFERROR(SPLIT(
TEXTJOIN("/ ",TRUE, 
  IFERROR(FILTER($U$2:$U,$T$2:$T=A2),""), 
  IFERROR(FILTER($Z$2:$Z,$Y$2:$Y=A2),""), 
  IFERROR(FILTER($AF$2:$AF,$AE$2:$AE=A2),""), 
  IFERROR(FILTER($AL$2:$AL,$AK$2:AK=A2),""), 
  IFERROR(FILTER($AR$2:$AR,AO$2:$AQ=A2),""), 
  IFERROR(FILTER($AX$2:$AX,$AW$2:$AW=A2),""), 
  IFERROR(FILTER($BD$2:$BD,$BC$2:$BC=A2),""), 
  IFERROR(FILTER($BJ$2:$BJ,$BI$2:$BI=A2),""), 
  IFERROR(FILTER($BP$2:$BP,$BO$2:$BO=A2),""), 
  IFERROR(FILTER($BV$2:$BV,$BU$2:$BU=A2),"") )
,"/"))),"|")))))))
```

<br />

##### Generate a list without duplicates from multiple column ranges
```
=SORT(VALUE(UNIQUE(QUERY({'Table'!T2:T;'Table'!Y2:Y;'Table'!AE2:AE;'Table'!AK2:AK;'Table'!AQ2:AQ;'Table'!AW2:AW;'Table'!BC2:BC;'Table'!BI2:BI;'Table'!BO2:BO;'Table'!BU2:BU},"select * where Col1 <>'' or Col1 is not null"))),1,1)
```

<br />

##### TEXTJOIN specified strings if cell (A2) found matches in multiple columns
```
=TEXTJOIN(", ",TRUE,IF(ISNA(MATCH(A2,$T$2:$T,0)),"","English"),IF(ISNA(MATCH(A2,$Y$2:$Y,0)),"","Deutsch"),IF(ISNA(MATCH(A2,$AE$2:$AE,0)),"","Español"),IF(ISNA(MATCH(A2,$AK$2:$AK,0)),"","Français"),IF(ISNA(MATCH(A2,$AQ$2:$AQ,0)),"","Italiano"),IF(ISNA(MATCH(A2,$AW$2:$AW,0)),"","Japanese"),IF(ISNA(MATCH(A2,$BC$2:$BC,0)),"","Nederlands"),IF(ISNA(MATCH(A2,$BI$2:$BI,0)),"","Svenska"),IF(ISNA(MATCH(A2,$BO$2:$BO,0)),"","Dansk"),IF(ISNA(MATCH(A2,$BU$2:$BU,0)),"","Português") )
```

<br />

#### [Removing everything but emojis in javascript for google sheets script](https://stackoverflow.com/questions/48755842/removing-everything-but-emojis-in-javascript-for-google-sheets-script)
```
=REGEXREPLACE(cell or range,"[[:print:]]","")
```

<br />

#### [Find and Replace emojis in Google Sheets](https://stackoverflow.com/questions/43501740/find-and-replace-emojis-in-google-sheets)
```
=ARRAYFORMULA(REGEXREPLACE(range,"[\x{1F300}-\x{1F64F}]|[\x{2702}-\x{27B0}]|[\x{1F68}-\x{1F6C}]|[\x{1F30}-\x{1F70}]|[\x{2600}-\x{26ff}]|[\x{D83C}-\x{DBFF}\x{DC00}-\x{DFFF}]",""))
```

<br />

#### [How Search and Find Emojis on Sheets](https://stackoverflow.com/questions/53883089/how-search-and-find-emojis-on-sheets)
```
=ARRAYFORMULA(REGEXREPLACE(range,"[^[:ascii:]]",))
```

<br />

#### TEXTJOIN only Emojis from a range
```
=REGEXREPLACE(TEXTJOIN("",0,range),"[[:print:]]","")
```
or
```
=REGEXREPLACE(TEXTJOIN("",0,range),"[[:ascii:]]","")
```

<br />

#### EN Alphanumeric only - remove ascii/emoji / non-English characters
```
=REGEXREPLACE(
LOWER(CLEAN(TRIM(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(
cell,
CHAR(127)," "),
CHAR(129)," "),
CHAR(141)," "),
CHAR(143)," "),
CHAR(144)," "),
CHAR(157)," "),
CHAR(160)," "),
CHAR(202)," "),
"!"," "),
"@"," "),
"#"," "),
"$"," "),
"%"," "),
"^"," "),
"&"," "),
"*"," "),
"("," "),
")"," "),
"?"," "),
"/"," "),
"‐"," "),
"−"," "),
"_"," "),
"+"," "),
"="," "),
"{"," "),
"["," "),
"}"," "),
"]"," "),
"|"," "),
"\"," "),
":"," "),
";"," "),
""""," "),
"<"," "),
">"," "),
","," "),
"."," "),
"⟨"," "),
"⟩"," "),
"、"," "),
" ‒ "," "),
"‒"," "),
"–"," "),
"—"," "),
"―"," "),
"…"," "),
"⋯"," "),
"᠁"," "),
"ฯ"," "),
"‹"," "),
"›"," "),
"«"," "),
"»"," "),
"⧸"," "),
"⁄"," "),
"·"," "),
"‱"," "),
"•"," "),
"†"," "),
"‡"," "),
"⹋"," "),
"‘"," "),
"’"," "),
"“"," "),
"”"," "),
"°"," "),
"〃"," "),
"¡"," "),
"¿"," "),
"※"," "),
"×"," "),
"•"," "),
"№"," "),
"÷"," "),
"º"," "),
"ª"," "),
"‰"," "),
"¶"," "),
"±"," "),
"∓"," "),
"′"," "),
"″"," "),
"‴"," "),
"§"," "),
"~"," "),
"‖"," "),
"¦"," "),
"©"," "),
"🄯"," "),
"℗"," "),
"®"," "),
"℠"," "),
"™"," "),
"¤"," "),
"؋"," "),
"​₳"," "),
"฿"," "),
"₿"," "),
"₵"," "),
"¢"," "),
"₡"," "),
"₢"," "),
"₠"," "),
"​₫"," "),
"​₻"," "),
"​₯"," "),
"​֏"," "),
"₠"," "),
"​€"," "),
"ƒ"," "),
"​₣"," "),
"​₶"," "),
"​₷"," "),
"₲"," "),
"₴"," "),
"₭"," "),
"₺"," "),
"​₾"," "),
"​₼"," "),
"​ℳ"," "),
"​ℛℳ"," "),
"​​₥"," "),
"​₦"," "),
"​₧"," "),
"​₱"," "),
"​₰"," "),
"​₴"," "),
"​£"," "),
"元"," "),
"​圆"," "),
"​圓"," "),
"﷼ "," "),
"​៛"," "),
"₽"," "),
"​​₹"," "),
"₨"," "),
"​₪"," "),
"৳"," "),
"₸"," "),
"₮"," "),
"​₩"," "),
"¥"," "),
"​​円"," "),
"৲"," "),
"৹"," "),
"৻"," "),
"𐆚"," "),
"𐆖"," "),
"𐆙"," "),
"𐆗"," "),
"𐆘"," "),
"߾"," "),
"߿"," "),
"𞲰"," "),
"⁂"," "),
"❧"," "),
"☞"," "),
"‽"," "),
"⸮"," "),
"◊"," "),
"⁀"," "),
"،"," ")
)))
,"[^[:ascii:]]",)
```

<br />
