=====================

### INDEX

 * [GENERAL](#GENERAL)
 * [MBIE](#MBIE)
 * [MARITIME](#MARITIME)
 * [MOE](#MOE)
 






# GENERAL
###### [LIBRARY](https://github.com/ScottypNZ/CODE-LIBRARY)   |   [INDEX](#INDEX)
-------------------------

LINKS 
https://opdhsblobprod04-secondary.blob.core.windows.net/contents/9b520b3cc93e43f7ba25a428fd5605f4/b10c45193088c620dcde2762fc236c51?sv=2018-03-28&sr=b&si=ReadPolicy&sig=ujmH19rSqr0MgAy4GPoNrFQymGs4im8K19miED2HgCs%3D&st=2021-07-25T04%3A57%3A19Z&se=2021-07-26T05%3A07%3A19Z

https://www.excelguru.ca/blog/2014/08/20/5-very-useful-text-formulas-power-query-edition/

https://docs.microsoft.com/en-us/powerquery-m/power-query-m-function-reference

https://www.goodly.co.in/extract-first-and-last-name-power-query/

https://www.myonlinetraininghub.com/extract-letters-numbers-symbols-from-strings-in-power-query-with-text-select-and-text-remove

```
ADD COLUMN AND FITER AT THE SAME TIME   
= Table.AddColumn(#"Replaced NULL", "SPECIAL", each Text.Select([Full Name],{"A".."z"," "})<>[Full Name])

COUNT NULL VALUES
=List.Count(List.RemoveNulls(Record.ToList(Record.RemoveFields(_, "Vehicle ID"))))

RUNNING TOTAL
= Table.AddColumn(#"Added SUB", "CUM", each 
List.Sum( List.FirstN( #"Added SUB"[SUB], [Index] ) ) , Int64.Type)

ADD ADDRESS AND FORMAT AT THE SAME TIME   
Table.AddColumn(#"Added WEEK CREATED", "ADDRESS", each 
Text.Upper(Text.Combine(List.Select(Text.Split(
[Address 1]&" "&[Address 2]&" "&[Suburb]&" "&[City]&" "&[Post Code]&" "&[Country],
" "),each _ <> "")," ")))

ADD CURRENT WEEK FILTER   
Table.AddColumn(#"Replaced NULL", "FILTER", each Date.From(Date.StartOfWeek(Date.AddDays(DateTime.LocalNow(),-7) ) )
    = Date.StartOfWeek([Created On]))

MODIFIED ON WEEK
Table.AddColumn(#"Added FILTER", "WEEK MODIFIED", each Text.From(Date.Year(Date.From([Modified On])))
    &"/"& Text.PadStart(Text.From(Date.WeekOfYear(Date.From([Modified On]))),2,"0"))

CREATED ON WEEK   
Table.AddColumn(#"Added WEEK MODIFIED", "WEEK CREATED", each Text.From(Date.Year(Date.From([Created On])))&"/"& Text.PadStart(Text.From(Date.WeekOfYear(Date.From([Created On]))),2,"0"))

COUNT BLANK COLUMNS   
= List.NonNullCount(Record.FieldValues(_))

**TABLE PIVOT **  
DEFAULT              = Table.Pivot(Source, List.Distinct(Source[Name]), "Name", "Value")
TABLE PIVOT MANUAL   = Table.Pivot(Source, {"value","@odata.nextLink"}, "Name", "Value")

CONVERT LIST TO TABLE   
= Table.AddColumn( #"Expanded rows", "Records", each Record.FromList([rows],[columnNames] ) )

O-DATA SVC   
= OData.Feed("http://mnz-16-app02-p/MaritimeNewZealand/XRMServices/2011/OrganizationData.svc/ContactSet?$select=Home2&$top=3")
= OData.Feed("http://mnz-16-app02-p/MaritimeNewZealand/XRMServices/2011/OrganizationData.svc/ContactSet?&$top=3")

O-DATA API  
= OData.Feed("https://crm/maritimenewzealand/api/data/v8.0/contacts")

JSON API  
= Json.Document(Web.Contents("https://crm/maritimenewzealand/api/data/v8.0/contacts?$top=10"))

REMOVE LINE BREAKS IN ADDRESS - (SINGLE LINE)
= Table.ReplaceValue(Reordered,"#(cr)#(lf)",", ",Replacer.ReplaceText,{"address1_composite"})

REMOVE NUMBERS NEW COLUMN
= Table.AddColumn(#"Replaced PHONE -", "MOBILE", each Text.Select([Mobile Phone],{"+".."9"}))

COUNT NON BLANK COLUMNS IN RANGE
=List.NonNullCount({[Col1],[Col2],[Col3]})

AVERAGE  
= Table.AddColumn(#"Grouped REFERENCE", "AVERAGE", each List.Average(#"Grouped REFERENCE"[COUNT]), type number)

ADD LEAVING ZERO TO NUMBER  
= Table.AddColumn(#"Changed Type", "Average", each Text.PadStart(Text.From([number]),3,"0"))

MAX  
= Table.AddColumn(#"Added MIN", "MAX", each List.Max({Number.From([START]),Number.From([END])}), type date)

MIN  
= Table.AddColumn(#"Added Room End", "MIN", each List.Min({Number.From([START]),Number.From([END])}), type date)

FILTER ROW  
= Table.SelectRows(#"Added Date Range", each [END] <> null)

FILTER ERRORS  
= Table.SelectRowsWithErrors(#"Replaced null rooms", {"HOTEL"})

FILTER OR AND  
= Table.SelectRows(#"Expanded Grouped REFERENCE", each ([ROOMS] = "0909" or [ROOMS] = "0937") and ([BHR] = "BHR-042420"))

CEILING / BUCKET  
= Table.AddColumn(#"Added Custom2", "BUCKET", each Number.RoundUp([COUNT]/20,0)*20, Int64.Type)

CREATE WITH 3DP  
TBC

TEXT CONTAINS  
= Table.AddColumn(#"Changed Type", "TEST", each Text.Contains([#"Flight #"],"1"))

COUNT WORDS   
=List.Count(Text.SplitAny([FULL NAME]," ")))

NET WORD COUNT (LESS BLANKS)  
=List.Count(List.RemoveItems(Text.SplitAny([FULL NAME]," #(tab)#(lf)"),{""})))

SHORT NAME UTILISING NET WORD COUNT (TO STOP SINGLE WORD DUPLICATION)  
=if [NetWordCount] = 1 then [FULL NAME] else Text.Upper(List.First(Text.Split([FULL NAME]," ")))&" "&Text.Upper(List.Last(Text.Split([FULL NAME]," ")))

SHORT NAME  
= Table.AddColumn(#"Added REFERENCE","TEST", each Text.Proper(List.First(Text.Split([FirstName]," ")) &" "& List.Last(Text.Split([Last.Name]," "))), type text)

REMOVE EXCESS SPACES  
= Table.AddColumn(#"Added RoomH", "ROOMI", each Text.Combine(List.Select(Text.Split([ROOMH]," "),each _ <> "")," "))

TRANSPOSE DATES BETWEEN A RANGE  
= Table.AddColumn(#"Filtered Rows", "DATE", each { Number.From([START])..Number.From([END]) })
 
MERGE QUERIES (AFTER GROUP BY)  
= Table.NestedJoin(#"Reordered Columns",{"REFERENCE"},#"Grouped Rows1",{"REFERENCE"},"Grouped Rows1",JoinKind.LeftOuter)

SPIT DATE   
= Table.SplitColumn(Table.TransformColumnTypes(#"Added SHORT NAME", {{"BirthDate", type text}}, "en-NZ"), "BirthDate", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"BirthDate.1", "BirthDate.2"})  

CONVERT DATE TO NUMBER (SO DONT NEED TO SPLIT)   
= Table.AddColumn(#"Added DOB", "DOB.1", each Text.PadStart(Text.From(Date.From(Number.From(Date.From([BirthDate])))),10,"0"))  

CONVERT DATE FROM YYYY/MM/DD    
 Date.ToText(Date.From(Text.From([BirthDate])), "dd/MM/yyyy"))    

ADD ROW IF TABLE IS EMPTY   
= if Table.IsEmpty(#"Reordered Columns") = true then Table.InsertRows(#"Reordered Columns", 0,{Record.FromList(List.Repeat({""},Table.ColumnCount(#"Reordered Columns")),Table.ColumnNames(#"Reordered Columns"))}) else #"Reordered Columns"

IMPORT ONLY SELECTED TABLES FROM EXISTING QUERY   
 = Table.SelectColumns(Source,{"MIAS DATE", "PassengerID"}),

TEXT REPLACEMENT   
= Text.ReplaceAll([Text], Replacements)

REFERENCE QUERY  
= "BHR"

ADD YEAR AND WEEK  
= Table.AddColumn(#"Added MONTH", "WEEK", each Date.ToText([DATE],"yyyy")&"/"&Text.From(Date.WeekOfYear([DATE])))

EXTRACT MONTH IN "MMM" FORMAT  
= Date.ToText([Date],"MMM" ) & "_" & Text.From(Date.Year([Date]))

EXTRACT MONTH IN "MM" FORMAT INCLUDING LEADING ZERO  
= Text.PadStart(Text.From(Date.Month([ArrivalDate]),2,"0")))

LEFT 10 CHARACTERS CONVERTED TO TEXT, TRIMMED AND THEN PADDED TO 10  
= Text.PadStart(Text.Trim(Text.Start(Text.From([START]),10)),10,"0"))

COUNT SPECIFIC NUMBER IN SPECIFIED LIST OF COLUMNS
= Table.AddColumn(#"Changed column type", "ISSUE", each 
List.Count(List.Select(Record.ToList(Record.SelectFields(_,{"INITIAL FIRST","INITIAL SECOND","INITIAL LAST"} ) ), each _ = "1") ) )

EXTRACT YEAR AND MONTH  
= Table.AddColumn(#"Added DATE REF", "MONTH", each Date.ToText([DATE],"yyyy")&"/"&Date.ToText([DATE],"MM"))
= Table.AddColumn(#"Added WEEK", "MONTH", each Date.ToText([DATE],"yyyy") &"/"& Text.PadStart(Text.From(Date.Month([DATE])),2,"0"), type text)

PAD START OF DATE AND SUBSTITUE NULL VALUES  
= Table.AddColumn(#"Changed Type", "REFERENCE", 
each if [ArrivalDate] = null then 
Text.PadStart(Text.From([PassengerID]),6,"0")&" "&"00/00/0000"
else 
Text.PadStart(Text.From([PassengerID]),6,"0")&" "&Text.PadStart(Text.From([ArrivalDate]),10,"0"))

REPLACE NULL WITH "XXXXXXX" WHEN SPLITTING FIRST NAME  
= Table.AddColumn(#"Changed Type", "NameRef",
each if [FirstName] = null then "XXXXXX"
else 
Text.Upper(Text.BeforeDelimiter([FirstName]," "))

REPLACE NULL WITH "XXXXXXX" WHEN SPLITTING LAST NAME  
= Table.AddColumn(#"Added NAME REF", "Custom", 
each if [LastName] = null then "XXXXXX"
else 
Text.Upper(Text.AfterDelimiter([LastName]," ")))

REPLACE VALUES WHERE CONDITION 
= Table.ReplaceValue(#"PREVIOUS STEP","064", 
each if Text.Start([Mobile Phone],3) = "064" then "0064" 
else #"PREVIOUS STEP",Replacer.ReplaceText,{"Mobile Phone"} )

OPTION1 GROUP BY ID (2 STEPS) "ALL" NEEDS TO BE ADDED   
= Table.Group(Source, {"ID"}, {{"MAX", each List.Max([YEAR]), type number}, {"All", each _, type table }})  
= Table.ExpandTableColumn(#"Grouped Rows", "All", {"YEAR", "ROW3", "ROW4", "ROW5"})
GROUPS BY ID
NEW COLUMN NAME MAX/MAX/YEAR (LAST IS MAX IF VALUE)   
NEW COLUMN NAME ALL/ALLROWS   
EXPAND ALL LESS ID WHICH IS ALREADY IN TABLE.   
UNTICK PREFIX   

OPTION 2 GROUP BY ID (3 STEPS) "ALL" DOES NOT NEED TO BE ADDED   
= Table.Group(Source, {"ID"}, {{"MAX", each List.Max([YEAR]), type number}})   
= Table.NestedJoin(Source, {"ID"}, Table1, {"ID"}, "Grouped Rows", JoinKind.FullOuter)   
= Table.ExpandTableColumn(#"Merged Queries", "Grouped Rows", {"MAX"}, {"MAX"})   
GROUPS BY ID   
NEW COLUMN NAME MAX/MAX/YEAR (LAST IS MAX IF VALUE)  
MERGE QUERIES FULL OUTER ON ID/ID   
EXPAND MAX   

```



FINANCIAL YEAR FROM DATE
```
= Table.AddColumn(#"Added YEAR", "Financial Year", each 
if Date.Month([LastModified]) < 7 then "FY" &
Text.End(Text.From(Date.Year([LastModified])-1),2) &"/"&
Text.End(Text.From(Date.Year([LastModified])),2)
 else "FY" &
Text.End(Text.From(Date.Year([LastModified])),2) &"/"&
Text.End(Text.From(Date.Year([LastModified])+1),2))
```

### COUNTIFS
```
# STEP 1 - GROUP BY PRIMARY FIELD
GROUP BY [SHORT NAME REFERENCE] | ALL | ALL ROWS

# STEP 2 - ADD COLUMN [FILTER], SECOND VARIABLE = FILTER. IN THIS CASE 1
=Table.SelectRows([ALL], each ([OCC] = 1))

# STEP 3 - ADD COLUMN2, [FILTER COUNT] 
=Table.RowCount([FILTER])

# STEP 4 REMOVE [FILTER] & [SHORT NAME REF]

# STEP 5 EXPAND ALL

# STEP 6 FILTER [COUNT FILTER] =1 

```

#COUNTIFS TEMPLATE

```
VBA
let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],

    #"Grouped Rows" = Table.Group(Source, {"Number"}, {{"Count", each _, type table [Type=text, Number=number, Countifs=number, Index=number]}}),

    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "SelectB", each Table.SelectRows([Count], each ([Type] = "B"))),

    #"Added Custom1" = Table.AddColumn(#"Added Custom", "CountB", each Table.RowCount([SelectB])),

    #"Removed Other Columns" = Table.SelectColumns(#"Added Custom1",{"Count", "CountB"}),
    
    #"Expanded Count" = Table.ExpandTableColumn(#"Removed Other Columns", "Count", {"Type", "Number", "Countifs"}, {"Type", "Number", "Countifs"})
in
    #"Expanded Count"
```

### OCCURANCE / RUNNING COUNT
```VBA
1 ADD INDEX COLUMN
2 GROUP BY RELEVANT VARIABLE (ALL ROWS [ALL] & COUNT ROWS [COUNT]
3 ADD CUSTOM COLUMN Table.AddIndexColumn([ALL], "OCCURANCE", 1, 1))
4 REMOVE ALL 
5 EXPAND OCCURANCE

NB: Count is not necessary.
NB: Make sure Group name is distinct
NB: Make sure occurance is not in the data 
NB: Index is to make sure the data stays in the correct order, without it it will be unsorted.
NB: LINK https://musamomodu.com/count-occurrences-countif-in-power-query/
```




### REMOVE DUPLICATES SYSTEM FUNCTION
```VBA
    #"Kept Duplicates" = 
    let columnNames = {"REFERENCE"}, 
    addCount = Table.Group(#"ReportView Application_Seafarer Application ALL", columnNames, {{"Count.1", Table.RowCount, type number}}), 
    selectDuplicates = Table.SelectRows(addCount, each [Count.1] > 1), 
    removeCount = Table.RemoveColumns(selectDuplicates, "Count.1") 
    in Table.Join(#"ReportView Application_Seafarer Application ALL", columnNames, removeCount, columnNames, JoinKind.Inner)
```

### SHORT NAME REF IGNOREING INITIALS THAT GIVE A FALSE POSITIVE [STRING STANDARDISATION]

```VBA
= Table.AddColumn(#"Renamed DOB", "SHORT NAME", each 
if [COUNT FIRST] < 2 then "XXXXXX" &" "& Text.Upper(List.Last(Text.Split([FULL NAME]," "))) 
else if [COUNT LAST] < 1 then Text.Upper(List.First(Text.Split([FULL NAME]," "))) &" "& "XXXXXX" 
else Text.Upper(List.First(Text.Split([FULL NAME]," "))) &" "& List.Last(Text.Split([FULL NAME]," ")))
```
# MBIE
###### [LIBRARY](https://github.com/ScottypNZ/CODE-LIBRARY)   |   [INDEX](#INDEX)
-------------------------

### NAME REF   
```VBA
let
    Source = Excel.CurrentWorkbook(){[Name="MIAS"]}[Content],
    #"Removed Other Columns" = Table.SelectColumns(Source,{"FirstName", "OtherNames", "LastName", "BirthDate", "ArrivalDate","PassportNumber"}),
    #"Added FULL NAME" = Table.AddColumn(#"Removed Other Columns", "FULL NAME", each Text.Upper(Text.Combine(List.Select({[FirstName],[OtherNames],[LastName]}, each _<> "" and _ <> null)," "))),
    #"Added FIRST" = Table.AddColumn(#"Added FULL NAME", "FIRST",
    each if [FirstName] = null then "XXXXXX" else Text.Upper(List.First(Text.Split([FirstName]," ")))),
    #"Added LAST" = Table.AddColumn(#"Added FIRST", "LAST", each if [LastName] = null then "XXXXXX" else Text.Upper(List.First(Text.Split([LastName]," ")))),
    #"Added SHORT NAME XXXXXX" = Table.AddColumn(#"Added LAST", "SHORT NAME XXXXXX",  
each if [FirstName] = null then "" else Text.Upper(List.First(Text.Split([FirstName]," ")))&" "&
Text.Upper(List.Last(Text.Split([FULL NAME]," ")))),
    #"Added SHORT NAME" = Table.AddColumn(#"Added SHORT NAME XXXXXX", "SHORT NAME", each [FIRST]&" "&[LAST]),
    #"Added NEW DOB" = Table.AddColumn(#"Added SHORT NAME", "NEW DOB", each Date.ToText(Date.From(Text.From([BirthDate])), "dd/MM/yyyy")),
    #"Added DOB SHORT NAME" = Table.AddColumn(#"Added NEW DOB", "DOB SHORT NAME", each Date.ToText(Date.From(Text.From([BirthDate])), "dd/MM/yyyy")&" "&[SHORT NAME]),
    #"Added ARRIVAL DOB SHORT NAME" = Table.AddColumn(#"Added DOB SHORT NAME", "ARRIVAL DOB SHORT NAME", each Date.ToText(Date.From(Text.From([ArrivalDate])), "dd/MM/yyyy")&" "&[DOB SHORT NAME]),
    #"Added ARRIVAL DOB PASSPORT SHORT NAME" = Table.AddColumn(#"Added ARRIVAL DOB SHORT NAME", "ARRIVAL DOB PASSPORT SHORT NAME", each [NEW DOB]&" "&Text.From([PassportNumber])&" "&[SHORT NAME])
in
    #"Added ARRIVAL DOB PASSPORT SHORT NAME"
```

### APP MAX
```VBA
let
    Source = Excel.CurrentWorkbook(){[Name="APP"]}[Content],
    #"Filtered Rows" = Table.SelectRows(Source, each ([EXCLUDE] = false)),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"DOB SHORT NAME PASSPORT REF", "DOB PASSPORT REF", "DOB SHORT NAME REF", "MovementId"}),
    #"Grouped Rows" = Table.Group(#"Removed Other Columns", {"DOB SHORT NAME PASSPORT REF", "DOB PASSPORT REF", "DOB SHORT NAME REF"},{{"MAX", each List.Max([MovementId]), type number}})
in
    #"Grouped Rows"
```

### GROUPING MATRIX

```VBA
let
    Source = Excel.CurrentWorkbook(){[Name="BHR"]}[Content],
    #"Removed Other Columns" = Table.SelectColumns(Source,{"MIAS GROUP ID", "ARRIVAL", "Email", "Bubble", "ARRIVAL ROOM REF"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"ARRIVAL", type text}, {"ARRIVAL ROOM REF", type text}, {"Email", type text}, {"Bubble", type text}, {"MIAS GROUP ID", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",null,"XXXX",Replacer.ReplaceValue,{"ARRIVAL ROOM REF", "Email", "Bubble"}),
    #"Added Email REF" = Table.AddColumn(#"Replaced Value", "EMAIL REF", each [ARRIVAL]&" "&[Email]),
    #"Added Bubble REF" = Table.AddColumn(#"Added Email REF", "BUBBLE REF", each [ARRIVAL]&" "&[Bubble]),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Added Bubble REF", {"MIAS GROUP ID", "ARRIVAL", "Email", "Bubble"}, "Attribute", "Value"),
    #"Filtered Rows" = Table.SelectRows(#"Unpivoted Columns", each ([MIAS GROUP ID] <> "")),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each not Text.Contains([Value], "XXXX")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows1",{"ARRIVAL", "Email", "Bubble", "Attribute"})
in
    #"Removed Columns"
```

_________________________________________________________________

### ROOMING
```VBA
let
    Source = #"MIAS",
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"ROOMING INDEX", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([ROOMING INDEX] = 1)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Group count", Order.Ascending},{"Group reference", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "PAX", 1, 1, Int64.Type),
    #"Removed Columns" = Table.RemoveColumns(#"Added Index",{"ROOMING INDEX", "Booking id", "Room", "ROOMING INDEX BY ROOM"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"PAX", "Hotel", "Room type", "Check in", "Check out", "First name", "Last name", "Date of birth", "Group count", "Adjoining room", "Accessibility", "Special requirements", "Dietary requirements", "Group reference", "Flight number", "Unaccompanied minor", "Speak english", "Languages spoken", "Other language", "Email", "Phone", "Primary contact", "Group id", "ROOMING LIST", "ROOMING LIST BY ROOM"}),
    #"Add blank" = if Table.IsEmpty(#"Reordered Columns") = true then Table.InsertRows(#"Reordered Columns", 0,{Record.FromList(List.Repeat({""},Table.ColumnCount(#"Reordered Columns")),Table.ColumnNames(#"Reordered Columns"))}) else #"Reordered Columns"
in
    #"Add blank"
```
### MERGE FILES BASIC

```VBA
let
    Source = Folder.Files("C:\LOCAL DATA\MIAS\MIAS DATA"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Filtered Rows" = Table.SelectRows(#"Filtered Hidden Files1", each ([Extension] = ".xlsx")),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Rows", "Transform File from MIAS", each #"Transform File from MIAS"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File from MIAS"}),
    #"Sorted Rows1" = Table.Sort(#"Removed Other Columns1",{{"Source.Name", Order.Ascending}}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Sorted Rows1", "Transform File from MIAS", Table.ColumnNames(#"Transform File from MIAS"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Authorisation checkbox", type text}, {"AuthoriseToValidateInformation", type text}, {"RawFlags", type text}, {"GroupBucket", type text}, {"PassengerBucket", type text}, {"TentativeReleaseDate", type text}, {"Registration declarations", type text}, {"Notes", type text}, {"ArrivalDate", type text}, {"FlightArrivalAirport", type text}, {"FlightArrivalTime", type text}, {"FlightDepartureCity", type text}, {"FlightDepartureTime", type text}, {"Flight", type text}, {"FlightID", type text}, {"AirlineBookingReference", type text}, {"Airline", type text}, {"NumberOfNights", type text}, {"HotelCity", type text}, {"HotelCheckOut", type text}, {"HotelCheckIn", type text}, {"HotelRoomType", type text}, {"HotelRoom", type text}, {"HotelRoomID", type text}, {"Hotel", type text}, {"BookingID", type text}, {"UnaccompaniedMinor", type text}, {"MinorHasParentOrGuardian", type text}, {"VoucherLastSent", type text}, {"VoucherID", type text}, {"DietaryRequirements", type text}, {"HotelRoomSpecialRequests", type text}, {"HotelRoomAccessibilityRequirements", type text}, {"HotelRoomPreferredTypes", type text}, {"HotelRoomAdjoining", type text}, {"PreferredNumberOfRooms", type text}, {"NumberOfPassengersInGroup", type text}, {"GroupName", type text}, {"GroupID", type text}, {"UpdatedAt", type text}, {"CreatedAt", type text}, {"UserID", type text}, {"FinalArrivalDestinationInNewZealand", type text}, {"TemporaryOrPermanentAddress", type text}, {"Contact_Postcode", type text}, {"Contact_City", type text}, {"Contact_Suburb", type text}, {"Contact_StreetName", type text}, {"Contact_ApartmentOrStreetNumber", type text}, {"KnownResidentialAddressInNewZealand", type text}, {"INZClientNumber", type text}, {"Charges-leftBeforeMarch19", type text}, {"Charges-ordinarilyOnMarch19", type text}, {"Charges-DepartAfterAug11", type text}, {"Charges-DurationOfStay", type text}, {"VisaTemporaryOtherTypeSpecified", type text}, {"VisaTemporaryOtherType", type text}, {"VisaDetail", type text}, {"VisaStatus", type text}, {"NewZealandTemporaryVisaHolder", type text}, {"AustralianCitizenOrPermanentResident", type text}, {"NewZealandPermanentResident", type text}, {"NewZealandCitizen", type text}, {"PassportDateOfExpiry", type text}, {"PassportDateOfIssue", type text}, {"PassportIssuingAuthority", type text}, {"PassportNumber", type text}, {"Phone", type text}, {"Email", type text}, {"Nationality", type text}, {"OtherLanguage", type text}, {"LanguagesSpoken", type text}, {"SpeakEnglish", type text}, {"CountryOfDeparture", type text}, {"BirthDate", type text}, {"GenderDiverse", type text}, {"Gender", type text}, {"LastName", type text}, {"OtherNames", type text}, {"FirstName", type text}, {"PrimaryContact", type text}, {"Status", type text}, {"PassengerID", type text}, {"DATE", type text}, {"MIAS DATE", type text}, {"Source.Name", type text}}),
    #"ADD Duplicate" = Table.AddColumn( #"Changed Type", "DUPLICATING", each Text.Combine({[GroupID],[FirstName],[OtherNames],[LastName],[PassportNumber],[INZClientNumber],[FlightID],[Flight],[ArrivalDate],[RawFlags]}," | "), type text),
    #"Reordered Columns" = Table.ReorderColumns(#"ADD Duplicate",{"DUPLICATING", "Source.Name", "MIAS DATE", "DATE", "PassengerID", "Status", "PrimaryContact", "FirstName", "OtherNames", "LastName", "Gender", "GenderDiverse", "BirthDate", "CountryOfDeparture", "SpeakEnglish", "LanguagesSpoken", "OtherLanguage", "Nationality", "Email", "Phone", "PassportNumber", "PassportIssuingAuthority", "PassportDateOfIssue", "PassportDateOfExpiry", "NewZealandCitizen", "NewZealandPermanentResident", "AustralianCitizenOrPermanentResident", "NewZealandTemporaryVisaHolder", "VisaStatus", "VisaDetail", "VisaTemporaryOtherType", "VisaTemporaryOtherTypeSpecified", "Charges-DurationOfStay", "Charges-DepartAfterAug11", "Charges-ordinarilyOnMarch19", "Charges-leftBeforeMarch19", "INZClientNumber", "KnownResidentialAddressInNewZealand", "Contact_ApartmentOrStreetNumber", "Contact_StreetName", "Contact_Suburb", "Contact_City", "Contact_Postcode", "TemporaryOrPermanentAddress", "FinalArrivalDestinationInNewZealand", "UserID", "CreatedAt", "UpdatedAt", "GroupID", "GroupName", "NumberOfPassengersInGroup", "PreferredNumberOfRooms", "HotelRoomAdjoining", "HotelRoomPreferredTypes", "HotelRoomAccessibilityRequirements", "HotelRoomSpecialRequests", "DietaryRequirements", "VoucherID", "VoucherLastSent", "MinorHasParentOrGuardian", "UnaccompaniedMinor", "BookingID", "Hotel", "HotelRoomID", "HotelRoom", "HotelRoomType", "HotelCheckIn", "HotelCheckOut", "HotelCity", "NumberOfNights", "Airline", "AirlineBookingReference", "FlightID", "Flight", "FlightDepartureTime", "FlightDepartureCity", "FlightArrivalTime", "FlightArrivalAirport", "ArrivalDate", "Notes", "Registration declarations", "TentativeReleaseDate",  "PassengerBucket", "GroupBucket", "RawFlags", "AuthoriseToValidateInformation", "Authorisation checkbox"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"PassengerID", Order.Ascending}})
in
    #"Sorted Rows"
  ```
### ADD REFERENCE DETAILS
```VBA
let
    Source = #"MIAS ALL SIMPLE",
    #"Removed Duplicates" = Table.Distinct(Source, {"DUPLICATING"}),
    #"Changed NAME/DATE" = Table.TransformColumnTypes(#"Removed Duplicates",
{{"ArrivalDate", type date},
{"FirstName", type text},
{"OtherNames", type text},
{"LastName", type text},
{"DATE", type date}}),
    #"Added REFERENCE" = Table.AddColumn(#"Changed NAME/DATE", "REFERENCE", each if [ArrivalDate] = null then 
Text.PadStart(Text.From([PassengerID]),6,"0")&" "&"00/00/0000"
else 
Text.PadStart(Text.From([PassengerID]),6,"0")&" "&Text.PadStart(Text.From([ArrivalDate]),10,"0")),
    #"Added MONTH" = Table.AddColumn(#"Added REFERENCE", "MONTH", each if [ArrivalDate] = null then "0000/00" else Text.From(Date.Year([ArrivalDate]))&"/"&Text.PadStart(Text.From(Date.Month([ArrivalDate])),2,"0")),
    #"Added FIRST" = Table.AddColumn(#"Added MONTH", "FIRST",
each if [FirstName] = null then "XXXXXX" else Text.Upper(List.First(Text.Split([FirstName]," ")))),
    #"Added LAST" = Table.AddColumn(#"Added FIRST", "LAST", each if [LastName] = null then "XXXXXX" else Text.Upper(List.First(Text.Split([LastName]," ")))),
    #"Added SHORT NAME" = Table.AddColumn(#"Added LAST", "SHORT NAME", each [FIRST]&" "&[LAST]),
    #"Split DATE" = Table.SplitColumn(Table.TransformColumnTypes(#"Added SHORT NAME", {{"BirthDate", type text}}, "en-NZ"), "BirthDate", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"BirthDate.1", "BirthDate.2"}),
    #"Added BIRTHDATE" = Table.AddColumn(#"Split DATE", "BirthDate.3", each [BirthDate.1]),
    #"Replaced DASH" = Table.ReplaceValue(#"Added BIRTHDATE","-","/",Replacer.ReplaceText,{"BirthDate.3"}),
    #"Added NUMBER" = Table.AddColumn(#"Replaced DASH", "NUMBER", each Number.From(Date.From([BirthDate.3]))),
    #"Added DOB" = Table.AddColumn(#"Added NUMBER", "DOB", each Text.PadStart(Text.From(Date.From([NUMBER])),10,"0")),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Added DOB", {{"DOB", "00/00/0000"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Replaced Errors",null,"00/00/0000",Replacer.ReplaceValue,{"DOB"}),
    #"Added NAME-REF" = Table.AddColumn(#"Replaced Value", "NAME-REF", 
each if [DOB] = null 
then "00/00/0000"&" "&[FIRST]&" "&[LAST] 
else [DOB]&" "&[FIRST]&" "&[LAST]),
    #"Added ARRIVAL" = Table.AddColumn(#"Added NAME-REF", "ARRIVAL", each Text.PadStart(Text.From([ArrivalDate]),10,"0")),
    #"Replaced NULL" = Table.ReplaceValue(#"Added ARRIVAL",null,"00/00/0000",Replacer.ReplaceValue,{"ARRIVAL"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Replaced NULL",{"MIAS DATE", "DATE", "PassengerID", "Status", "FirstName", "LastName", "PassportNumber", "UserID", "GroupID", "GroupName", "NumberOfPassengersInGroup", "PreferredNumberOfRooms", "VoucherID", "RawFlags", "REFERENCE", "MONTH", "SHORT NAME", "DOB", "NAME-REF", "ARRIVAL"}),
    #"Sorted Rows1" = Table.Sort(#"Removed Other Columns",{{"MIAS DATE", Order.Descending}})
in
    #"Sorted Rows1"
```


### MERGE EXCEL FILES AND PROCESS
```VBA
let
    Source = Folder.Files("C:\LOCAL DATA\MIAS\DATA"),
    #"Filtered Hidden Files" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function" = Table.AddColumn(#"Filtered Hidden Files", "Transform File from MIAS", each #"Transform File from MIAS"([Content])),
    #"Sorted File" = Table.Sort(#"Invoke Custom Function",{{"Name", Order.Descending}}),
    #"Expanded Transform MIAS" = Table.ExpandTableColumn(#"Sorted File", "Transform File from MIAS", {"MIAS DATE", "DATE", "PassengerID", "Status", "PrimaryContact", "FirstName", "OtherNames", "LastName", "Gender", "GenderDiverse", "BirthDate", "CountryOfDeparture", "SpeakEnglish", "LanguagesSpoken", "OtherLanguage", "Nationality", "Email", "NZ contact number", "NZ contact number value", "Other contact number", "PassportNumber", "PassportIssuingAuthority", "PassportDateOfIssue", "PassportDateOfExpiry", "NewZealandCitizen", "NewZealandPermanentResident", "AustralianCitizenOrPermanentResident", "NewZealandTemporaryVisaHolder", "VisaStatus", "VisaDetail", "VisaTemporaryOtherType", "VisaTemporaryOtherTypeSpecified", "Charges-DurationOfStay", "Charges-DepartAfterAug11", "Charges-ordinarilyOnMarch19", "Charges-leftBeforeMarch19", "INZClientNumber", "KnownResidentialAddressInNewZealand", "Contact_ApartmentOrStreetNumber", "Contact_StreetName", "Contact_Suburb", "Contact_City", "Contact_Postcode", "TemporaryOrPermanentAddress", "FinalArrivalDestinationInNewZealand", "UserID", "AllowedDuplicateVouchers", "CreatedAt", "UpdatedAt", "GroupID", "GroupName", "NumberOfPassengersInGroup", "PreferredNumberOfRooms", "HotelRoomAdjoining", "HotelRoomPreferredTypes", "HotelRoomAccessibilityRequirements", "HotelRoomSpecialRequests", "DietaryRequirements", "VoucherID", "VoucherLastSent", "MinorHasParentOrGuardian", "UnaccompaniedMinor", "BookingID", "Hotel", "HotelRoomID", "HotelRoom", "HotelRoomType", "HotelCheckIn", "HotelCheckOut", "HotelCity", "NumberOfNights", "Airline", "AirlineBookingReference", "FlightID", "Flight", "FlightDepartureTime", "FlightDepartureCity", "FlightArrivalTime", "FlightArrivalAirport", "ArrivalDate", "Notes", "Registration declarations", "TentativeReleaseDate", "PassengerBucket", "GroupBucket", "RawFlags", "BulkUploadNotes", "Vaccinated", "AuthoriseToValidateInformation", "Authorisation checkbox", "ScanTransportID", "ScanTransportCity", "ScanTransportHotel", "ScanTransportOperator", "ScanTransportTime", "ScanTransportInteractionTime", "ScanTransportInteractionType", "ScanHotel", "ScanHotelInteractionTime", "ScanHotelInteractionType", "ScanHotelOperator", "ScanNumberOfRooms", "ScanRoomNumbers"}, {"MIAS DATE", "DATE", "PassengerID", "Status", "PrimaryContact", "FirstName", "OtherNames", "LastName", "Gender", "GenderDiverse", "BirthDate", "CountryOfDeparture", "SpeakEnglish", "LanguagesSpoken", "OtherLanguage", "Nationality", "Email", "NZ contact number", "NZ contact number value", "Other contact number", "PassportNumber", "PassportIssuingAuthority", "PassportDateOfIssue", "PassportDateOfExpiry", "NewZealandCitizen", "NewZealandPermanentResident", "AustralianCitizenOrPermanentResident", "NewZealandTemporaryVisaHolder", "VisaStatus", "VisaDetail", "VisaTemporaryOtherType", "VisaTemporaryOtherTypeSpecified", "Charges-DurationOfStay", "Charges-DepartAfterAug11", "Charges-ordinarilyOnMarch19", "Charges-leftBeforeMarch19", "INZClientNumber", "KnownResidentialAddressInNewZealand", "Contact_ApartmentOrStreetNumber", "Contact_StreetName", "Contact_Suburb", "Contact_City", "Contact_Postcode", "TemporaryOrPermanentAddress", "FinalArrivalDestinationInNewZealand", "UserID", "AllowedDuplicateVouchers", "CreatedAt", "UpdatedAt", "GroupID", "GroupName", "NumberOfPassengersInGroup", "PreferredNumberOfRooms", "HotelRoomAdjoining", "HotelRoomPreferredTypes", "HotelRoomAccessibilityRequirements", "HotelRoomSpecialRequests", "DietaryRequirements", "VoucherID", "VoucherLastSent", "MinorHasParentOrGuardian", "UnaccompaniedMinor", "BookingID", "Hotel", "HotelRoomID", "HotelRoom", "HotelRoomType", "HotelCheckIn", "HotelCheckOut", "HotelCity", "NumberOfNights", "Airline", "AirlineBookingReference", "FlightID", "Flight", "FlightDepartureTime", "FlightDepartureCity", "FlightArrivalTime", "FlightArrivalAirport", "ArrivalDate", "Notes", "Registration declarations", "TentativeReleaseDate", "PassengerBucket", "GroupBucket", "RawFlags", "BulkUploadNotes", "Vaccinated", "AuthoriseToValidateInformation", "Authorisation checkbox", "ScanTransportID", "ScanTransportCity", "ScanTransportHotel", "ScanTransportOperator", "ScanTransportTime", "ScanTransportInteractionTime", "ScanTransportInteractionType", "ScanHotel", "ScanHotelInteractionTime", "ScanHotelInteractionType", "ScanHotelOperator", "ScanNumberOfRooms", "ScanRoomNumbers"}),
    #"Removed Non core" = Table.SelectColumns(#"Expanded Transform MIAS",{"MIAS DATE", "PassengerID", "PrimaryContact", "FirstName", "OtherNames", "LastName", "BirthDate", "PassportNumber", "INZClientNumber", "UserID", "GroupID", "GroupName", "NumberOfPassengersInGroup", "PreferredNumberOfRooms", "VoucherID", "BookingID", "HotelRoom", "Flight", "ArrivalDate", "RawFlags"}),
    #"Change all to Text" = Table.TransformColumnTypes(#"Removed Non core",{{"MIAS DATE", type text}, {"PassengerID", type text}, {"FirstName", type text}, {"OtherNames", type text}, {"LastName", type text}, {"BirthDate", type text}, {"PassportNumber", type text}, {"INZClientNumber", type text}, {"UserID", type text}, {"GroupID", type text}, {"GroupName", type text}, {"NumberOfPassengersInGroup", type text}, {"PreferredNumberOfRooms", type text}, {"VoucherID", type text}, {"BookingID", type text}, {"HotelRoom", type text}, {"Flight", type text}, {"ArrivalDate", type text}, {"RawFlags", type text}}),
    #"Added COMBINED" = Table.AddColumn(#"Change all to Text", "COMBINED", each Text.Combine({[VoucherID], [PassengerID],[FirstName],[OtherNames],[LastName],[BirthDate],[PassportNumber],[Flight],[ArrivalDate],[RawFlags]}," ")),
    #"Removed Duplicates" = Table.Distinct(#"Added COMBINED", {"COMBINED"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Duplicates",{{"MIAS DATE", type date},{"ArrivalDate", type date}, {"PrimaryContact", Int64.Type}}),
    #"Added MONTH" = Table.AddColumn(#"Changed Type", "MONTH", each if [ArrivalDate] = null then "XXXX/XX" else Text.From(Date.Year([ArrivalDate]))&"/"&Text.PadStart(Text.From(Date.Month([ArrivalDate])),2,"0")),
    #"Merged FULL NAME" = Table.AddColumn(#"Added MONTH", "FULL NAME", each Text.Upper(Text.Combine({[FirstName], [OtherNames], [LastName]}, " ")), type text),
    #"Added AllWordCount" = Table.AddColumn(#"Merged FULL NAME", "AllWordCount", each List.Count(List.RemoveItems(Text.SplitAny([FULL NAME]," #(tab)#(lf)"),{""}))),
    #"Added LastCount" = Table.AddColumn(#"Added AllWordCount", "LastCount", each try List.Count(Text.SplitAny([LastName]," ")) otherwise 0),
    #"Added SHORT NAME" = Table.AddColumn(#"Added LastCount", "SHORT NAME", 
each if [AllWordCount] = 1 and [LastCount] = 1 then "XXXXXXX" &" "&Text.Upper(List.Last(Text.Split([FULL NAME]," "))) 
else if [AllWordCount] = 1 and [LastCount] = 0 then Text.Upper(List.First(Text.Split([FULL NAME]," "))) &" "& "XXXXXXX"
else if [AllWordCount] = 0 then "XXXXXXX" &" "& "XXXXXXX"
else Text.Upper(List.First(Text.Split([FULL NAME]," "))) &" "& Text.Upper(List.Last(Text.Split([FULL NAME]," ")))),
    #"Split DATE" = Table.SplitColumn(Table.TransformColumnTypes(#"Added SHORT NAME", {{"BirthDate", type text}}, "en-NZ"), "BirthDate", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"BirthDate", "BirthDate.2"}),
    #"Added BIRTHDATE" = Table.AddColumn(#"Split DATE", "BirthDate.3", each [BirthDate]),
    #"Replaced DASH" = Table.ReplaceValue(#"Added BIRTHDATE","-","/",Replacer.ReplaceText,{"BirthDate.3"}),
    #"Added NUMBER" = Table.AddColumn(#"Replaced DASH", "NUMBER", each Number.From(Date.From([BirthDate.3]))),
    #"Added DOB" = Table.AddColumn(#"Added NUMBER", "DOB", each Text.PadStart(Text.From(Date.From([NUMBER])),10,"0")),
    #"Replaced Errors" = Table.ReplaceErrorValues(#"Added DOB", {{"DOB", "XX/XX/XXXX"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Replaced Errors",null,"XX/XX/XXXX",Replacer.ReplaceValue,{"DOB"}),
    #"Added NAME-REF" = Table.AddColumn(#"Replaced Value", "NAME-REF", 
each if [DOB] = null 
then "XX/XX/XXXX"&" "& [SHORT NAME] else [DOB]&" "&[SHORT NAME]),
    #"Added ARRIVAL" = Table.AddColumn(#"Added NAME-REF", "ARRIVAL", each Text.PadStart(Text.From([ArrivalDate]),10,"0")),
    #"Replaced NULL" = Table.ReplaceValue(#"Added ARRIVAL",null,"XX/XX/XXXX",Replacer.ReplaceValue,{"ARRIVAL"}),
    #"Added REFERENCE" = Table.AddColumn(#"Replaced NULL", "REFERENCE", each 
if [ArrivalDate] = null
then Text.PadStart(Text.From([PassengerID]),6,"0") &" "& "XX/XX/XXXX" &" "& [#"NAME-REF"]
else Text.PadStart(Text.From([PassengerID]),6,"0") &" "& Text.PadStart(Text.From([ArrivalDate]),10,"0") &" "& [#"NAME-REF"]),
    #"Added VALID" = Table.AddColumn(#"Added REFERENCE", "INVALID", each [ArrivalDate]=null or [Flight] =null),
    #"Grouped Rows" = Table.Group(#"Added VALID", {"PassengerID"}, {{"MAX", each List.Max([MIAS DATE]), type date}, {"ALL", each _, type table }}),
    #"Expanded ALL" = Table.ExpandTableColumn(#"Grouped Rows", "ALL", {"COMBINED", "PrimaryContact", "MIAS DATE", "FirstName", "OtherNames", "LastName", "BirthDate",  "PassportNumber", "INZClientNumber", "UserID", "GroupID", "GroupName", "NumberOfPassengersInGroup", "PreferredNumberOfRooms", "VoucherID", "BookingID", "HotelRoom", "Flight", "ArrivalDate", "RawFlags", "MONTH", "FULL NAME", "AllWordCount", "LastCount", "SHORT NAME", "BirthDate.2","BirthDate.3", "NUMBER", "DOB", "NAME-REF", "ARRIVAL", "REFERENCE", "INVALID"}),
    #"Added MAX ID" = Table.AddColumn(#"Expanded ALL", "MAX ID", each [MAX]=[MIAS DATE])
in
    #"Added MAX ID"
 ```
 ROOM UTILISATION
 ```VBA
 let
    Source = Excel.CurrentWorkbook(){[Name="BHR_TBL"]}[Content],
    #"Removed other columns" = Table.SelectColumns(Source,{"TYPE", "City", "HOTEL", "ROOM REFERENCE", "PEOPLE", "DAYS", "BHR", "START", "END", "ROOMS"}),
    #"Added START" = Table.AddColumn(#"Removed other columns", "Custom", each Text.PadStart(Text.Trim(Text.Start(Text.From([START]),10)),10,"0"), type date),
    #"Changed Type" = Table.TransformColumnTypes(#"Added START",{{"START", type date}, {"END", type date}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"START"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Removed Columns",{"City", "HOTEL", "ROOM REFERENCE", "PEOPLE", "DAYS", "BHR", "Custom", "END", "ROOMS"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns1",{{"Custom", "START"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "Custom", each Text.PadStart(Text.Trim(Text.Start(Text.From([END]),10)),10,"0"), type date),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom1",{"END"}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Removed Columns1",{"City", "HOTEL", "ROOM REFERENCE", "PEOPLE", "DAYS", "BHR", "START", "Custom", "ROOMS"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Reordered Columns2",{{"Custom", "END"}}),
    #"Changed Type all" = Table.TransformColumnTypes(#"Renamed Columns1",{{"HOTEL", type text}, {"ROOM REFERENCE", type text}, {"PEOPLE", Int64.Type}, {"DAYS", Int64.Type}, {"BHR", type text}, {"START", type date}, {"END", type date}, {"ROOMS", type text}, {"City", type text}}),
    #"Replaced null rooms" = Table.ReplaceValue(#"Changed Type all",null,"0000",Replacer.ReplaceValue,{"ROOMS"}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Replaced null rooms", {"City"}),
    #"Added Joined" = Table.AddColumn(#"Removed Errors", "JOINED", each try
if Text.Contains([#"ROOM NUMBER"],"-") then "TRUE" else
if Text.Contains([#"ROOM NUMBER"],"/") then "TRUE" else "FALSE"
otherwise "FALSE"),
    #"Added RoomA" = Table.AddColumn(#"Added Joined", "ROOMA", each Text.Trim(Text.Remove( [ROOMS], { "A".."z","&"..")"} ))),
    #"Reordered Columns" = Table.ReorderColumns(#"Added RoomA",{"JOINED", "ROOM REFERENCE", "PEOPLE", "DAYS", "BHR", "START", "END", "ROOMS"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"ROOM REFERENCE", Order.Descending}}),
    #"Replaced 266267268" = Table.ReplaceValue(#"Sorted Rows","266267268","266 267 268",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced Phone" = Table.ReplaceValue(#"Replaced 266267268","11080272026065","1108",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced Phonex" = Table.ReplaceValue(#"Replaced Phone","11080272026065","x11080272026065",Replacer.ReplaceText,{"ROOMS"}),
    #"Replaced Comma" = Table.ReplaceValue(#"Replaced Phonex",","," ",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced Dash" = Table.ReplaceValue(#"Replaced Comma","-"," ",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced Slash" = Table.ReplaceValue(#"Replaced Dash","/"," ",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced BackSlash" = Table.ReplaceValue(#"Replaced Slash","\"," ",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced +" = Table.ReplaceValue(#"Replaced BackSlash","+"," ",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced Negative" = Table.ReplaceValue(#"Replaced +","0909   1","0909",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced Text" = Table.TransformColumns(#"Replaced Negative",{{"ROOMA", Text.Trim, type text}}),
    #"Replaced 08163" = Table.ReplaceValue(#"Replaced Text","08163","816",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced 18111" = Table.ReplaceValue(#"Replaced 08163","1810 18111","1810 1811",Replacer.ReplaceText,{"ROOMA"}),
    #"Replaced 1101/2" = Table.ReplaceValue(#"Replaced 18111","1101  2","1101  1102",Replacer.ReplaceText,{"ROOMA"}),
    #"Added RoomB" = Table.AddColumn(#"Replaced 1101/2", "ROOMB", 
each if [ROOM REFERENCE] = "12 | 04 | 04"
then Text.Start([#"ROOMA"],3) 
&" "& Text.Start([#"ROOMA"],1) & Text.Middle([#"ROOMA"],4,2)
&" "& Text.Start([#"ROOMA"],1) & Text.Middle([#"ROOMA"],7,2)
&" "& Text.Start([#"ROOMA"],1) & Text.Middle([#"ROOMA"],10,2)
else [#"ROOMA"]),
    #"Added RoomC" = Table.AddColumn(#"Added RoomB", "ROOMC", 
each if [ROOM REFERENCE] = "10 | 03 | 05"
then Text.Start([#"ROOMA"],4) 
&" "& Text.Start([#"ROOMA"],2) & Text.Middle([#"ROOMA"],5,2)
&" "& Text.Start([#"ROOMA"],2) & Text.Middle([#"ROOMA"],8,2)
else [#"ROOMB"]),
    #"Added RoomD" = Table.AddColumn(#"Added RoomC", "ROOMD", 
each if [ROOM REFERENCE] = "09 | 04 | 04"
then Text.Start([#"ROOMA"],3) 
&" "& Text.Start([#"ROOMA"],2) & Text.Middle([#"ROOMA"],4,1)
&" "& Text.Start([#"ROOMA"],2) & Text.Middle([#"ROOMA"],6,1)
&" "& Text.Start([#"ROOMA"],2) & Text.Middle([#"ROOMA"],8,1)
else [#"ROOMC"]),
    #"Added RoomE" = Table.AddColumn(#"Added RoomD", "ROOME", 
each if [ROOM REFERENCE] = "09 | 03 | 04"
then Text.Start([#"ROOMA"],3) 
&" "& Text.Middle([#"ROOMA"],4,3)
&" "& Text.Middle([#"ROOMA"],4,2) & Text.Middle([#"ROOMA"],8,1)
else [#"ROOMD"]),
    #"Added RoomF" = Table.AddColumn(#"Added RoomE", "ROOMF", 
each if [ROOM REFERENCE] = "07 | 02 | 05"
then Text.Start([#"ROOMA"],4) 
&" "& Text.Start([#"ROOMA"],2)&Text.Middle([#"ROOMA"],5,2)
else [#"ROOME"]),
    #"Added RoomG" = Table.AddColumn(#"Added RoomF", "ROOMG", 
each if [ROOM REFERENCE] = "06 | 02 | 05"
then Text.Start([#"ROOMA"],4) 
&" "& Text.Start([#"ROOMA"],3)&Text.Middle([#"ROOMA"],5,1)
else [#"ROOMF"]),
    #"Added RoomH" = Table.AddColumn(#"Added RoomG", "ROOMH", 
each if [ROOM REFERENCE] = "05 | 02 | 04"
then Text.Start([#"ROOMA"],3) 
&" "& Text.Start([#"ROOMA"],2)&Text.Middle([#"ROOMA"],4,1)
else [#"ROOMG"]),
    #"Added Length" = Table.AddColumn(#"Added RoomH", "LENGTH", each Text.Length([ROOMH])),
    #"Added RoomI" = Table.AddColumn(#"Added Length", "ROOMI", 
each if [ROOM REFERENCE] = "06 | 01 | 00" and [LENGTH] = 6
then Text.Start([#"ROOMA"],3) 
&" "& Text.Middle([#"ROOMA"],3,3)
else [#"ROOMH"]),
    #"Removed Rooms" = Table.RemoveColumns(#"Added RoomI",{"ROOMA", "ROOMB", "ROOMC", "ROOMD", "ROOME", "ROOMF", "ROOMG", "ROOMH", "LENGTH"}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Removed Rooms", "ROOMI", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"ROOMI.1", "ROOMI.2", "ROOMI.3", "ROOMI.4", "ROOMI.5", "ROOMI.6", "ROOMI.7", "ROOMI.8"}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Split Column by Delimiter", {"TYPE", "City", "JOINED", "HOTEL", "ROOM REFERENCE", "PEOPLE", "DAYS", "BHR", "START", "END", "ROOMS"}, "Attribute", "Value"),
    #"Added Room End" = Table.AddColumn(#"Unpivoted Columns", "ROOM END", each Text.PadStart(Text.From([Value]),4,"0")),
    #"Added MIN" = Table.AddColumn(#"Added Room End", "MIN", each List.Min({Date.From([START]),Date.From([END])}), type date),
    #"Added MAX" = Table.AddColumn(#"Added MIN", "MAX", each List.Max({Date.From([START]),Date.From([END])}), type date),
    #"Removed Old Dates" = Table.RemoveColumns(#"Added MAX",{"START", "END"}),
    #"Renamed New Dates" = Table.RenameColumns(#"Removed Old Dates",{{"MIN", "START"}, {"MAX", "END"}}),

    #"Added Flat dates" = Table.AddColumn(#"Renamed New Dates", "DATE", each { Number.From([START])..Number.From([END])}),
    #"Expanded DATE" = Table.ExpandListColumn(#"Added Flat dates", "DATE"),
    #"Changed Date Type" = Table.TransformColumnTypes(#"Expanded DATE",{{"DATE", type date}}),
    #"Removed Start/End/Value" = Table.RemoveColumns(#"Changed Date Type",{"START", "END", "Value"}),
    #"Added WEEK" = Table.AddColumn(#"Removed Start/End/Value", "WEEK", each Date.ToText([DATE],"yyyy")&"/"&Text.PadStart(Text.From(Date.WeekOfYear([DATE])),2,"0"), type text),
    #"Added MONTH" = Table.AddColumn(#"Added WEEK", "MONTH", each Date.ToText([DATE],"yyyy") &"/"& Text.PadStart(Text.From(Date.Month([DATE])),2,"0"), type text)

/* = Text.PadStart(Text.From(Date.Month([ArrivalDate]),2,"0"))) */,
    #"Added REFERENCE" = Table.AddColumn(#"Added MONTH", "REFERENCE", each 
Text.PadStart(Text.From([DATE]),10,"0")&" | "& Text.From([HOTEL])&" | "& Text.From([ROOM END]), type text),
    #"Added SHORT" = Table.AddColumn(#"Added REFERENCE", "SHORT", each if [DAYS] < 13 then  "TRUE" else"FALSE"),
    #"Added ROOM REF" = Table.AddColumn(#"Added SHORT", "ROOM REF", each Text.From([HOTEL])&" | "& Text.Start([ROOM END],4), type text),
    #"Removed Duplicates" = Table.Distinct(#"Added ROOM REF", {"REFERENCE"}),
    #"Added SUB" = Table.AddColumn(#"Removed Duplicates", "ROW", each "SUB", type text),
    #"Added MIQ" = Table.AddColumn(#"Added SUB", "OFFICE", each "MIQ", type text)
in
    #"Added MIQ"
```
# API   
___________________________________________


JSON REPORT LIST
```VBA
let
    Source = Json.Document(Web.Contents("http://mnz-16-app02-p/MaritimeNewZealand/api/data/v8.0/")), value = Source[value],   
    #"Converted to Table" = Table.FromList(value, Splitter.SplitByNothing(), null, null, ExtraValues.Error),   
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"name", "kind", "url"}, {"Column1.name", "Column1.kind", "Column1.url"})
in
    #"Expanded Column1"
```    

# MARITIME
###### [LIBRARY](https://github.com/ScottypNZ/CODE-LIBRARY)   |   [INDEX](#INDEX)
-------------------------

ODATA API CALL - LIST AND JOIN METHOD (TO USE LESS MEMORY, CHANGE NOT NOTICEABLE)
 ```VBA
let
    Source = Table.ColumnNames(OData.Feed("https://crm/maritimenewzealand/api/data/v8.0/contacts")),
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Filtered Rows1" = Table.SelectRows(#"Converted to Table", each (
[Column1] = "createdon" or 
[Column1] = "modifiedon" or
[Column1] = "mnz_alert" or
[Column1] = "mnz_customernumber" or
[Column1] = "mnz_tritonpersonid" or
[Column1] = "mnz_tradingas"  or
[Column1] = "birthdate" or
[Column1] = "fullname" or
[Column1] = "firstname" or
[Column1] = "middlename" or
[Column1] = "lastname" or
[Column1] = "emailaddress1" or
[Column1] = "address1_line1" or
[Column1] = "address1_line2" or
[Column1] = "address1_line3" or
[Column1] = "address1_county" or
[Column1] = "address1_city" or
[Column1] = "address1_postalcode" or
[Column1] = "mobilephone" or
[Column1] = "mnz_otherphone" or
[Column1] = "mnz_phones" or
[Column1] = "telephone1" or
[Column1] = "telephone2" or
[Column1] = "mnz_custmnzsolomonid" or
[Column1] = "mnz_custopfsolomonid" or
[Column1] = "mnz_custsarsolomonid" or
[Column1] = "createdby" 
)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows1", "Custom", each null),
    #"Pivoted Column" = Table.Pivot(#"Added Custom", List.Distinct(#"Added Custom"[Column1]), "Column1", "Custom"),
    #"Reordered Columns" = Table.ReorderColumns(#"Pivoted Column",{"createdon", "modifiedon", "mnz_alert", "mnz_customernumber", "mnz_tritonpersonid", "mnz_tradingas", "birthdate", "fullname", "firstname", "middlename", "lastname", "emailaddress1", "address1_line1", "address1_line2", "address1_line3", "address1_county", "address1_city", "address1_postalcode", "mobilephone", "mnz_otherphone", "mnz_phones", "telephone1", "telephone2", "mnz_custmnzsolomonid", "mnz_custopfsolomonid", "mnz_custsarsolomonid", "createdby"}),
    Join = Table.NestedJoin(#"Reordered Columns", {"mnz_customernumber"}, OData.Feed("https://crm/maritimenewzealand/api/data/v8.0/contacts"), {"mnz_customernumber"}, "test", JoinKind.FullOuter)
in
    Join
```

JSON API CALL - ALL IN ONE STEP

 ```VBA
= Table.TransformColumns(Table.ExpandRecordColumn(Table.FromList(Json.Document(Web.Contents(
"https://crm/maritimenewzealand/api/data/v8.0/contacts")) [value] ,   
Splitter.SplitByNothing(), null, null, ExtraValues.Error),    
"Column1", {"createdon", "modifiedon", "mnz_customernumber"}, 
 {{"createdon", each Date.From(DateTimeZone.From(_)), type date}, {"modifiedon", each Date.From(DateTimeZone.From(_)), type date}})
```


XML API CALL (VIA JSON FUNCTION) - MULTISTEP
 ```VBA
let
    Source = Json.Document(Web.Contents("https://crm/maritimenewzealand/api/data/v8.0/contacts?fetchXml=%3Cfetch+version%3D%221.0%22+output-format%3D%22xml-platform%22+mapping%3D%22logical%22+distinct%3D%22false%22%3E%0D%0A++%3Centity+name%3D%22contact%22%3E%0D%0A++++%3Cattribute+name%3D%22fullname%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22statecode%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22emailaddress1%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_customernumber%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_alert%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22birthdate%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22address1_line3%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22address1_postalcode%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22address1_line2%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22address1_country%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22address1_city%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22address1_line1%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_tradingas%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22contactid%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_tritonpersonid%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22modifiedon%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22createdon%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22createdby%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_otherphone%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mobilephone%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22telephone2%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22telephone1%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_custsarsolomonid%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_custopfsolomonid%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_custmnzsolomonid%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_address1city%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22mnz_address1city%22+%2F%3E%0D%0A
%3Cattribute+name%3D%22firstname%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22middlename%22+%2F%3E%0D%0A++++%3Cattribute+name%3D%22lastname%22+%2F%3E%0D%0A
++++%3Corder+attribute%3D%22fullname%22+descending%3D%22false%22+%2F%3E%0D%0A++%3C%2Fentity%3E%0D%0A%3C%2Ffetch%3E")),
    #"Converted to Table" = Table.FromRecords({Source}),
    #"Removed Columns" = Table.RemoveColumns(#"Converted to Table",{"@odata.context", "@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"}),
    #"Expanded List" = Table.ExpandListColumn(#"Removed Columns", "value"),
    #"Expanded value" = Table.ExpandRecordColumn(#"Expanded List", "value", {"createdon", "modifiedon",  "createdby", "mnz_alert", "mnz_tradingas", "mnz_customernumber", "mnz_tritonpersonid", "fullname", "firstname", "middlename", "lastname", "birthdate", "emailaddress1", "address1_composite", "mobilephone", "telephone1", "telephone2", "mnz_otherphone", "mnz_custmnzsolomonid", "mnz_custsarsolomonid", "mnz_custopfsolomonid"}),
    #"Added Index" = Table.AddIndexColumn(#"Expanded value", "Index", 1, 1),
    #"Reordered INDEX" = Table.ReorderColumns(#"Added Index",{"Index", "createdon", "modifiedon", "createdby", "mnz_alert", "mnz_tradingas", "mnz_customernumber", "mnz_tritonpersonid", "fullname", "firstname", "middlename", "lastname", "birthdate", "emailaddress1", "address1_composite", "mobilephone", "telephone1", "telephone2", "mnz_otherphone", "mnz_custmnzsolomonid", "mnz_custsarsolomonid", "mnz_custopfsolomonid"}),
    #"Replaced Value" = Table.ReplaceValue(#"Reordered INDEX","#(cr)#(lf)",", ",Replacer.ReplaceText,{"address1_composite"}),
    #"Parsed Date" = Table.TransformColumns(#"Replaced Value",{{"createdon", each Date.From(DateTimeZone.From(_)), type date}, {"modifiedon", each Date.From(DateTimeZone.From(_)), type date}})
in
    #"Parsed Date"
```

### JSON API ODATA PAGENATION LOOP - PRE FILTERED DATA 

GetData as list
 ```VBA

(Page as number) as list =>
let StartURL = "https://crm/maritimenewzealand/api/data/v8.0/contacts?$skiptoken=%3Ccookie%20pagenumber=%22",
    EndURL   = "%22%20pagingcookie=%22%253ccookie%2520page%253d%25221%2522%253e%253ccontactid%2520last%253d%2522%257bD0E5305F-0085-E211-9FD1-000C29854771%257d%2522%2520first%253d%2522%257b609AF16C-120C-4E2C-9498-00015D9B0068%257d%2522%2520%252f%253e%253c%252fcookie%253e%22%20istracking=%22False%22%20/%3E&$select=createdon,modifiedon,mnz_alert,mnz_customernumber,mnz_tritonpersonid,mnz_tradingas,birthdate,fullname,firstname,middlename,lastname,emailaddress1,address1_line1,address1_line2,address1_line3,address1_county,address1_city,address1_postalcode,mobilephone,mnz_otherphone,mnz_phones,telephone1,telephone2,mnz_custmnzsolomonid,mnz_custopfsolomonid,mnz_custsarsolomonid",
    source =Record.ToTable(Json.Document(Web.Contents( StartURL & Number.ToText(Page) & EndURL))) {1} [Value]
in  source
```
GetData as table from list
 ```VBA
(Pages as number) as table =>
let StartURL = "https://crm/maritimenewzealand/api/data/v8.0/contacts?$skiptoken=%3Ccookie%20pagenumber=%22",
    EndURL   = "%22%20pagingcookie=%22%253ccookie%2520page%253d%25221%2522%253e%253ccontactid%2520last%253d%2522%257bD0E5305F-0085-E211-9FD1-000C29854771%257d%2522%2520first%253d%2522%257b609AF16C-120C-4E2C-9498-00015D9B0068%257d%2522%2520%252f%253e%253c%252fcookie%253e%22%20istracking=%22False%22%20/%3E&$select=createdon,modifiedon,mnz_alert,mnz_customernumber,mnz_tritonpersonid,mnz_tradingas,birthdate,fullname,firstname,middlename,lastname,emailaddress1,address1_line1,address1_line2,address1_line3,address1_county,address1_city,address1_postalcode,mobilephone,mnz_otherphone,mnz_phones,telephone1,telephone2,mnz_custmnzsolomonid,mnz_custopfsolomonid,mnz_custsarsolomonid",
    Source =Table.FromList(Json.Document(Web.Contents( StartURL & Number.ToText(Pages) & EndURL)) [value] , Splitter.SplitByNothing(), null, null, ExtraValues.Error)
in  Source
```

GetData as table from records
 ```VBA
(Page1 as number)  =>
let StartURL = "https://crm/maritimenewzealand/api/data/v8.0/contacts?$skiptoken=%3Ccookie%20pagenumber=%22",
    EndURL   = "%22%20pagingcookie=%22%253ccookie%2520page%253d%25221%2522%253e%253ccontactid%2520last%253d%2522%257bD0E5305F-0085-E211-9FD1-000C29854771%257d%2522%2520first%253d%2522%257b609AF16C-120C-4E2C-9498-00015D9B0068%257d%2522%2520%252f%253e%253c%252fcookie%253e%22%20istracking=%22False%22%20/%3E&$select=createdon,modifiedon,mnz_alert,mnz_customernumber,mnz_tritonpersonid,mnz_tradingas,birthdate,fullname,firstname,middlename,lastname,emailaddress1,address1_line1,address1_line2,address1_line3,address1_county,address1_city,address1_postalcode,mobilephone,mnz_otherphone,mnz_phones,telephone1,telephone2,mnz_custmnzsolomonid,mnz_custopfsolomonid,mnz_custsarsolomonid",
    Source =Table.FromRecords(Json.Document(Web.Contents( StartURL & Number.ToText(Page1) & EndURL)) [value] )
in  Source
```

List Generate Varients

GetDataListRecordToTable   
GetDataTableFromList   
GetDataTableFromRecords   



 ```VBA
let ListGenerate = List.Generate( () =>
        [Result= try GetDataListRecordToTable(1) otherwise null, Page = 1], 
        each not List.IsEmpty([Result]),
        each [Result = try GetDataListRecordToTable([Page]+1) otherwise null, 
        Page =[Page]+1], each [Result]),
    ConvertedTable = Table.FromList(ListGenerate, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    ExpandedToList = Table.ExpandListColumn(ConvertedTable, "Column1"),
    ExpandedToRecords = Table.ExpandRecordColumn(ExpandedToList, "Column1", {"createdon", "modifiedon", "mnz_alert", "mnz_customernumber", "mnz_tritonpersonid", "mnz_tradingas", "birthdate", "fullname", "firstname", "middlename", "lastname", "emailaddress1", "address1_line1", "address1_line2", "address1_line3", "address1_county", "address1_city", "address1_postalcode",  "address1_composite", "mobilephone", "mnz_otherphone", "mnz_phones", "telephone1", "telephone2", "mnz_custmnzsolomonid", "mnz_custopfsolomonid", "mnz_custsarsolomonid"})
in  ExpandedToRecords
```

### APPLICATION EFFICIENCY
###### STEP 1 | Create decision column giving status (approved etc), remove other columns, group min associated date, merge back
###### STEP 2 | Create ready column giving admin ready (if received different), remove other columns, group min associated date, merge back
###### STEP 3 | Grouped MAX to give final modified date. You then have the dates for received, ready to process, approved, closed
###### STEP 4 | Advisor Pivot, remove other, group / count advisors, creating occurance, merge back, then pivoting horizontally

### STEP 1
```VBA
let source = Table1

  #"---DESCISION---" = source,

  #"Added DECISION" = Table.AddColumn(#"---DESCISION---", "DECISION", each [status] = "Application Approved" or [status] = "Application Declined" or [status] = "Application Approved Re-Assessment" or [status] = "Certificate Issued" or [sub_status] = "Approved" or [sub_status] = "Declined", type text),
  #"Removed other columns" = Table.SelectColumns(#"Added DECISION", {"code", "DECISION", "TYPE CODE", "updated_datetime_local"}),
  #"Changed column type" = Table.TransformColumnTypes(#"Removed other columns", {{"updated_datetime_local", type date}}),
  #"Filtered DECISION" = Table.SelectRows(#"Changed column type", each ([DECISION] = true)),
  #"Grouped rows" = Table.Group(#"Filtered DECISION", {"TYPE CODE"}, {{"DECISION", each List.Min([updated_datetime_local]), type nullable datetime}}),
  #"Merged queries" = Table.NestedJoin(#"Grouped rows", {"TYPE CODE"}, #"Added DECISION", {"TYPE CODE"}, "Grouped rows", JoinKind.LeftOuter),
  #"Expanded Grouped rows" = Table.ExpandTableColumn(#"Merged queries", "Grouped rows", {"TypeName", "SubClassName", "status", "bpm_previous_status", "code", "COUNT", "OCCURANCE", "updated_datetime_local", "created_by", "generated_description", "sub_status", "sub_status_changed_datetime", "ADVISOR", "ADMIN", "created_datetime_local", "Index", "REFERENCE", "YEAR"}) ,

  #"---READY---" = #"Expanded Grouped rows",

  #"Added READY" =Table.AddColumn(#"---READY---", "READY", each [status] = "Ready for Assessment" or [bpm_previous_status] = "Ready for Assessment", type text),
  #"Removed NON READY" = Table.SelectColumns(#"Added READY", {"code", "READY", "TYPE CODE", "updated_datetime_local"}),
  #"Changed READY" = Table.TransformColumnTypes(#"Removed NON READY", {{"updated_datetime_local", type date}}),
  #"Filtered READY" = Table.SelectRows(#"Changed READY", each ([READY] = true)),
  #"Grouped READY" = Table.Group(#"Filtered READY", {"TYPE CODE"}, {{"READY",  each List.Min([updated_datetime_local]), type nullable datetime}}),
  #"Merged READY" = Table.NestedJoin(#"Grouped READY", {"TYPE CODE"}, #"Added READY", {"TYPE CODE"}, "Grouped READY", JoinKind.LeftOuter),
  #"Expanded Grouped READY" = Table.ExpandTableColumn(#"Merged READY", "Grouped READY", {"DECISION", "TypeName", "SubClassName", "status", "bpm_previous_status", "code", "COUNT", "OCCURANCE", "updated_datetime_local", "created_by", "generated_description", "sub_status", "sub_status_changed_datetime", "ADVISOR", "ADMIN", "created_datetime_local", "Index", "REFERENCE", "YEAR"}),

  #"---MAX---" = #"Expanded Grouped READY",

  #"Grouped MAX" = Table.Group(#"---MAX---", {"TYPE CODE"}, {{"MAX", each List.Max([updated_datetime_local]), type nullable date}, 
{"ALL", each _, type nullable table[#"TYPE CODE" = nullable text, #"READY" = nullable Int64.Type, #"DECISION" = nullable Int64.Type, TypeName = nullable text, SubClassName = nullable text, status = nullable text, bpm_previous_status = nullable text, code = nullable text, COUNT = nullable Int64.Type, OCCURANCE = nullable Int64.Type, updated_datetime_local = nullable Int64.Type, created_by = nullable text, generated_description = nullable text, sub_status = nullable text, sub_status_changed_datetime = nullable Int64.Type, ADVISOR = nullable text, ADMIN = nullable text, created_datetime_local = nullable Int64.Type, Index = nullable Int64.Type, REFERENCE = nullable text, YEAR = nullable number]}}),
  #"Expanded ALL" = Table.ExpandTableColumn(#"Grouped MAX", "ALL", {"READY", "DECISION", "TypeName", "SubClassName", "status", "bpm_previous_status", "code", "COUNT", "OCCURANCE", "updated_datetime_local", "created_by", "generated_description", "sub_status", "sub_status_changed_datetime", "ADVISOR", "ADMIN", "created_datetime_local", "Index", "REFERENCE", "YEAR"}),
  #"Changed column type 1" = Table.TransformColumnTypes(#"Expanded ALL", {{"DECISION", type date}, {"READY", type date}, {"created_datetime_local", type date}, {"MAX", type date}}),
  #"Added TOTAL TIME" = Table.AddColumn(#"Changed column type 1", "TOTAL TIME", each [MAX] - [created_datetime_local]),
  #"Added DECISION TIME" = Table.AddColumn(#"Added TOTAL TIME", "DECISION TIME", each [DECISION]  - [READY], Int64.Type),
  #"Added ADMIN TIME" = Table.AddColumn(#"Added DECISION TIME", "ADMIN TIME", each [READY]  - [created_datetime_local], Int64.Type),
  #"Changed column type 2" = Table.TransformColumnTypes(#"Added ADMIN TIME", {{"ADMIN TIME", Int64.Type}, {"TOTAL TIME", Int64.Type}, {"DECISION TIME", Int64.Type}}),
  #"Reordered columns" = Table.ReorderColumns(#"Changed column type 2", {"TYPE CODE", "TypeName", "SubClassName", "status", "bpm_previous_status", "code", "COUNT", "OCCURANCE", "updated_datetime_local", "created_by", "generated_description", "sub_status", "sub_status_changed_datetime", "ADVISOR", "ADMIN", "Index", "REFERENCE", "YEAR", "TOTAL TIME", "DECISION TIME", "ADMIN TIME", "created_datetime_local", "READY", "DECISION", "MAX"}),

  #"---ADVISOR---" = #"Reordered columns",

  #"Removed non CODE" = Table.SelectColumns(#"---ADVISOR---", {"TYPE CODE", "ADVISOR"}),
  #"Filtered NULL ADVISOR" = Table.SelectRows(#"Removed non CODE", each ([ADVISOR] <> "")),
  #"Filtered BLANK CODE" = Table.SelectRows(#"Filtered NULL ADVISOR", each [TYPE CODE] <> null and [TYPE CODE] <> ""),
  #"Removed duplicates" = Table.Distinct(#"Filtered BLANK CODE", {"TYPE CODE", "ADVISOR"}),
  #"Grouped TYPE" = Table.Group(#"Removed duplicates", {"TYPE CODE"}, {{"Advisor Count", each Table.RowCount(_), Int64.Type}, {"ALL", each _, type nullable table[#"TYPE CODE" = nullable text, ADVISOR = nullable text]}}),
  #"Added OWNER OCCURANCE" = Table.AddColumn(#"Grouped TYPE", "OCC", each Table.AddIndexColumn([ALL], "OWNER", 1, 1)),
  #"Expanded OCC" = Table.ExpandTableColumn(#"Added OWNER OCCURANCE", "OCC", {"ADVISOR", "OWNER"}, {"ADVISOR", "OWNER"}),
  #"Removed ALL" = Table.RemoveColumns(#"Expanded OCC", {"ALL"}),
  #"Pivoted ADVISOR" = Table.Pivot(Table.TransformColumnTypes(#"Removed ALL", {{"OWNER", type text}}), List.Distinct(Table.TransformColumnTypes(#"Removed ALL", {{"OWNER", type text}})[OWNER]), "OWNER", "ADVISOR"),
  #"Merged ADVISOR" = Table.NestedJoin(#"Pivoted ADVISOR", {"TYPE CODE"}, #"Reordered columns", {"TYPE CODE"}, "Pivoted column", JoinKind.LeftOuter),
  #"Changed ADVISOR TYPE" = Table.TransformColumnTypes(#"Merged ADVISOR", {{"1", type text}, {"2", type text}, {"3", type text}, {"4", type text}, {"5", type text}}),
  #"Reordered GROUP" = Table.ReorderColumns(#"Changed ADVISOR TYPE", {"TYPE CODE", "Pivoted column", "Advisor Count", "1", "2", "3", "4", "5"}),
  #"Expanded Pivoted column" = Table.ExpandTableColumn(#"Reordered GROUP", "Pivoted column", {"TypeName", "SubClassName", "status", "bpm_previous_status", "code", "COUNT", "OCCURANCE", "updated_datetime_local", "created_by", "generated_description", "sub_status", "sub_status_changed_datetime", "ADVISOR", "ADMIN", "Index", "REFERENCE", "YEAR", "TOTAL TIME", "DECISION TIME", "ADMIN TIME", "created_datetime_local", "READY", "DECISION", "MAX"}, {"TypeName", "SubClassName", "status", "bpm_previous_status", "code", "COUNT", "OCCURANCE", "updated_datetime_local", "created_by", "generated_description", "sub_status", "sub_status_changed_datetime", "ADVISOR", "ADMIN", "Index", "REFERENCE", "YEAR", "TOTAL TIME", "DECISION TIME", "ADMIN TIME", "created_datetime_local", "READY", "DECISION", "MAX"})

in #"Expanded Pivoted column"

```

## DATA FLOW

##### PULL DATA FROM DATA FLOW AND MERGE WITHOUT PREMIUM

```VBA 
let 

Seafarer_Application = 
   
PowerBI.Dataflows(null) [Data] {0} 
{[dataflowId="a8224ddc-9ecc-40c2-96d2-dd42060baf08"]}[Data] 
{[entity="ReportView Application_Seafarer Application DEDUPE"]}[Data],

Seafarer_Renewal = 

PowerBI.Dataflows(null) [Data] {0} 
{[dataflowId="a8224ddc-9ecc-40c2-96d2-dd42060baf08"]}[Data] 
{[entity="ReportView Application_Seafarer Certificate Renewal Application DEDUPE"]}[Data],

Endorsement_Application = 

PowerBI.Dataflows(null) [Data] {0} 
{[dataflowId="a8224ddc-9ecc-40c2-96d2-dd42060baf08"]}[Data] 
{[entity="ReportView Application_Endorsement Application DEUPE"]}[Data],

Personnel_Certificate = 

PowerBI.Dataflows(null) [Data] {0} 
{[dataflowId="a8224ddc-9ecc-40c2-96d2-dd42060baf08"]}[Data] 
{[entity="ReportView Application_Personnel Certificate Application DUPE"]}[Data],

source = Table.Combine({Seafarer_Application, Seafarer_Renewal, Endorsement_Application, Personnel_Certificate})

in source
```

### FUZZY IND

##### STEP 1 | RAW   | Get data function and data preprocessing. Used for Quality reports.
##### STEP 2 | CLEAN | Process phone numbers and name data for Fuzzy. Used for Fuzzy, seperated for speed.
##### STEP 3 | FUZZY | Fuzzy merge clean data into useable report.

### STEP 1
```VBA
let
  GetData = (Page as number) as list =>
    let MidURL = "?$skiptoken=%3Ccookie%20pagenumber=%22",
    PageNUM = Text.From(Page),
    EndURL =    "%22%20pagingcookie=%22%253ccookie%2520page%253d%25221%2522%253e%253ccontactid%2520last%253d%2522%257bD0E5305F-0085-E211-9FD1-000C29854771%257d%2522%2520first%253d%2522%257b609AF16C-120C-4E2C-9498-00015D9B0068%257d%2522%2520%252f%253e%253c%252fcookie%253e%22%20istracking=%22False%22%20/%3E&$select=createdon,modifiedon,mnz_alert,statuscode,mnz_customernumber,mnz_tritonpersonid,mnz_tradingas,birthdate,fullname,firstname,middlename,lastname,emailaddress1,address1_line1,address1_line2,address1_line3,mnz_address1city,mnz_address1country,address1_postalcode,mobilephone,mnz_otherphone,telephone1,telephone2,mnz_custmnzsolomonid,mnz_custopfsolomonid,mnz_custsarsolomonid,_createdby_value",
    source =Record.ToTable(Json.Document(Web.Contents("https://crm/maritimenewzealand/api/data/v8.0/contacts",
    [ RelativePath=MidURL & PageNUM & EndURL,
    Headers=[Prefer="odata.include-annotations=""OData.Community.Display.V1.FormattedValue"""]]
    ))) {1} [Value]
in  source,
  ListGenerate = List.Generate( () =>
    [Result= try GetData(1) otherwise null, Page = 1],
    each not List.IsEmpty([Result]),
    each [Result = try GetData([Page]+1) otherwise null,
    Page =[Page]+1], each [Result]),
  ConvertedTable = Table.FromList(ListGenerate, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
  ExpandedToList = Table.ExpandListColumn(ConvertedTable, "Column1"),
  #"Expanded Column" = Table.ExpandRecordColumn(ExpandedToList, "Column1", {"@odata.etag", "createdon", "modifiedon", "mnz_alert@OData.Community.Display.V1.FormattedValue", "statuscode@OData.Community.Display.V1.FormattedValue", "mnz_customernumber", "mnz_tritonpersonid", "mnz_tradingas", "birthdate", "fullname", "firstname", "middlename", "lastname", "emailaddress1", "address1_line1", "address1_line2", "address1_line3", "mnz_address1city@OData.Community.Display.V1.FormattedValue", "mnz_address1country@OData.Community.Display.V1.FormattedValue", "address1_postalcode", "mobilephone", "mnz_otherphone", "telephone1", "telephone2", "mnz_custmnzsolomonid", "mnz_custopfsolomonid", "mnz_custsarsolomonid", "_createdby_value@OData.Community.Display.V1.FormattedValue", "contactid"}, {"@odata.etag", "createdon", "modifiedon", "mnz_alert", "statuscode", "mnz_customernumber", "mnz_tritonpersonid", "mnz_tradingas", "birthdate", "fullname", "firstname", "middlename", "lastname", "emailaddress1", "address1_line1", "address1_line2", "address1_line3", "mnz_address1city", "mnz_address1country", "address1_postalcode", "mobilephone", "mnz_otherphone", "telephone1", "telephone2", "mnz_custmnzsolomonid", "mnz_custopfsolomonid", "mnz_custsarsolomonid", "createdby", "contactid"}),
  #"Parsed Date" = Table.TransformColumns(#"Expanded Column", {{"createdon", each Date.From(DateTimeZone.From(_)), type date}, {"modifiedon", each Date.From(DateTimeZone.From(_)), type date}}),
  #"Renamed Columns" = Table.RenameColumns(#"Parsed Date", {{"mnz_tritonpersonid", "TRITON"}, {"fullname", "FULL NAME"}, {"createdon", "Created On"}, {"modifiedon", "Modified On"}, {"mnz_customernumber", "Customer Number"}, {"birthdate", "Date of Birth"}, {"mnz_alert", "Alert"}, {"mnz_tradingas", "Trading As"}, {"firstname", "First Name"}, {"middlename", "Middle Name"}, {"lastname", "Last Name"}, {"emailaddress1", "Email"}, {"address1_line1", "Address 1"}, {"address1_line2", "Address 2"}, {"address1_line3", "Suburb"}, {"mnz_address1city", "City"}, {"address1_postalcode", "Post Code"}, {"mobilephone", "Mobile Phone"}, {"telephone1", "Business Phone"}, {"telephone2", "Home Phone"}, {"mnz_custmnzsolomonid", "Customer MNZ Solomon ID"}, {"mnz_custopfsolomonid", "Customer OPF Solomon ID"}, {"mnz_custsarsolomonid", "Customer SAR Solomon ID"}, {"createdby", "Created By"}, {"statuscode", "Status"}, {"mnz_address1country", "Country"}, {"mnz_otherphone", "Other Phone"}}),
  #"Replaced NULL" = Table.ReplaceValue(#"Renamed Columns", null, "", Replacer.ReplaceValue, {"@odata.etag", "Created On", "Modified On", "Alert", "Status", "Customer Number", "TRITON", "Trading As", "Date of Birth", "FULL NAME", "First Name", "Middle Name", "Last Name", "Email", "Address 1", "Address 2", "Suburb", "City", "Country", "Post Code", "Mobile Phone", "Other Phone", "Business Phone", "Home Phone", "Customer MNZ Solomon ID", "Customer OPF Solomon ID", "Customer SAR Solomon ID", "Created By", "contactid"}),
  #"Added FILTER" = Table.AddColumn(#"Replaced NULL", "FILTER", each Date.From(Date.StartOfWeek(Date.AddDays(DateTime.LocalNow(),-7) ) )
    = Date.StartOfWeek([Created On])),
  #"Added WEEK MODIFIED" = Table.AddColumn(#"Added FILTER", "WEEK MODIFIED", each Text.From(Date.Year(Date.From([Modified On])))
    &"/"& Text.PadStart(Text.From(Date.WeekOfYear(Date.From([Modified On]))),2,"0")),
  #"Added WEEK CREATED" = Table.AddColumn(#"Added WEEK MODIFIED", "WEEK CREATED", each Text.From(Date.Year(Date.From([Created On])))&"/"& Text.PadStart(Text.From(Date.WeekOfYear(Date.From([Created On]))),2,"0")),
  #"Added ADDRESS" = Table.AddColumn(#"Added WEEK CREATED", "ADDRESS", each Text.Upper(Text.Combine(List.Select(Text.Split(
    [Address 1]&" "&[Address 2]&" "&[Suburb]&" "&[City]&" "&[Post Code]&" "&[Country],
    " "),each _ <> "")," "))),
  #"Added DOB" = Table.AddColumn(#"Added ADDRESS", "DOB", each if [Date of Birth] = "" then "00/00/0000" else 
Date.ToText(Date.From(Text.From([Date of Birth])), "dd/MM/yyyy")),
  #"Added FULL NAME REF" = Table.AddColumn(#"Added DOB", "FULL NAME REF", each [DOB]&" "& Text.Upper([FULL NAME])),
  #"Changed Type" = Table.TransformColumnTypes(#"Added FULL NAME REF", {{"Status", type text}, {"@odata.etag", type text}, {"Created On", type text}, {"Modified On", type text}, {"Alert", type text}, {"Customer Number", type text}, {"TRITON", type text}, {"Trading As", type text}, {"Date of Birth", type text}, {"FULL NAME", type text}, {"First Name", type text}, {"Middle Name", type text}, {"Last Name", type text}, {"Email", type text}, {"Address 1", type text}, {"Address 2", type text}, {"Suburb", type text}, {"City", type text}, {"Country", type text}, {"Post Code", type text}, {"Mobile Phone", type text}, {"Other Phone", type text}, {"Business Phone", type text}, {"Home Phone", type text}, {"Customer MNZ Solomon ID", type text}, {"Customer OPF Solomon ID", type text}, {"Customer SAR Solomon ID", type text}, {"Created By", type text}, {"contactid", type text}, {"FILTER", type text}, {"WEEK MODIFIED", type text}, {"WEEK CREATED", type text}, {"ADDRESS", type text}, {"DOB", type text}, {"FULL NAME REF", type text}}),
  #"Added index" = Table.AddIndexColumn(#"Changed Type", "Index", 1, 1, Int64.Type),
  #"Sorted rows" = Table.Sort(#"Added index", {{"Index", Order.Descending}})
in
  #"Sorted rows"
```

### STEP 2

```VBA
let
  Source = PowerBI.Dataflows(null) [Data] {0},
  #"89198763-4791-4d4f-a39e-496037f58a15" = Source{[dataflowId = "89198763-4791-4d4f-a39e-496037f58a15"]}[Data],
  #"Removed details" = Table.RemoveColumns(#"89198763-4791-4d4f-a39e-496037f58a15", {"IsLeaf"}),
  #"CRM RAW" = #"Removed details"{[entity = "CRM RAW"]}[Data],
  #"Replaced PHONE ""-""" = Table.ReplaceValue(#"CRM RAW", "-", "", Replacer.ReplaceText, {"Business Phone", "Home Phone", "Mobile Phone", "Other Phone"}),
  #"Replaced PHONE "" """ = Table.ReplaceValue(#"Replaced PHONE ""-""", " ", "", Replacer.ReplaceText, {"Mobile Phone", "Other Phone", "Business Phone", "Home Phone"}),
  #"Replaced PHONE +64" = Table.ReplaceValue(#"Replaced PHONE "" """, "+64", "0", Replacer.ReplaceText, {"Business Phone", "Home Phone", "Mobile Phone", "Other Phone"}),
  #"Replaced PHONE 064" = Table.ReplaceValue(#"Replaced PHONE +64", "064", each if 
Text.Start([Mobile Phone],3) = "064" or
Text.Start([Home Phone],3) = "064" or
Text.Start([Business Phone],3) = "064" or
Text.Start([Other Phone],3) = "064" then "0" else null, Replacer.ReplaceText,{"Business Phone", "Home Phone", "Mobile Phone", "Other Phone"}),
  #"Replaced PHONE 64" = Table.ReplaceValue(#"Replaced PHONE 064", "64", each if 
Text.Start([Mobile Phone],2) = "64" or
Text.Start([Home Phone],2) = "64" or
Text.Start([Business Phone],2) = "64" or
Text.Start([Other Phone],2) = "64" then "0" else null, Replacer.ReplaceText,{"Business Phone", "Home Phone", "Mobile Phone", "Other Phone"}),
  #"Replaced PHONE 00" = Table.ReplaceValue(#"Replaced PHONE 64", "00", each if 
Text.Start([Mobile Phone],2) = "00" or
Text.Start([Home Phone],2) = "00" or
Text.Start([Business Phone],2) = "00" or
Text.Start([Other Phone],2) = "00" then "0" else null, Replacer.ReplaceText,{"Business Phone", "Home Phone", "Mobile Phone", "Other Phone"}),
  #"Replaced CHARACTERS" = List.Accumulate(Text.ToList("(){}[];',./\<>?:_+-=*&^%$#@!"),#"Replaced PHONE 00",(table,old)=> Table.ReplaceValue(table,old,"", Replacer.ReplaceText,{"Business Phone", "Home Phone", "Mobile Phone", "Other Phone"})),
  #"Replaced PHONE 00 INT" = Table.ReplaceValue(#"Replaced CHARACTERS", "00", each if 
Text.Start([Mobile Phone],2) = "00" or
Text.Start([Home Phone],2) = "00" or
Text.Start([Business Phone],2) = "00" or
Text.Start([Other Phone],2) = "00" then "+" else null, Replacer.ReplaceText,{"Business Phone", "Home Phone", "Mobile Phone", "Other Phone"}),
  #"Added PHONE ERROR" = Table.AddColumn(#"Replaced PHONE 00 INT", "PHONE ERROR", each Text.Select([Mobile Phone],   {"A".."z","!".."*",","..",",":".."@","[".."`","{".."~"}) &
        Text.Select([Business Phone], {"A".."z","!".."*",","..",",":".."@","[".."`","{".."~"}) &
        Text.Select([Other Phone],    {"A".."z","!".."*",","..",",":".."@","[".."`","{".."~"}) &
        Text.Select([Home Phone],     {"A".."z","!".."*",","..",",":".."@","[".."`","{".."~"})),
  #"Replaced '-'" = Table.ReplaceValue(#"Added PHONE ERROR", "'-'", "", Replacer.ReplaceText, {"Last Name"}),
  #"Replaced -" = Table.ReplaceValue(#"Replaced '-'", "-", "", Replacer.ReplaceText, {"Last Name"}),
  #"Replaced ," = Table.ReplaceValue(#"Replaced -", ",", "", Replacer.ReplaceText, {"Last Name"}),
  #"Replaced ." = Table.ReplaceValue(#"Replaced ,", ".", "", Replacer.ReplaceText, {"Last Name"}),
  #"Replaced '_'" = Table.ReplaceValue(#"Replaced .", "'_'", "", Replacer.ReplaceText, {"Last Name"}),
  #"Replaced FIRST -" = Table.ReplaceValue(#"Replaced '_'", "-", "", Replacer.ReplaceText, {"First Name"}),
  #"Replaced FIRST '" = Table.ReplaceValue(#"Replaced FIRST -", "'", "", Replacer.ReplaceText, {"First Name"}),
  #"Replaced FIRST ." = Table.ReplaceValue(#"Replaced FIRST '", ".", "", Replacer.ReplaceText, {"First Name"}),
  #"Replaced FIRST """ = Table.ReplaceValue(#"Replaced FIRST .", """", "", Replacer.ReplaceText, {"First Name"}),
  #"Replaced FIRST ," = Table.ReplaceValue(#"Replaced FIRST """, ",", "", Replacer.ReplaceText, {"First Name"}),
  #"Replaced MIDDLE null" = Table.ReplaceValue(#"Replaced FIRST ,", null, "", Replacer.ReplaceValue, {"Middle Name"}),
  #"Added COUNT FULL" = Table.AddColumn(#"Replaced MIDDLE null", "COUNT", each List.Count(Text.Split([FULL NAME]," ")), Int64.Type),
  #"Added COUNT FIRST" = Table.AddColumn(#"Added COUNT FULL", "COUNT FIRST", each if [First Name] <> "" then Text.Length(List.First(Text.Split([First Name]," "))) else 0),
  // = Table.AddColumn(#"Filtered FULL NULL", "COUNT", each List.Count(Text.Split([Last Name]," ")))
  #"Added COUNT LAST" = Table.AddColumn(#"Added COUNT FIRST", "COUNT LAST", each if [Last Name] <> "" then List.Count(Text.Split([Last Name]," ")) 
else 0, Int64.Type),
  #"Added LEN FIRST" = Table.AddColumn(#"Added COUNT LAST", "LEN FIRST", each Text.Length([First Name]), Int64.Type),
  #"Added SHORT" = Table.AddColumn(#"Added LEN FIRST", "SHORT NAME", each if [COUNT FIRST] < 2 then "XXXXXX" &" "& Text.Upper(List.Last(Text.Split([FULL NAME]," "))) 
else if [COUNT LAST] < 1 then Text.Upper(List.First(Text.Split([FULL NAME]," "))) &" "& "XXXXXX" else
Text.Upper(List.First(Text.Split([FULL NAME]," "))) &" "& List.Last(Text.Split([FULL NAME]," "))),
  #"Added SHORT NAME REF" = Table.AddColumn(#"Added SHORT", "SHORT NAME REF", each if [Date of Birth] = null then 
"00/00/0000 "&[SHORT NAME]
else 
Text.PadStart(Text.From([Date of Birth]),10,"0")&" "&[SHORT NAME], type text),
  #"Removed Columns" = Table.RemoveColumns(#"Added SHORT NAME REF", {"@odata.etag", "Created On", "Modified On", "Alert", "Trading As", "Address 1", "Address 2", "Suburb", "City", "Country", "Post Code", "contactid", "COUNT", "COUNT FIRST", "COUNT LAST", "LEN FIRST", "SHORT NAME", "Status"}),
  #"Added FINANCE" = Table.AddColumn(#"Removed Columns", "FINANCE", each if [Customer MNZ Solomon ID] <> null then [Customer MNZ Solomon ID] else if [Customer OPF Solomon ID] <> null then [Customer OPF Solomon ID] else [Customer SAR Solomon ID]),
  #"Added COUNT" = Table.AddColumn(#"Added FINANCE", "FINANCE COUNT", each List.NonNullCount({[Customer MNZ Solomon ID], [Customer OPF Solomon ID], [Customer SAR Solomon ID]})),
  #"Reordered Columns" = Table.ReorderColumns(#"Added COUNT", {"FILTER", "Created By", "WEEK MODIFIED", "WEEK CREATED", "Customer Number", "TRITON", "Date of Birth", "First Name", "Middle Name", "Last Name", "FULL NAME", "SHORT NAME REF", "ADDRESS", "Email", "Mobile Phone", "Home Phone", "Business Phone", "Other Phone", "Customer MNZ Solomon ID", "Customer OPF Solomon ID", "Customer SAR Solomon ID"}),
  #"Sorted Rows" = Table.Sort(#"Reordered Columns", {{"WEEK MODIFIED", Order.Descending}}),
  #"Changed type text" = Table.TransformColumnTypes(#"Sorted Rows", {{"FILTER", type text}, {"Created By", type text}, {"WEEK MODIFIED", type text}, {"WEEK CREATED", type text}, {"Customer Number", type text}, {"TRITON", type text}, {"Date of Birth", type text}, {"First Name", type text}, {"Middle Name", type text}, {"Last Name", type text}, {"FULL NAME", type text}, {"SHORT NAME REF", type text}, {"ADDRESS", type text}, {"Email", type text}, {"Mobile Phone", type text}, {"Home Phone", type text}, {"Business Phone", type text}, {"Other Phone", type text}, {"Customer MNZ Solomon ID", type text}, {"Customer OPF Solomon ID", type text}, {"DOB", type text}, {"FULL NAME REF", type text}, {"Index", type text}, {"PHONE ERROR", type text}, {"Customer SAR Solomon ID", type text}, {"FINANCE", type text}, {"FINANCE COUNT", type text}})
in
  #"Changed type text"
```

### STEP 3

###### A | FUZZY JOIN ON SHORT NAME REF
###### B | GROUP BY SHORT NAME AND ADD INDEX / GROUPING
###### C | ADD OCCURANCE TO FILTER OUT DUPLICATES
###### D | ADD COUNT FIRST OCCURANCES, ENSURES IF ONLY ONE PARTY IN THE GROUP REMAIING AFTER FILTER THE WHOLE GROUP REMAINS

```VBA
let
  Source = PowerBI.Dataflows(null) [Data] {0},
  DataFlow = Source{[dataflowId = "69326da1-6ec0-49da-bccb-54b5bd817e14"]}[Data],
  #"Removed columns" = Table.RemoveColumns(DataFlow, {"Tags", "ItemKind", "ItemName", "IsLeaf", "dataCategory"}),
  Entity = #"Removed columns"{[entity = "CRM CLEAN"]}[Data],
  #"--- FUZZY---" = Entity,
  #"FUZZY JOIN" = Table.FuzzyNestedJoin(#"--- FUZZY---", {"SHORT NAME REF"}, #"--- FUZZY---", {"SHORT NAME REF"}, "FUZZY", JoinKind.LeftOuter, [IgnoreCase = true, IgnoreSpace = true, Threshold = 0.8]),
  #"Expanded FUZZY" = Table.ExpandTableColumn(#"FUZZY JOIN", "FUZZY", {"FILTER", "Created By", "WEEK MODIFIED", "WEEK CREATED", "Customer Number", "TRITON", "Date of Birth", "First Name", "Middle Name", "Last Name", "FULL NAME", "SHORT NAME REF", "ADDRESS", "Email", "Mobile Phone", "Home Phone", "Business Phone", "Other Phone", "PHONE ERROR", "Customer MNZ Solomon ID", "Customer OPF Solomon ID", "Customer SAR Solomon ID", "FINANCE", "FINANCE COUNT"}, {"FUZZY.FILTER", "FUZZY.Created By", "FUZZY.WEEK MODIFIED", "FUZZY.WEEK CREATED", "FUZZY.Customer Number", "FUZZY.TRITON", "FUZZY.Date of Birth", "FUZZY.First Name", "FUZZY.Middle Name", "FUZZY.Last Name", "FUZZY.FULL NAME", "FUZZY.SHORT NAME REF", "FUZZY.ADDRESS", "FUZZY.Email", "FUZZY.Mobile Phone", "FUZZY.Home Phone", "FUZZY.Business Phone", "FUZZY.Other Phone", "FUZZY.PHONE ERROR", "FUZZY.Customer MNZ Solomon ID", "FUZZY.Customer OPF Solomon ID", "FUZZY.Customer SAR Solomon ID", "FUZZY.FINANCE", "FUZZY.FINANCE COUNT"}),
  #"---GROUP SHORT NAME---" = #"Expanded FUZZY",
  #"Grouped SHORT NAME" = Table.Group(#"---GROUP SHORT NAME---", {"SHORT NAME REF"}, {{"GROUP", each _, type table[#"Customer Number" = text, #"SHORT NAME REF" = text, #"FUZZY.Customer Number" = nullable text, #"FUZZY.SHORT NAME REF" = text, FUZZY.FILTER = text, #"FUZZY.Created By" = text, #"FUZZY.WEEK MODIFIED" = text, #"FUZZY.WEEK CREATED" = text, FUZZY.TRITON = text, #"FUZZY.Date of Birth" = text, #"FUZZY.First Name" = text, #"FUZZY.Middle Name" = text, #"FUZZY.Last Name" = text, #"FUZZY.FULL NAME" = text, FUZZY.ADDRESS = text, FUZZY.Email = text, #"FUZZY.Mobile Phone" = text, #"FUZZY.Home Phone" = text, #"FUZZY.Business Phone" = text, #"FUZZY.Other Phone" = text, #"FUZZY.PHONE ERROR" = text, #"FUZZY.Customer MNZ Solomon ID" = text, #"FUZZY.Customer OPF Solomon ID" = text, #"FUZZY.Customer SAR Solomon ID" = text, FUZZY.FINANCE = text, #"FUZZY.FINANCE COUNT" = text]}, {"COUNT", each Table.RowCount(_), Int64.Type}}),
  #"Filtered SINGLE COUNT" = Table.SelectRows(#"Grouped SHORT NAME", each [COUNT] <> 1),
  #"Sorted SHORT NAME" = Table.Sort(#"Filtered SINGLE COUNT", {{"COUNT", Order.Descending}, {"SHORT NAME REF", Order.Ascending}}),
  #"Added Index" = Table.AddIndexColumn(#"Sorted SHORT NAME", "Index", 1, 1, Int64.Type),
  #"Added Custom" = Table.AddColumn(#"Added Index", "GROUPING", each Text.PadStart(Text.From([Index]),5,"0"), type text),
  #"Expanded GROUP" = Table.ExpandTableColumn(#"Added Custom", "GROUP", {"Customer Number", "FUZZY.Customer Number", "FUZZY.SHORT NAME REF", "FUZZY.FILTER", "FUZZY.Created By", "FUZZY.WEEK MODIFIED", "FUZZY.WEEK CREATED", "FUZZY.TRITON", "FUZZY.Date of Birth", "FUZZY.First Name", "FUZZY.Middle Name", "FUZZY.Last Name", "FUZZY.FULL NAME", "FUZZY.ADDRESS", "FUZZY.Email", "FUZZY.Mobile Phone", "FUZZY.Home Phone", "FUZZY.Business Phone", "FUZZY.Other Phone", "FUZZY.PHONE ERROR", "FUZZY.Customer MNZ Solomon ID", "FUZZY.Customer OPF Solomon ID", "FUZZY.Customer SAR Solomon ID", "FUZZY.FINANCE", "FUZZY.FINANCE COUNT"}, {"Customer Number", "FUZZY.Customer Number", "FUZZY.SHORT NAME REF", "FUZZY.FILTER", "FUZZY.Created By", "FUZZY.WEEK MODIFIED", "FUZZY.WEEK CREATED", "FUZZY.TRITON", "FUZZY.Date of Birth", "FUZZY.First Name", "FUZZY.Middle Name", "FUZZY.Last Name", "FUZZY.FULL NAME", "FUZZY.ADDRESS", "FUZZY.Email", "FUZZY.Mobile Phone", "FUZZY.Home Phone", "FUZZY.Business Phone", "FUZZY.Other Phone", "FUZZY.PHONE ERROR", "FUZZY.Customer MNZ Solomon ID", "FUZZY.Customer OPF Solomon ID", "FUZZY.Customer SAR Solomon ID", "FUZZY.FINANCE", "FUZZY.FINANCE COUNT"}),
  #"---CUSTOMER NO OCCURANCE ---" = #"Expanded GROUP",
  #"Grouped Customer Occurance" = Table.Group(#"---CUSTOMER NO OCCURANCE ---", {"FUZZY.Customer Number"}, {{"RELATED", each Table.AddIndexColumn(_, "Occurence", 1,1), type table}}),
  #"Expanded Customer Occurance" = Table.ExpandTableColumn(#"Grouped Customer Occurance", "RELATED", {"GROUPING", "SHORT NAME REF", "Customer Number", "FUZZY.FILTER", "FUZZY.Created By", "FUZZY.WEEK MODIFIED", "FUZZY.WEEK CREATED", "FUZZY.TRITON", "FUZZY.Date of Birth", "FUZZY.First Name", "FUZZY.Middle Name", "FUZZY.Last Name", "FUZZY.FULL NAME", "FUZZY.SHORT NAME REF", "FUZZY.ADDRESS", "FUZZY.Email", "FUZZY.Mobile Phone", "FUZZY.Home Phone", "FUZZY.Business Phone", "FUZZY.Other Phone", "FUZZY.PHONE ERROR", "FUZZY.Customer MNZ Solomon ID", "FUZZY.Customer OPF Solomon ID", "FUZZY.Customer SAR Solomon ID", "FUZZY.FINANCE", "FUZZY.FINANCE COUNT", "Occurence", "COUNT", "Index"}),
  #"Reordered Customer Occurance" = Table.ReorderColumns(#"Expanded Customer Occurance", {"GROUPING", "Customer Number", "SHORT NAME REF", "FUZZY.Customer Number", "FUZZY.SHORT NAME REF", "FUZZY.FILTER", "FUZZY.Created By", "FUZZY.WEEK MODIFIED", "FUZZY.WEEK CREATED", "FUZZY.TRITON", "FUZZY.Date of Birth", "FUZZY.First Name", "FUZZY.Middle Name", "FUZZY.Last Name", "FUZZY.FULL NAME", "FUZZY.ADDRESS", "FUZZY.Email", "FUZZY.Mobile Phone", "FUZZY.Home Phone", "FUZZY.Business Phone", "FUZZY.Other Phone", "FUZZY.PHONE ERROR", "FUZZY.Customer MNZ Solomon ID", "FUZZY.Customer OPF Solomon ID", "FUZZY.Customer SAR Solomon ID", "FUZZY.FINANCE", "FUZZY.FINANCE COUNT", "Occurence", "COUNT"}),
  #"---GROUP SHORT NAME REF---" = #"Reordered Customer Occurance",
  #"Group FILTER SHORT" = Table.Group(#"---GROUP SHORT NAME REF---", {"SHORT NAME REF"}, {{"Count", each _, type table[GROUPING = text, #"Customer Number" = text, #"SHORT NAME REF" = text, #"FUZZY.Customer Number" = nullable text, #"FUZZY.SHORT NAME REF" = text, FUZZY.FILTER = text, #"FUZZY.Created By" = text, #"FUZZY.WEEK MODIFIED" = text, #"FUZZY.WEEK CREATED" = text, FUZZY.TRITON = text, #"FUZZY.Date of Birth" = text, #"FUZZY.First Name" = text, #"FUZZY.Middle Name" = text, #"FUZZY.Last Name" = text, #"FUZZY.FULL NAME" = text, FUZZY.ADDRESS = text, FUZZY.Email = text, #"FUZZY.Mobile Phone" = text, #"FUZZY.Home Phone" = text, #"FUZZY.Business Phone" = text, #"FUZZY.Other Phone" = text, #"FUZZY.PHONE ERROR" = text, #"FUZZY.Customer MNZ Solomon ID" = text, #"FUZZY.Customer OPF Solomon ID" = text, #"FUZZY.Customer SAR Solomon ID" = text, FUZZY.FINANCE = text, #"FUZZY.FINANCE COUNT" = text, Occurence = number, COUNT = number, Index = number]}}),
  #"Added FILTER" = Table.AddColumn(#"Group FILTER SHORT", "FILTER", each Table.SelectRows([Count], each ([Occurence] = 1))),
  #"Added FILTER COUNT" = Table.AddColumn(#"Added FILTER", "COUNT FILTER", each Table.RowCount([FILTER])),
  #"Removed FILTER" = Table.RemoveColumns(#"Added FILTER COUNT", {"SHORT NAME REF", "FILTER"}),
  #"Expanded Count" = Table.ExpandTableColumn(#"Removed FILTER", "Count", {"GROUPING", "Customer Number", "SHORT NAME REF", "FUZZY.Customer Number", "FUZZY.SHORT NAME REF", "FUZZY.FILTER", "FUZZY.Created By", "FUZZY.WEEK MODIFIED", "FUZZY.WEEK CREATED", "FUZZY.TRITON", "FUZZY.Date of Birth", "FUZZY.First Name", "FUZZY.Middle Name", "FUZZY.Last Name", "FUZZY.FULL NAME", "FUZZY.ADDRESS", "FUZZY.Email", "FUZZY.Mobile Phone", "FUZZY.Home Phone", "FUZZY.Business Phone", "FUZZY.Other Phone", "FUZZY.PHONE ERROR", "FUZZY.Customer MNZ Solomon ID", "FUZZY.Customer OPF Solomon ID", "FUZZY.Customer SAR Solomon ID", "FUZZY.FINANCE", "FUZZY.FINANCE COUNT", "Occurence", "COUNT", "Index"}, {"GROUPING", "Customer Number", "SHORT NAME REF", "FUZZY.Customer Number", "FUZZY.SHORT NAME REF", "FUZZY.FILTER", "FUZZY.Created By", "FUZZY.WEEK MODIFIED", "FUZZY.WEEK CREATED", "FUZZY.TRITON", "FUZZY.Date of Birth", "FUZZY.First Name", "FUZZY.Middle Name", "FUZZY.Last Name", "FUZZY.FULL NAME", "FUZZY.ADDRESS", "FUZZY.Email", "FUZZY.Mobile Phone", "FUZZY.Home Phone", "FUZZY.Business Phone", "FUZZY.Other Phone", "FUZZY.PHONE ERROR", "FUZZY.Customer MNZ Solomon ID", "FUZZY.Customer OPF Solomon ID", "FUZZY.Customer SAR Solomon ID", "FUZZY.FINANCE", "FUZZY.FINANCE COUNT", "Occurence", "COUNT.1", "Index"}),
  #"Transform columns" = Table.TransformColumnTypes(#"Expanded Count", {{"COUNT FILTER", type text}}),
  #"Replace errors" = Table.ReplaceErrorValues(#"Transform columns", {{"COUNT FILTER", null}})
in
  #"Replace errors"
```
### WORD ASSOCIATIONS PIVOT

```VBA

let
    Source = Excel.CurrentWorkbook(){[Name="COMPANY!Query_from_Microsoft_CRM"]}[Content],
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Removed Other Columns" = Table.SelectColumns(#"Promoted Headers",{"Organisation Name"}),
    #"Removed Duplicates1" = Table.Distinct(#"Removed Other Columns"),
    #"Uppercased Text" = Table.TransformColumns(#"Removed Duplicates1",{{"Organisation Name", Text.Upper, type text}}),
    #"Duplicated Column" = Table.DuplicateColumn(#"Uppercased Text", "Organisation Name", "Organisation"),
    #"Reordered Columns" = Table.ReorderColumns(#"Duplicated Column",{"Organisation", "Organisation Name"}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Reordered Columns", "Organisation Name", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"Organisation Name.1", "Organisation Name.2", "Organisation Name.3", "Organisation Name.4", "Organisation Name.5", "Organisation Name.6", "Organisation Name.7"}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Split Column by Delimiter", {"Organisation"}, "Attribute", "Value"),
    #"Removed Columns" = Table.RemoveColumns(#"Unpivoted Columns",{"Attribute"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "LEN", each Text.Length([Value])),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each ([LEN] <> 1 and [LEN] <> 2)),
    #"Added Custom1" = Table.AddColumn(#"Filtered Rows", "START", each Text.Start([Value],3)),
    #"Removed Other" = Table.SelectColumns(#"Added Custom1",{"START", "Value"}),
    #"Grouped Rows" = Table.Group(#"Removed Other", {"START"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Count", Order.Descending}}),
    #"Merged Queries" = Table.NestedJoin(#"Sorted Rows", {"START"}, #"Removed Other", {"START"}, "Removed Other", JoinKind.LeftOuter),
    #"Expanded Removed Other" = Table.ExpandTableColumn(#"Merged Queries", "Removed Other", {"Value"}, {"Value"}),
    #"Removed Duplicates" = Table.Distinct(#"Expanded Removed Other", {"Value"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Duplicates", "Index", 1, 1, Int64.Type),
    #"GroupedRows" = Table.Group(#"Added Index", {"START"}, {{"tmp", each Table.AddIndexColumn(_, "Occurence", 1,1), type table}}),
    #"Expanded tmp" = Table.ExpandTableColumn(GroupedRows, "tmp", {"Value", "Count", "Index", "Occurence"}, {"Value", "Count", "Index", "Occurence"}),
    #"Removed Index" = Table.RemoveColumns(#"Expanded tmp",{"Index"}),
    #"Grouped Rows1" = Table.Group(#"Removed Index", {"START"}, {{"WORDS", each Table.RowCount(_), Int64.Type}}),
    Merged = Table.NestedJoin(#"Grouped Rows1", {"START"}, #"Removed Index", {"START"}, "Removed Index", JoinKind.LeftOuter),
    #"Expanded Removed Index" = Table.ExpandTableColumn(Merged, "Removed Index", {"Value", "Count", "Occurence"}, {"Value", "Count", "Occurence"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Expanded Removed Index", {{"Occurence", type text}}, "en-NZ"), List.Distinct(Table.TransformColumnTypes(#"Expanded Removed Index", {{"Occurence", type text}}, "en-NZ")[Occurence]), "Occurence", "Value"),
    #"Sorted Rows1" = Table.Sort(#"Pivoted Column",{{"Count", Order.Descending}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Sorted Rows1",{"START", "Count", "WORDS", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", "81", "82", "83", "84"})
in
    #"Reordered Columns1"
```

### WORD COUNT

```VBA
let
    Source = Excel.CurrentWorkbook(){[Name="NZBN"]}[Content],
    #"Uppercased Text" = Table.TransformColumns(Source,{{"Organisation Name", Text.Upper, type text}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Uppercased Text", "Organisation Name", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"Organisation Name.1", "Organisation Name.2", "Organisation Name.3", "Organisation Name.4", "Organisation Name.5", "Organisation Name.6", "Organisation Name.7"}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Split Column by Delimiter", {"NZBN", "ENTITYSTATUS", "ENTITYNAME"}, "Attribute", "Value"),
    #"Grouped Rows" = Table.Group(#"Unpivoted Columns", {"Value"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "LEN", each Text.Length([Value])),
    #"Sorted Rows" = Table.Sort(#"Added Custom",{{"Count", Order.Descending}}),
    #"Filtered Rows" = Table.SelectRows(#"Sorted Rows", each ([LEN] <> 1)),
    #"Added Custom1" = Table.AddColumn(#"Filtered Rows", "PERC", each [Count] / Table.RowCount(#"Filtered Rows"))
in
    #"Added Custom1"
```
### FUZZY ADDRESS

```VBA
let
    Source = #"COMPANY",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Status] = "Active") and ([ADDRESS] <> "" and [ADDRESS] <> "ANK, CNK, NEW ZEALAND" and [ADDRESS] <> "NEW ZEALAND")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Customer Number", "Organisation Name", "MONTH", "ADDRESS", "FULL ORG NAME"}),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Other Columns",{"Organisation Name"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Removed Columns",{"MONTH", "Customer Number", "ADDRESS", "FULL ORG NAME"}),
    #"Grouped INDEX RELATED ORG" = Table.Group(#"Reordered Columns1", {"ADDRESS"}, {{"RELATED", each Table.AddIndexColumn(_, "Occurence", 1,1), type table}}),
    #"Expanded RELATED" = Table.ExpandTableColumn(#"Grouped INDEX RELATED ORG", "RELATED", {"MONTH", "Customer Number", "FULL ORG NAME", "Occurence"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded RELATED",{"Occurence", "MONTH", "ADDRESS", "Customer Number", "FULL ORG NAME"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"ADDRESS", Order.Ascending}, {"FULL ORG NAME", Order.Ascending}})
in
    #"Sorted Rows"
```

# MOE
###### [LIBRARY](https://github.com/ScottypNZ/CODE-LIBRARY)   |   [INDEX](#INDEX)
-------------------------

### REST API CALL

```VBA
let 
    BaseUrl = https://education-nz.atlassian.net/rest/api/3/search?jql=project in ('has'),


    Credentials = scott.phillips@education.govt.nz:ATATT3xFfGF0arW etc,

    JiraIDPerPage = 100,

    GetJson = (Url) =>
        let 
            RawData = Web.Contents(Url),
            Json    = Json.Document(RawData)
        in  Json,

    GetJiraIDCount = () =>
        let Url   = BaseUrl & "&maxResults=0",
            Json  = GetJson(Url),
            Count = Json[#"total"]
        in  Count,

    GetPage = (Index) =>
        let Skip  = "&startAt=" & Text.From(Index * JiraIDPerPage),
            Top   = "&maxResults=" & Text.From(JiraIDPerPage),
            Url   = BaseUrl & Skip & Top,
            Json  = GetJson(Url),
            Value = Json[#"issues"]
        in  Value,

    JiraIDCount = List.Max({ JiraIDPerPage, GetJiraIDCount() }),
    PageCount   = Number.RoundUp(JiraIDCount / JiraIDPerPage),
    PageIndices = { 0 .. PageCount - 1 },
    Pages       = List.Transform(PageIndices, each GetPage(_)),
    JiraID    = List.Union(Pages),
    Table       = Table.FromList(JiraID, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Renamed = Table.RenameColumns(Table,{{"Column1", "issues"}}),
    Expanded = Table.ExpandRecordColumn(Renamed, "issues", {"id", "key", "fields"}),
```
### SIMPLE REST API CALL

```VBA

let
    apiUrl = https://api.example.com/data,
    accessToken = "your_access_token_here",
    headers = [#"Authorization" = "Bearer " & accessToken],
    apiResponse = Web.Contents(apiUrl, [Headers=headers]),
    apiTable = Json.Document(apiResponse)
in
    apiTable
```

### LIST GENERATE API CALL

```VBA

= (startAt as number) as list  =>
let StartURL = "HTTPS://education-nz.atlassian.net/rest/api/3/search?jql=project=HAS&maxResults=100&startAt=",
source = Json.Document(Web.Contents( StartURL & Number.ToText(startAt) )) [issues] 
in  source

------

let ListGenerate = List.Generate( () =>
       [Result= try GetDataListRecordToTable(1) otherwise null, startAt = 0], 
       each not List.IsEmpty([Result]),
       each [Result = try GetDataListRecordToTable([startAt]+100) otherwise null, 
       startAt =[startAt]+100], each [Result]),

    ToTable = Table.Buffer(Table.FromList(ListGenerate, Splitter.SplitByNothing(), null, null, ExtraValues.Error)),
    ExpandTable = Table.ExpandListColumn(ToTable, "Column1")
in expand table

```



