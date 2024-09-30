=====================

### INDEX

 * [GENERAL](#GENERAL)
 * [MOE](#MOE)
 * [DATE](#DATE)
 
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

IF RULE PLACEMENT DATA
```
Employment & Training Placements = 
IF(or('FY 2017'[Job Progression]="First Job",'FY 2017'[Job Progression]="Career progression with new employer"),1,
IF('FY 2017'[Training Placement Category]="31 Days Training",1,0))
```
AGE CATEGORY
```
Age Category = 
IF(
    'FY 2018'[Calculated Age]<18,"Under 18",
    IF(AND( 'FY 2018'[Calculated Age]<25, 'FY 2018'[Calculated Age]>=18),"18-24",
    IF(AND( 'FY 2018'[Calculated Age]<50, 'FY 2018'[Calculated Age]>=25),"25-49",
    IF(AND( 'FY 2018'[Calculated Age]<65, 'FY 2018'[Calculated Age]>=50),"50-64","65+"))))
```

REGION
```
Provider Region = 
IF(CONTAINSSTRING('FY 2017'[Provider Name],"(AKLD)")=TRUE(),"Auckland, NZ",
IF(CONTAINSSTRING('FY 2017'[Provider Name],"(NTHLD)")=TRUE(),"Northland, NZ",
IF(CONTAINSSTRING('FY 2017'[Provider Name],"(WAI)")=TRUE(),"Waikato, NZ",
    IF(CONTAINSSTRING('FY 2017'[Provider Name],"(BOP)")=TRUE(),"Bay of Plenty, NZ",
        IF(CONTAINSSTRING('FY 2017'[Provider Name],"(HB)")=TRUE(),"Hawke's Bay, NZ",
            IF(CONTAINSSTRING('FY 2017'[Provider Name],"(MW)")=TRUE(),"Manawatu-Whanganui, NZ",
                IF(CONTAINSSTRING('FY 2017'[Provider Name],"(WLG)")=TRUE(),"Wellington, NZ",
                IF(CONTAINSSTRING('FY 2017'[Provider Name],"(NMT)")=TRUE(),"Nelson-Marlborough-Tasman, NZ",
                IF(CONTAINSSTRING('FY 2017'[Provider Name],"(CAN)")=TRUE(),"Canterbury, NZ",
                IF(CONTAINSSTRING('FY 2017'[Provider Name],"(OTAGO)")=TRUE(),"Otago, NZ",
                IF(CONTAINSSTRING('FY 2017'[Provider Name],"(STHLD)")=TRUE(),"Southland, NZ","")))))))))))
```

QUARTER
```
Quarter = 
IF(AND('FY 2017'[Date].[Date]>=DATE(2016,07,1),'FY 2017'[Date].[Date]<DATE(2016,10,1)),"Q1",
    IF(AND('FY 2017'[Date].[Date]>=DATE(2016,10,1),'FY 2017'[Date].[Date]<DATE(2017,1,1)),"Q2",
    IF(AND('FY 2017'[Date].[Date]>=DATE(2017,1,1),'FY 2017'[Date].[Date]<DATE(2017,4,1)),"Q3",
    IF(AND('FY 2017'[Date].[Date]>=DATE(2017,4,1),'FY 2017'[Date].[Date]<DATE(2017,7,1)),"Q4",""))))
```

AGE
```
Calculated Age = 
DATEDIFF('FY 2017'[DOB].[Date],'FY 2017'[Date].[Date],YEAR)
```

LEADS BY CREATION DATE
```
Leads by Creation Date = 
CALCULATE (
    [Leads ALL],
    TREATAS (
        VALUES ( 'Calendar'[Date] ),
        'Activation and Coordination'[Date - Created]
    )
)
```

LEADS ALL
```
Leads ALL = 
VAR X =
COUNTROWS ( 'Activation and Coordination' )

RETURN

IF(X = BLANK(), "", X)
```


LEAD DAYS AVERAGE
```
Lead Days Average = 
VAR X =
    AVERAGEX (
        'Activation and Coordination',
        'Activation and Coordination'[(C) Lead Age DAYS]
    )
RETURN
    IF ( X = BLANK (), "", X )
```

DYNAMIC SINGLE VALUE
```
DynamicSingleValue = 
VAR StartDate = DATE(2018,1,2)
VAR EndDate = MAX(SELECTEDVALUE('Report Tracker'[Report Due]), SELECTEDVALUE('Report Tracker'[Report Submitted Date]))

RETURN 
CALCULATE(
    [(m) Reports Submitted],
    DATESBETWEEN('Calendar'[Dates], DATE(2023,2,1), DATE(2023,4,1))
)
```

MAP LABEL
```
Map Label = 

VAR _count = COUNTX('Applications', 'Applications'[PACO ID])
VAR _Region = SELECTEDVALUE(Region[Region Short])
VAR _amount = FORMAT([(m) $ Total Funds Requested ex GST], "$#,###")

RETURN

COMBINEVALUES(" ",_count, _Region, " Applicants rqstg ", _amount)
```

CROSS JOIN
```
(m) CPF Education Apps = 
CALCULATE (
    [(m) Applications Received],
    FILTER (
        CROSSJOIN (
            ALL ( 'Documents'[$ - CPF Education Amount] ),
            ALL ( Initiatives[Initiative Type] )
        ),
        'Documents'[$ - CPF Education Amount] > 0
            || Initiatives[Initiative Type] = "CPF Education"
    )
)
```

DECLINE REASONS
```
(m) Declined Reasons = 
IF (
    HASONEVALUE ( 'Payment Tracker'[PACOF ID] ),
    CONCATENATEX (
        'Declined Reasons',
        'Declined Reasons'[Declined / Withdrawn Types],
        UNICHAR ( 10 )
    ),
    BLANK ()
)
```


DECISION DAYS
```
(m) Average Age DAYS Decision = 
AVERAGEX (
    'Applications',
    'Applications'[(c) Age of Application DAYS]
)
```

TURNAROUND TIMES
```
(m) Assessment Turnaround = 
VAR Single =
    IF (
            VALUES ( 'Documents'[Assessor_ID] ) <> BLANK () &&
            Values ( 'Documents'[Modified] ) = BLANK (),
        DATEDIFF (
            FIRSTDATE ( 'Last Refresh'[Last Refresh Date] ),
            Values ( 'Documents'[Modified] ),
            DAY
        ),
        IF (
            Values ( 'Applications'[Submitted Date] )
                > Values ( 'Documents'[Modified] ),
            DATEDIFF (
                Values ( 'Applications'[Submitted Date] ),
                Values ( 'Documents'[Modified] ),
                DAY
            ),
            DATEDIFF (
                Values ( 'Documents'[Assessor.Modified] ),
                Values ( 'Documents'[Modified] ),
                DAY
            )
        )
    )

VAR __table = SUMMARIZE('Documents',[PACOF ID],"__value", Single)
RETURN
IF(HASONEVALUE('Documents'[PACOF ID]),Single,AVERAGEX(__table,[__value]))
```


SUBTRACTING DATES
```
(m) Applications Received - Submitted Date = 
CALCULATE (
    [(m) Applications Received],
    TREATAS (
        VALUES ( 'Calendar'[Dates] ),
        'Applications'[Submitted Date]
    )
)
```

CONCAT NAME
```
1st Contact - Name = 
CONCATENATE('Main Tracking'[First Name]&" ", 'Main Tracking'[Last Name])
```

FUNDING GROUP
```
Funding Group = 
SWITCH(
    TRUE(),
    'Main Tracking'[Proposed Amount (excl. GST)] = BLANK(), BLANK(),
    'Main Tracking'[Proposed Amount (excl. GST)] < 5000, "< $5000",
    'Main Tracking'[Proposed Amount (excl. GST)] <= 10000, "$5000 - $10,000",
    'Main Tracking'[Proposed Amount (excl. GST)] <= 20000, "$10,001 - $20,000",
    "$20,001 - $30,000"
)
```

FUNDS OUTSTANING
```
(m) - Funds Outstanding % = 
DIVIDE( [(m) $ - Funds Outstanding], [(m) $ - Funding Value], 0 )
```

FUND VALUE LESS GST
```
(m) $ - Funding Value = 
CALCULATE (
    SUM ( 'Main Tracking'[$Approved (minus GST)] ),
    FILTER ( 'Main Tracking', ('Main Tracking'[(m) Approved Applications]) > 0 )
)
```

RUNNING TOTAL RECIEVED AFTER
```
(m) Accountability Reports Received running total = 
CALCULATE(
	[(m) Accountability Reports Received],
	FILTER(
		ALLSELECTED('Main Tracking'[Date AR received]),
		ISONORAFTER('Main Tracking'[Date AR received], MAX('Main Tracking'[Date AR received]), DESC)
	)
```

DYNAMIC Y AXIS
```
Dynamic Y-Axis = {
    ("Māori", NAMEOF('Sheet1'[ShowMaori-Graph2]), 3),
    ("Non-Pacific", NAMEOF('Sheet1'[ShowNonPacific-Graph2]), 1),
    ("NZ European", NAMEOF('Sheet1'[ShowNZEuropean-Graph2]), 4),
    ("Pacific", NAMEOF('Sheet1'[ShowPacificRate-Graph2]), 0),
    ("Total NZ population", NAMEOF('Sheet1'[ShowTotal-Graph2]), 2)
}
```

NAVIGATION PANE
```
NavigationPage4 = SELECTEDVALUE('Goal 4 Navigation'[Outcome])
```


HIGHLIGHT 
```
CF - Highlight row = 
VAR go = MAX(Sheet1[Goal])
VAR CF =
IF(go = 1, "#ff3300", 
IF(go = 2, "#990099",
IF(go = 3, "#E6AF24",
IF(go = 4, "#009999","grey"))))
RETURN
CF
```

GRAPH TYPE
```
GraphType_measure = IF(NOT(ISBLANK(SUM(Sheet1[Pacifica Rate]))) && ISBLANK(COUNT(Sheet1[Sub indicator])), 1, 
IF(ISBLANK(SUM(Sheet1[Pacifica Rate])) && ISBLANK(COUNT(Sheet1[Sub indicator])) && NOT(ISBLANK(SUM(Sheet1[Pacifica Number]))), 2, 
IF(NOT(ISBLANK(COUNT(Sheet1[Sub indicator]))) && ISBLANK(SUM(Sheet1[Pacifica Rate])) && NOT(ISBLANK(SUM(Sheet1[Pacifica Number]))), 3,
IF(NOT(ISBLANK(COUNT(Sheet1[Sub indicator]))) && NOT(ISBLANK(SUM(Sheet1[Pacifica Rate]))), 4,
    BLANK()))))
```

HIDE GRAPH
```
HideGraph = IF(ISFILTERED(Sheet1[Indicator name]),CALCULATE(SUM(Sheet1[Pacifica Rate])),CALCULATE(SUM(Sheet1[Pacifica Rate]), Sheet1[IsLatestDateByOutcome] = TRUE))
```


NZ EUROPEAN GRAPH
```
NZEuropean-Graph = IF(NOT(ISBLANK(Sheet1[Pacifica Rate])) && ISBLANK(Sheet1[Sub indicator]), Sheet1[NZ European Rate], 
IF(ISBLANK(Sheet1[Pacifica Rate]) && ISBLANK(Sheet1[Sub indicator]) && NOT(ISBLANK(Sheet1[Pacifica Number])), Sheet1[NZ European Number], 
IF(NOT(ISBLANK(Sheet1[Sub indicator])) && ISBLANK(Sheet1[Pacifica Rate]) && NOT(ISBLANK(Sheet1[Pacifica Number])), Sheet1[NZ European Number],
IF(NOT(ISBLANK(Sheet1[Sub indicator])) && NOT(ISBLANK(Sheet1[Pacifica Rate])), Sheet1[NZ European Rate],
    BLANK()))))
```

SCROLL MESSAGE
```Scroll-message = 
VAR min_val = MIN(Sheet1[index])--This code will show the narrative corresponding to the Indicator name even when nothing is selected
RETURN
    CALCULATE(MIN(Sheet1[Scroll-message_calc]), Sheet1[index] = min_val)
```

TOTALS
```
Total-Graph = IF(NOT(ISBLANK(Sheet1[Pacifica Rate])) && ISBLANK(Sheet1[Sub indicator]), Sheet1[Total Rate], 
IF(ISBLANK(Sheet1[Pacifica Rate]) && ISBLANK(Sheet1[Sub indicator]) && NOT(ISBLANK(Sheet1[Pacifica Number])), Sheet1[Total Number], 
IF(NOT(ISBLANK(Sheet1[Sub indicator])) && ISBLANK(Sheet1[Pacifica Rate]) && NOT(ISBLANK(Sheet1[Pacifica Number])), Sheet1[Total Number],
IF(NOT(ISBLANK(Sheet1[Sub indicator])) && NOT(ISBLANK(Sheet1[Pacifica Rate])), Sheet1[Total Rate],
    BLANK()))))
```

ALL CITY SUMMARY
```
All City Summary = 
    CALCULATE(
        COUNTA(ToloaGrantRequests[ID]),
        ALLEXCEPT(ToloaGrantRequests,ToloaGrantRequests[City]))&" "&
ToloaGrantRequests[City]&" recipients receiving "&
    FORMAT(
    CALCULATE(
        SUM(ToloaGrantRequests[Total Paid]),
        ALLEXCEPT(ToloaGrantRequests,ToloaGrantRequests[City])),
        "$#,###")
```

PROVIDER SUMMARY
```
Steam Providers Summary = 
    IF(ToloaGrantRequests[Grant Type Groups]="Steam Provider",
        CALCULATE(
        COUNTA(ToloaGrantRequests[ID]),
        ALLEXCEPT(ToloaGrantRequests,ToloaGrantRequests[City]))&" "&
ToloaGrantRequests[City]&" recipients receiving "&
    FORMAT(
    CALCULATE(
        SUM(ToloaGrantRequests[Total Paid]),
        ALLEXCEPT(ToloaGrantRequests,ToloaGrantRequests[City])),
        "$#,###"))
```

PERCENTAGE OF TOTAL
```
INELIGIBLE APPS = Format( SUM(CRM_Applicatons[Not Eligible] )  /  max(CRM_Applicatons[CUM ALL]) , "Percent" )
```

LABEL
```
Map Label = 
   CALCULATE(
        COUNTA('3  2022/23 Toloa Scholarship Applications Database'[TOL ID]),'3  2022/23 Toloa Scholarship Applications Database'[App Status]="Complete"&&'3  2022/23 Toloa Scholarship Applications Database'[ApplicationOutcome]<>"Declined",
        ALLEXCEPT('3  2022/23 Toloa Scholarship Applications Database','3  2022/23 Toloa Scholarship Applications Database'[Residence NZ Region]))&" "&
'3  2022/23 Toloa Scholarship Applications Database'[Residence Region]&" successful applicants"
```

INVOICE PAID
```
Paid = 
IF(OR('Invoice Tracker July 2021-2022'[19. Paid]=BLANK(),'Invoice Tracker July 2021-2022'[19. Paid]="No"),"Unpaid","Paid")
``` 

INVOICE STATUS
```
Invoice Status = 
IF('Invoice Tracker July 2021-2022'[Paid]="Paid","Paid",
IF('Invoice Tracker July 2021-2022'[3. Fin Yr]="2020/21"&&'Invoice Tracker July 2021-2022'[13. 2020/21 Recon]=BLANK(),"Awaiting Provider Recon",
IF('Invoice Tracker July 2021-2022'[3. Fin Yr]="2020/21"&&'Invoice Tracker July 2021-2022'[13. 2020/21 Recon]="No","Awaiting Provider Recon",
IF('Invoice Tracker July 2021-2022'[Paid]="Unpaid"&&'Invoice Tracker July 2021-2022'[To be approved]=0,"Awaiting Verification",
if('Invoice Tracker July 2021-2022'[Paid]="Unpaid"&&'Invoice Tracker July 2021-2022'[To be approved]=1&&'Invoice Tracker July 2021-2022'[16. Mgmt Approval]=BLANK(),"To send for Mgmt Approval",
IF('Invoice Tracker July 2021-2022'[Paid]="Unpaid"&&'Invoice Tracker July 2021-2022'[To be approved]=1&&'Invoice Tracker July 2021-2022'[16. Mgmt Approval]<>BLANK()&&'Invoice Tracker July 2021-2022'[18. Finance]=BLANK(),"To be sent to Finance",
IF('Invoice Tracker July 2021-2022'[Paid]="Unpaid"&&'Invoice Tracker July 2021-2022'[16. Mgmt Approval]<>BLANK()&&'Invoice Tracker July 2021-2022'[To be approved]=1&&'Invoice Tracker July 2021-2022'[18. Finance]<>BLANK(),"Awaiting Pmt","")))))))```
```



APPROVAL TIME
```
Invoice approval time = 
DATEDIFF(
    IF('Invoice Tracker July 2021-2022'[16. Mgmt Approval]<>BLANK()&&'Invoice Tracker July 2021-2022'[Created]>'Invoice Tracker July 2021-2022'[16. Mgmt Approval],'Invoice Tracker July 2021-2022'[16. Mgmt Approval],'Invoice Tracker July 2021-2022'[Created]),
IF('Invoice Tracker July 2021-2022'[16. Mgmt Approval]=BLANK(),TODAY(),'Invoice Tracker July 2021-2022'[16. Mgmt Approval]),DAY)
```

12 MONTH SPEND
```
(C) 12 mths emp$ = 
VAR FY = 'Verified MARs ALL'[MAR FY]
VAR Activity = "12 mths emp"
VAR _Rate =
    CALCULATE (
        SELECTEDVALUE ( Rates[Rate]),
        Rates[Activity] = Activity,
        Rates[FY] = FY
    )
RETURN
    IF (
        'Verified MARs ALL'[Continuous Employment > 12 mths] <> BLANK(),
        'Verified MARs ALL'[Continuous Employment > 12 mths] * _Rate,
        BLANK ()
    )
```

RUNNING TOTAL NB: Sub is a binary text criteria in adjecent cell
```DAX
= Table.AddColumn(#"Added SUB", "CUM", each 
List.Sum( List.FirstN( #"Added SUB"[SUB], [Index] ) ) , Int64.Type)
```

MAXIF
```DAX
LastCertOfGivenType = 
VAR maxDate = CALCULATE(
              MAX('Navigator Certificate'[Expiry Due Date]),
                FILTER('Navigator Certificate', Navigator Certificate'[VesselCertType] 
		    = EARLIER('Navigator Certificate'[VesselCertType]) ) )
RETURN IF('Navigator Certificate'[Expiry Due Date] = maxDate,1,0)
```
FILTERS A FIELD WITHIN A TABLE WHERE CONDITION IN ISSUE COLUMN IS 'DUPE'
```DAX
=COUNTX(FILTER(PEOPLE_CLEAN,PEOPLE_CLEAN[ISSUE]="DUPE"),PEOPLE_CLEAN[ISSUE])
```
FILTERS AND COUNTS THE COLUMN WHERE VALUE IS NOT BLANK (MATCHED COLUMN)
```
=COUNTX(FILTER(PEOPLE_CLEAN,PEOPLE_CLEAN[TRITON]<>""),PEOPLE_CLEAN[TRITON])
```
SUBTRACTS COUNT OF ANY FIELD (ROW COUNT) FROM THE ABOVE. THIS RETURNS A COUNT OF THE BLANK FIELDS
```
=COUNT([Customer Number])-[MATCHED]
```

#### CREATE SUBTOTAL
```DAX
TABLE1 = DATATABLE ( "Name", STRING, "Ordinal", STRING, { { "", "" },  { "", "" }, { "", "" } } )
```

```DAX
TABLE1 = DATATABLE (
    "Name", STRING,
    "Ordinal", STRING,
    {
        { "", "" },
        { "", "" },
        { "", "" }
    }
)
```

#### ALMOST SUBTOTAL   
```DAX
 =IF ( not ( isfiltered ( BHR[ROW] ) ) ,  [COUNT ROOM REF] / 30.4,  [COUNT ROOM REF] /  
CALCULATE ( DISTINCTCOUNT(BHR[ROOM REF]),ALL(BHR[MONTH] ) ) *30.4 )	
```

CALC ALL
```
=COUNT([Customer Number])
```

CALC PERCENT
```
=COUNT([Customer Number]) / [CALC ALL]
```

SUBTOTAL
```
=IF ( not ( isfiltered ( TABLE[CRM] ) ) ,  [CALC PERC] , COUNT([Customer Number] )  )
```



#### DATE TABLE 
```VBA
Date = 
//************** Script developed by RADACAD - edition: July 2021 
//************** set the variables below for your custom date table setting 
var _fromYear=2021 // set the start year of the date dimension. dates start from 1st of January of this year
var _toYear=2022   // set the end year of the date dimension. dates end at 31st of December of this year
var _startOfFiscalYear=7 // set the month number that is start of the financial year. example; if fiscal year start is July, value is 7
//************** 
var _today=TODAY()
return
ADDCOLUMNS(
    CALENDAR(
                DATE(_fromYear,1,1),
                DATE(_toYear,12,31)
),
"Year",YEAR([Date]),
"Start of Year",DATE( YEAR([Date]),1,1),
"End of Year",DATE( YEAR([Date]),12,31),
"Month",MONTH([Date]),
"Start of Month",DATE( YEAR([Date]), MONTH([Date]), 1),
"End of Month",EOMONTH([Date],0),
"Days in Month",DATEDIFF(DATE( YEAR([Date]), MONTH([Date]), 1),EOMONTH([Date],0),DAY)+1,
"Year Month Number",INT(FORMAT([Date],"YYYYMM")),
"Year Month Name",FORMAT([Date],"YYYY-MMM"),
"Day",DAY([Date]),
"Day Name",FORMAT([Date],"DDDD"),
"Day Name Short",FORMAT([Date],"DDD"),
"Day of Week",WEEKDAY([Date]),
"Day of Year",DATEDIFF(DATE( YEAR([Date]), 1, 1),[Date],DAY)+1,
"Month Name",FORMAT([Date],"MMMM"),
"Month Name Short",FORMAT([Date],"MMM"),
"Quarter",QUARTER([Date]),
"Quarter Name","Q"&FORMAT([Date],"Q"),
"Year Quarter Number",INT(FORMAT([Date],"YYYYQ")),
"Year Quarter Name",FORMAT([Date],"YYYY")&" Q"&FORMAT([Date],"Q"),
"Start of Quarter",DATE( YEAR([Date]), (QUARTER([Date])*3)-2, 1),
"End of Quarter",EOMONTH(DATE( YEAR([Date]), QUARTER([Date])*3, 1),0),
"Week of Year",WEEKNUM([Date]),
"Start of Week", [Date]-WEEKDAY([Date])+1,
"End of Week",[Date]+7-WEEKDAY([Date]),
"Fiscal Year",if(_startOfFiscalYear=1,YEAR([Date]),YEAR([Date])+ QUOTIENT(MONTH([Date])+ (13-_startOfFiscalYear),13)),
"Fiscal Quarter",QUARTER( DATE( YEAR([Date]),MOD( MONTH([Date])+ (13-_startOfFiscalYear) -1 ,12) +1,1) ),
"Fiscal Month",MOD( MONTH([Date])+ (13-_startOfFiscalYear) -1 ,12) +1,
"Day Offset",DATEDIFF(_today,[Date],DAY),
"Month Offset",DATEDIFF(_today,[Date],MONTH),
"Quarter Offset",DATEDIFF(_today,[Date],QUARTER),
"Year Offset",DATEDIFF(_today,[Date],YEAR)
)
```


# MOE
###### [LIBRARY](https://github.com/ScottypNZ/CODE-LIBRARY)   |   [INDEX](#INDEX)
-------------------------

#### MOE GENERAL

```vba
Average number of new hires = 
AVERAGEX(
	KEEPFILTERS(VALUES('New_Date Table'[Month_Year])),
	CALCULATE(DISTINCTCOUNT('Hire Data'[Employee Number]))
)
```

```vba
Average FTE of new hires = 
AVERAGEX('Hire Data','Hire Data'[FTE])
```

```vba
_Avg core = 
AVERAGEX(
    KEEPFILTERS(VALUES('New_Date Table'[Month_Year])),
    CALCULATE(DISTINCTCOUNT('All Staff'[Employee Number]),
    'All Staff'[Core Unplanned Employee] = "Yes", 
    'All Staff'[Flag] = "Include" 
    )
)
```

```vba
_Ethnicity = LOOKUPVALUE(_Ethnicity[ETHNICITY GRP],_Ethnicity[Employee Number],'All Staff'[Employee Number])
```

```vba
_Ethnicity % = divide( COUNT('All Staff'[Employee Number]),CALCULATE(COUNT('All Staff'[Employee Number]), ALL('All Staff'[ETHNICITY GRP])))
```

```vba
_FTE change MONTH = 
calculate(sum('All Staff'[FTE]),filter(all('New_Date Table'),'New_Date Table'[Month_Year]>min('New_Date Table'[Month_Year])))-
calculate(sum('All Staff'[FTE]),PREVIOUSMONTH('All Staff'[As of]),'All Staff'[As of]<>date(2023,04,03))
```

```vba
_FTE PREVIOUS = calculate(sum('All Staff'[FTE]),PREVIOUSMONTH('All Staff'[As of]))
```

```vba
_Gender_Diversity_Male = divide(calculate(COUNT('All Staff'[Employee Number]), 'All Staff'[_gender]="Male"),CALCULATE(COUNT('All Staff'[Employee Number]), ALL('All Staff'[_gender])))
```

```vba
_MonthFilter = MONTH(MAX('All Staff'[As of]))=month('All Staff'[As of])
```

```vba
_Senior Leadership_Diversity = divide( CALCULATE(COUNT('All Staff'[Employee Number]), 'All Staff'[Employment Type]="Senior Manager"),CALCULATE(COUNT('All Staff'[Employee Number]), 'All Staff'[Employment Type]="Senior Manager", ALL('All Staff'[_gender])))
```

```vba
Age Group = if('All Staff'[Age]<15, "Unknown",if(and('All Staff'[Age]>15,'All Staff'[Age]<=25), "<25", if(and('All Staff'[Age]>25,'All Staff'[Age]<=35), "25 - 35", if(and('All Staff'[Age]>35,'All Staff'[Age]<=45), "35 - 45",if(and('All Staff'[Age]>45,'All Staff'[Age]<=55), "45 - 55",if(and('All Staff'[Age]>55,'All Staff'[Age]<=65), "55 - 65",if(and('All Staff'[Age]>65,'All Staff'[Age]<=85), "Above 65","Unknown")))))))
```

```vba
Annual Leave Balance = LOOKUPVALUE(Singletons[annual_leave_balance],Singletons[Date & Employee Number], 'All Staff'[Date & Employee Number])
```

```vba
Annual Leave Balance -in weeks = Divide('All Staff'[Annual Leave Balance],5,0)
```

```vba
Annual Leave Balance Category = if('All Staff'[Gross_unplanned Employee]="No","",if(isblank('All Staff'[Annual Leave Balance]),"Not applicable",if('All Staff'[Annual Leave Balance]<=15, "15 days and less", if(and( 'All Staff'[Annual Leave Balance]>15, 'All Staff'[Annual Leave Balance]<=30), "15 to 30 days",  if(and( 'All Staff'[Annual Leave Balance]>30, 'All Staff'[Annual Leave Balance]<=50), "30 to 50 days", if('All Staff'[Annual Leave Balance]>50, "More than 50 days"))))))
```

```vba
Annual Leave Balance Category_Weeks = if('All Staff'[Gross_unplanned Employee]="No","",if(isblank('All Staff'[Annual Leave Balance -in weeks]),"Not applicable", if('All Staff'[Annual Leave Balance -in weeks]<=3, "< 3 weeks",   if(and( 'All Staff'[Annual Leave Balance -in weeks]>3,'All Staff'[Annual Leave Balance -in weeks]<=6), "3 to 6 weeks", if(and( 'All Staff'[Annual Leave Balance -in weeks]>6,'All Staff'[Annual Leave Balance -in weeks]<=10), "6 to 10 weeks",if('All Staff'[Annual Leave Balance -in weeks]>=10, "More than 10 weeks"))))))
```

```vba
Annual Leave Balance Category_Weeks = if('All Staff'[Gross_unplanned Employee]="No","",if(isblank('All Staff'[Annual Leave Balance -in weeks]),"Not applicable", if('All Staff'[Annual Leave Balance -in weeks]<=3, "< 3 weeks",   if(and( 'All Staff'[Annual Leave Balance -in weeks]>3,'All Staff'[Annual Leave Balance -in weeks]<=6), "3 to 6 weeks", if(and( 'All Staff'[Annual Leave Balance -in weeks]>6,'All Staff'[Annual Leave Balance -in weeks]<=10), "6 to 10 weeks",if('All Staff'[Annual Leave Balance -in weeks]>=10, "More than 10 weeks"))))))
```

```vba
month FTE change = 
calculate(sum('All Staff'[FTE]),filter(all('New_Date Table'),'New_Date Table'[Month_Year]>min('New_Date Table'[Month_Year])))-
calculate(sum('All Staff'[FTE]),PREVIOUSMONTH('All Staff'[As of]),'All Staff'[As of]<>date(2023,04,03))
```

```vba
_VARheadcount = CALCULATE (
                    COUNTROWS ( 'All Staff' ),
                    'All Staff'[Gross_unplanned Employee] = "Yes", 
                    'All Staff'[Flag] = "Include"
)
```

#### MOE Contractor & previous values

```vba
Check_change in contract date = if(or(ISBLANK('2 All Staff'[Contract End Date]), isblank('2 All Staff'[Contract end date at hire])),"Okay",if ('2 All Staff'[Contract End Date]>'2 All Staff'[Contract end date at hire],"Contract end date extended",if ('2 All Staff'[Contract End Date]='2 All Staff'[Contract end date at hire],"Contract end date same",if ('2 All Staff'[Contract End Date]<'2 All Staff'[Contract end date at hire],"Contract end date brought forward","Okay"))))
```

```vba
Contract end date at hire = if(isblank(LOOKUPVALUE('5 Hires'[Contract End Date],'5 Hires'[Employee Number], '2 All Staff'[Employee Number], '5 Hires'[Effective Start Date],'2 All Staff'[Hire Date])),'2 All Staff'[Contract End Date],LOOKUPVALUE('5 Hires'[Contract End Date],'5 Hires'[Employee Number], '2 All Staff'[Employee Number], '5 Hires'[Effective Start Date],'2 All Staff'[Hire Date]))
```

```vba
Contract end date by financial year = Lookupvalue('1 Date Table'[FY],'1 Date Table'[Date],'2 All Staff'[Contract End Date])
```

```vba
Contract end term extension = if('2 All Staff'[Contract extended for number of days]=0,"",if('2 All Staff'[Contract End Date]>(lookupvalue(
    '2 All Staff'[Contract End Date],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]), '2 All Staff'[Index],
 '2 All Staff'[Index]-1
    )),"Contract extension",""))
```

```vba
Contract extended for number of days = if('2 All Staff'[Check_change in contract date]="Contract end date extended", '2 All Staff'[Updated contract in days]-'2 All Staff'[Contract period in days _from date of hire],0)
```

```vba
Contract period in days _from date of hire = if ('2 All Staff'[Check_change in contract date]="Okay",0, '2 All Staff'[Contract end date at hire]-'2 All Staff'[Hire Date])
```

```vba
Contractor calc = COUNTROWS(FILTER('2 All Staff',[Employee Type (groups)]="Contractor")) / COUNTROWS('2 All Staff')
```

```vba
Contractors = COUNTROWS(FILTER('2 All Staff',[Employee Type (groups)]="Contractor"))
```

```vba
Next Business Group = lookupvalue(
    '2 All Staff'[Business Group],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]+1
    )
```

```vba
Next SSC Flag = lookupvalue(
    '2 All Staff'[Flag],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]+1
    )
```

```vba
Next Unit = lookupvalue(
    '2 All Staff'[Unit],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]+1
    )
```

```vba
Next_Division = lookupvalue(
    '2 All Staff'[Position Title],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]), '2 All Staff'[Index],
 '2 All Staff'[Index]+1
    )
```

```vba
Next_EmpType = lookupvalue(
    '2 All Staff'[Employee Type],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]), '2 All Staff'[Index],
 '2 All Staff'[Index]+1
    )
```

```vba
Next_Position code = lookupvalue(
    '2 All Staff'[Position code],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]+1
    )
```

```vba
Next_Position title = lookupvalue(
    '2 All Staff'[Position Title],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]), '2 All Staff'[Index],
 '2 All Staff'[Index]+1
    )
```

```vba
Original contract period = if('2 All Staff'[Contract period in days _from date of hire]=0,"" ,if('2 All Staff'[Contract period in days _from date of hire]<=90,"3 months or less", if (and ('2 All Staff'[Contract period in days _from date of hire]>90,'2 All Staff'[Contract period in days _from date of hire]<180),"3 to 6 months", if (and ('2 All Staff'[Contract period in days _from date of hire]>180,'2 All Staff'[Contract period in days _from date of hire]<=365),"6 to 12 months", if (and ('2 All Staff'[Contract period in days _from date of hire]>365,'2 All Staff'[Contract period in days _from date of hire]<=550),"12 to 18 months", if (and ('2 All Staff'[Contract period in days _from date of hire]>550,'2 All Staff'[Contract period in days _from date of hire]<=730),"18 to 24 months", "24 months or more"))))))
```

```vba
Position creation date = lookupvalue('3 Position analysis'[Position creation date], '3 Position analysis'[Position Code], '2 All Staff'[Position code], 0)
```

```vba
Position Status = 
VAR _Poscode = '2 All Staff'[Position code]
VAR _FirstRec =
    CALCULATE (
        MIN ( '2 All Staff'[As of] ),
        FILTER ( '2 All Staff', '2 All Staff'[Position code] = _Poscode )
    )
RETURN
    //407367
    IF ( '2 All Staff'[As of] > _FirstRec, "Old", "New" )
```

```vba
Prev_Employee_status = lookupvalue(
    '2 All Staff'[Employee Status],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]-1
    )
```

```vba
Prev_EmpType = lookupvalue(
    '2 All Staff'[Final employee type (groups)],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]-1
    )
```

```vba
Prev_FTE = lookupvalue(
    '2 All Staff'[FTE],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]-1
    )
```

```vba
Prev_SSC Flag = lookupvalue(
    '2 All Staff'[Flag],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]-1
    )
```

```vba
Previous Business Group = lookupvalue(
    '2 All Staff'[Business Group],
    '2 All Staff'[Employee Number],value('2 All Staff'[Employee Number]),
    '2 All Staff'[Index], '2 All Staff'[Index]-1
    )
```

```vba
Previous Leave Sattus = search("Leave", '2 All Staff'[Prev_Employee_status],1,0)
```

```vba
 Average number of position = 
Averagex(
    KEEPFILTERS(values('3 Position analysis'[As of])),
    calculate(distinctcount('3 Position analysis'[Position Code])))
```

```vba
 Updated contract period = if('2 All Staff'[Contract period in days _from date of hire]=0,"" ,if('2 All Staff'[Updated contract in days]<=90,"3 months or less", if (and ('2 All Staff'[Updated contract in days]>90,'2 All Staff'[Updated contract in days]<180),"3 to 6 months", if (and ('2 All Staff'[Updated contract in days]>180,'2 All Staff'[Updated contract in days]<=365),"6 to 12 months", if (and ('2 All Staff'[Updated contract in days]>365,'2 All Staff'[Updated contract in days]<=550),"12 to 18 months", if (and ('2 All Staff'[Updated contract in days]>550,'2 All Staff'[Updated contract in days]<=730),"18 to 24 months", "24 months or more"))))))
```

```vba
 Updated contract in days = if ('2 All Staff'[Check_change in contract date]="Okay",0, '2 All Staff'[Contract end date]-'2 All Staff'[Hire Date])
```

```vba
 Transfer to other Unit = if('2 All Staff'[unit]<>'2 All Staff'[Next Unit]&& '2 All Staff'[Flag]="Include"&& '2 All Staff'[Next SSC Flag]="Include"&& '2 All Staff'[Next Unit]<>"","Transfer to other business unit")
```

```vba
 Transfer to other BG = if('2 All Staff'[Business Group]<>'2 All Staff'[Next Business Group]&& '2 All Staff'[Flag]="Include"&& '2 All Staff'[Next SSC Flag]="Include"&& '2 All Staff'[Next Business Group]<>"","Transfer to other business group")
```

```vba
 Transfer from other BG = if('2 All Staff'[Business Group]<>'2 All Staff'[Previous Business Group]&& '2 All Staff'[Flag]="Include"&& '2 All Staff'[Prev_SSC Flag]="Include","Transfer from other business group")
```

```vba
 Termination status = IF (
                '2 All Staff'[Next_EmpType] = "" && '2 All Staff'[Flag]="Include"
                    && '2 All Staff'[As of] <> MAX ( '2 All Staff'[As of] ),
                "Terminated")
```

```vba
 Termination date = LOOKUPVALUE('6 Terminations'[Event Date], '6 Terminations'[Employee Number], '2 All Staff latest manual'[Employee Number], '6 Finishers'[Position code], '2 All Staff latest manual'[Position number])
```

```vba
 Term Month = if('2 All Staff'[As of]=eomonth('2 All Staff'[Termination date],0),'2 All Staff'[As of], if('2 All Staff'[As of]=max('2 All Staff'[As of]),EOMONTH('2 All Staff'[Termination date],0),if('2 All Staff'[Termination status]="Terminated", EOMONTH('2 All Staff'[As of]+1,0),if('2 All Staff'[As of]=max('2 All Staff'[As of]),EOMONTH('2 All Staff'[Termination date],0)))))
```

```vba
 No. of position change = 
calculate(distinctcount('3 Position analysis'[Position Code])-
calculate(distinctcount('3 Position analysis'[Position Code]),PREVIOUSMONTH('3 Position analysis'[As of])))
```

```vba
 Position vacancy period = if(isblank('3 Position analysis'[Position vacant since (days)]),"", if ('3 Position analysis'[Position vacant since (days)]<=90,"0 to 3 months", if (and('3 Position analysis'[Position vacant since (days)]>90,'3 Position analysis'[Position vacant since (days)]<=180),"3 to 6 months", if (and('3 Position analysis'[Position vacant since (days)]>180,'3 Position analysis'[Position vacant since (days)]<=365),"6 to 12 months", "More than 12 months"))))
```


```vba
Senior Leadership_Diversity_Gender = divide( CALCULATE(COUNT('All Staff'[Employee Number]), 'All Staff'[Employment Type]="Senior Manager"),CALCULATE(COUNT('All Staff'[Employee Number]), 'All Staff'[Employment Type]="Senior Manager", ALL('All Staff'[_gender])))
```

```vba
Measure = switch (SELECTEDMEASURENAME(), [Starting_FTE], [Ending_FTE], [New Starters_New Position], [Terminations],BLANK())
```

```vba
Fixed Term to ALl Staff = divide( CALCULATE(COUNT('All Staff'[Employee Number]), 'All Staff'[Employee Type (groups)]="Fixed Term"),CALCULATE(COUNT('All Staff'[Employee Number]), ALL('All Staff'[Employee Type (groups)])))
```

```vba
Ending FTE = 
CALCULATE(
    SUM( 'All Staff'[FTE] ),
    LASTDATE( 'All Staff'[As of] ))
```

```vba
Employee to Manager = divide(( CALCULATE(COUNT('All Staff'[Employee Number]), 'All Staff'[Employment Type (groups)]="Manager")+CALCULATE(COUNT('All Staff'[Employee Number]), 'All Staff'[Employment Type (groups)]="Senior Manager")+CALCULATE(COUNT('All Staff'[Employee Number]), 'All Staff'[Employment Type (groups)]="Service Manager")),CALCULATE(COUNT('All Staff'[Employee Number]), ALL('All Staff'[Employment Type (groups)]))
```

```vba
_VARTurnoverX = 
//Calculates current headcount 
VAR Headcount1 = CALCULATE (
                    COUNTROWS ( 'All Staff' ),
                    'All Staff'[Gross_unplanned Employee] = "Yes", 
                    'All Staff'[Flag] = "Include"
)


VAR Exits1 =
    CALCULATE (
        SWITCH (
            TRUE (),
            COUNTROWS ( 'WFG Exits (New)' ) <> 0, COUNTROWS ( 'WFG Exits (New)' ),
            0
        ),
        DATESINPERIOD (
            'New_Date Table'[Date],
            LASTDATE ( 'New_Date Table'[Date] ),
            -1,
            YEAR
        )
    )
RETURN
  
             DIVIDE ( Exits1, Headcount1, BLANK ()  )
	 
```

```vba
_VARTurnover_Gross = 
//Calculates current headcount 
VAR _CurrHC = CALCULATE (
                    COUNTROWS ( 'All Staff' ),
                    'All Staff'[Gross_unplanned Employee] = "Yes", 
                    'All Staff'[Flag] = "Include"
)
//Calculates headcount 12 months ago
VAR _PrevHC = CALCULATE (
                        COUNTROWS ( 'All Staff' ),
                        'All Staff'[Gross_unplanned Employee] = "Yes",
                        'All Staff'[Flag] = "Include",
                        PARALLELPERIOD ( 'New_Date Table'[Date], -12, MONTH )
)
//Calculates average headcount
VAR _HeadcountAve =
    SWITCH (
        TRUE (),
        _PrevHC <> 0,
            DIVIDE (_CurrHC + _PrevHC,
                2,
                BLANK ()
            ),
        _CurrHC
    )
VAR _Exits =
    CALCULATE (
        SWITCH (
            TRUE (),
            COUNTROWS ( 'WFG Exits (New)' ) <> 0, COUNTROWS ( 'WFG Exits (New)' ),
            0
        ),
        DATESINPERIOD (
            'New_Date Table'[Date],
            LASTDATE ( 'New_Date Table'[Date] ),
            -1,
            YEAR
        )
    )
RETURN
   IF (
       _CurrHC = 0,
         "",
             DIVIDE ( _Exits, _HeadcountAve, BLANK () 
    )
)
```

```vba

 Status = 
if('2 All Staff'[Index]=2,"June 2019 data",    If('2 All Staff'[Flag]="Include"&&'2 All Staff'[Prev_SSC Flag]=""&& '2 All Staff'[Position Status]="Old","New Starter_Existing Position",                                                                                                                                                                                                                                If('2 All Staff'[Flag]="Include"&&'2 All Staff'[Prev_SSC Flag]=""&& '2 All Staff'[Position Status]="New","New Starter_New Position",                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    if('2 All Staff'[Flag]="Exclude" && '2 All Staff'[Prev_SSC Flag]="Include" && '2 All Staff'[EMployee on leave]<>0,"Started Leave",                                    IF('2 All Staff'[Employee Status]="Active"&&'2 All Staff'[Prev_SSC Flag]="Exclude"&& '2 All Staff'[Flag]="Include"&&'2 All Staff'[Previous Leave Sattus]<>0, "Back from leave",                                                                                                                                                                  if('2 All Staff'[Flag]="Include"&& ('2 All Staff'[Employee Type (groups)]="Permanent"||'2 All Staff'[Employee Type (groups)]="Fixed Term") &&'2 All Staff'[Prev_SSC Flag]="Exclude"&& ('2 All Staff'[Prev_EmpType]="Contractor"||'2 All Staff'[Prev_EmpType]="Casual"|| '2 All Staff'[Prev_EmpType]="Temporary"|| '2 All Staff'[Prev_EmpType]="Seconded into the Ministry"|| '2 All Staff'[Prev_EmpType]="Non-contractor External Service Provide"),"Conversion to Permanent/Fixed Term", 
                                                                                                                                                                                                                                          if('2 All Staff'[Prev_SSC Flag]="Include"&& ('2 All Staff'[Prev_EmpType]="Permanent"||'2 All Staff'[Prev_EmpType]="Fixed Term"||'2 All Staff'[Prev_EmpType]="Internal Secondment") &&'2 All Staff'[Flag]="Exclude"&& ('2 All Staff'[Employee Type (groups)]="Contractor"||'2 All Staff'[Employee Type (groups)]="Casual"|| '2 All Staff'[Employee Type (groups)]="Temporary"|| '2 All Staff'[Employee Type (groups)]="Seconded into the Ministry"|| '2 All Staff'[Employee Type (groups)]="Non-contractor External Service Provide"),"Conversion from Permanent/Fixed Term",                                                                                                                                                                                                                                                                                                                                                                                                                                              if ('2 All Staff'[FTE]<>'2 All Staff'[Prev_FTE]&&'2 All Staff'[Flag]="Include"&& '2 All Staff'[Prev_SSC Flag]="Include" && '2 All Staff'[Employee Status]="Active", "Change in FTE"                                          ))))))))
```

#### WATERFALL

```vba

Waterfall = switch (SELECTEDVALUE('Status for waterfall'[Index]),                                                                                                                                                            1, [Starting_FTE], 
2, [New Starters_New Position], 
3,  [New Starters_Existing Position],
4, [Back from leave],
5, [Change in Hours],
6, [Conversion to permanent/fixed-term],
7, [Started external secondment],
8, [Back from external secondment],
9, [Terminations],
10, [Conversion from permanent/fixed-term],
11, [Started leave],
12, [Ending_FTE],
Blank()
)

```

```vba
Terminations = -1*
CALCULATE(
    sum( 'All Staff'[FTE] ), 'All Staff'[Flag]="Include", 'All Staff'[Status]="Terminated",PREVIOUSMONTH(
 LASTDATE('All Staff'[As of]))
    )
```

```vba
Starting_Vacancy = 
CALCULATE(
    SUM( 'All Staff'[FTE]), 'All Staff'[Flag]="Include",
    FIRSTDATE( 'All Staff'[As of] )
)
```

```vba
Starting_FTE = 
CALCULATE(
    SUM( 'All Staff'[FTE] ), 'All Staff'[Flag]="Include",PREVIOUSMONTH(
  LASTDATE('All Staff'[As of])
    ))
```


```vba
Started leave = -1*
CALCULATE(
  sum( 'All Staff'[Prev_FTE] ),   'All Staff'[Status]="Started Leave",
  LASTDATE('All Staff'[As of])
    )
```

```vba
Started external secondment = -1*
CALCULATE(
   sum( 'All Staff'[FTE] ),'All Staff'[Flag]="Include",  'All Staff'[Status]="Started External Secondment",
  LASTDATE('All Staff'[As of])
    )
```


```vba
New Starters_New Position = 
CALCULATE(
    sum( 'All Staff'[FTE] ),'All Staff'[Flag]="Include",  'All Staff'[Status]="New Starter_New Position",
  LASTDATE('All Staff'[As of])
    )
```


```vba
New Starters_Existing Position = 
CALCULATE(
    sum( 'All Staff'[FTE] ),'All Staff'[Flag]="Include",  'All Staff'[Status]="New Starter_Existing Position",
  LASTDATE('All Staff'[As of])
    )
```


```vba
Ending_FTE = 
CALCULATE(
    SUM( 'All Staff'[FTE] ), 'All Staff'[Flag]="Include",
   LASTDATE('All Staff'[As of])
    )
```

```vba
Conversion to permanent/fixed-term = 
CALCULATE(
    sum( 'All Staff'[FTE] ),'All Staff'[Flag]="Include",   'All Staff'[Status]="Conversion to permanent/fixed term",
 LASTDATE('All Staff'[As of])
    )
```

```vba
Conversion from permanent/fixed-term = -1*
CALCULATE(
    sum( 'All Staff'[Prev_FTE] ), 'All Staff'[Status]="Conversion from permanent/fixed term",
 LASTDATE('All Staff'[As of])
    )
```

```vba
Change in Hours = 
CALCULATE(
    sum( 'All Staff'[FTE] ),'All Staff'[Flag]="Include",  'All Staff'[Status]="Change in FTE",
 LASTDATE('All Staff'[As of]))-CALCULATE(
    sum( 'All Staff'[Prev_FTE] ),'All Staff'[Flag]="Include",  'All Staff'[Status]="Change in FTE",
 LASTDATE('All Staff'[As of]))
```

```vba
Back from leave = 
CALCULATE(
    sum( 'All Staff'[FTE] ),'All Staff'[Flag]="Include",  'All Staff'[Status]="Back from leave",
  LASTDATE('All Staff'[As of])
    )
```

```vba
Back from external secondment = 
CALCULATE(
    sum( 'All Staff'[FTE] ),'All Staff'[Flag]="Include",  'All Staff'[Status]="Back from external secondment",
  LASTDATE('All Staff'[As of])
    )
```

#### MAORI Capability

```vba
_Rehire = if('Exits'[Hire Date]>'Exits'[Last Date Worked],"Rehire","")
```

```vba
_EmployeeType = 
if ( 'All Staff'[Employee Type] = "Casual", "Casual",
if ( 'All Staff'[Employee Type] = "Contractor - Backfill OPEX", "Contractor",
if ( 'All Staff'[Employee Type] = "Contractor - CAPEX", "Contractor",
if ( 'All Staff'[Employee Type] = "Contractor - Project OPEX", "Contractor",
if ( 'All Staff'[Employee Type] = "Fixed Term", "Fixed Term",
if ( 'All Staff'[Employee Type] = "Temporary", "Temporary",
if ( 'All Staff'[Employee Type] = "Seconded into the Ministry", "Seconded into the Ministry",
if ( 'All Staff'[Employee Type] = "External Secondment", "Permanent",
if ( 'All Staff'[Employee Type] =  "External Secondment - Ministry Office", "Permanent",
if ( 'All Staff'[Employee Type] =  "Internal Secondment", "Permanent",
if ( 'All Staff'[Employee Type] =  "Permanent - Transitional", "Permanent",
if ( 'All Staff'[Employee Type] = "Permanent", "Permanent",
if ( 'All Staff'[Employee Type] = "Other External Service Providers", "Other External Service Provider",
if ( 'All Staff'[Employee Type] = "Other External Service Providers", "Professional Services / Specialist Provider (non-contractor)","XXXXXX"
))))))))))))))
```

```vba
_Index = countrows (filter(values('All Staff'[As of]), 'All Staff'[As of]<EARLIER('All Staff'[As of])+1))
```

```vba
_Maori = 
if ( 'All Staff'[Ethnicity First] = "Maori" , "Maori",
if ( 'All Staff'[Ethnicity Second] = "Maori" , "Maori",
if ( 'All Staff'[Ethnicity Third] = "Maori" , "Maori",
if ( 'All Staff'[Ethnicity First] = "" , "Unknown","Non-Maori")
)))
```

```vba
_Tenure = if( (DATEDIFF( 'All Staff'[Hire Date],'All Staff'[As of],MONTH)/12)<0, DATEDIFF('All Staff'[Hire Date],'All Staff'[As of],MONTH)/12, DATEDIFF( 'All Staff'[Hire Date],'All Staff'[As of],MONTH)/12)
```

```vba
_Tenure Group = if('All Staff'[_Tenure]<1,"0 - 1 year", if(and('All Staff'[_Tenure]>=1,'All Staff'[_Tenure]<2),"1 - 2 years", if(and('All Staff'[_Tenure]>=2,'All Staff'[_Tenure]<5),"2 - 5 years",if(and('All Staff'[_Tenure]>=5,'All Staff'[_Tenure]<10),"5 - 10 years","More than 10 years"))))
```

```vba
_Tenure_leavers = if('Exits'[_Tenure]<1, "0 - 1 year", if(and('Exits'[_Tenure]>=1, 'Exits'[_Tenure]<2), "1 - 2 years", if(and('Exits'[_Tenure]>=2, 'Exits'[_Tenure]<5), "2 - 5 years",  if(and('Exits'[_Tenure]>=5,'Exits'[_Tenure]<10), "5 - 10 years", "More than 10 years"))))
```

##### CUMULATIVE

```vba
_WEEKLY OPEN = [_WEEKLY CREATED running] - [_WEEKLY RESOLVED running]
```

```vba
_WEEKLY CREATED running = 
CALCULATE(
	SUM('_PIVOT'[CREATED]),
	FILTER(
		ALLSELECTED('_PIVOT'[WEEK]),
		ISONORAFTER('_PIVOT'[WEEK], MAX('_PIVOT'[WEEK]), DESC)
	)
)
```

```vba
_WEEKLY RESOLVED running = 
CALCULATE(
	SUM('_PIVOT'[RESOLVED]),
	FILTER(
		ALLSELECTED('_PIVOT'[WEEK]),
		ISONORAFTER('_PIVOT'[WEEK], MAX('_PIVOT'[WEEK]), DESC)
	)
)
```

```vba
_DAY CREATED = 

CALCULATE(
	SUM('_PIVOT'[CREATED]),
	FILTER(
		ALLSELECTED('_PIVOT'[Date]),
		ISONORAFTER('_PIVOT'[Date], MAX('_PIVOT'[Date]), DESC)
	)
)
```

```vba

```


# DATE

###### [LIBRARY](https://github.com/ScottypNZ/CODE-LIBRARY)   |   [INDEX](#INDEX)
-------------------------

```vba
 Date = 
//************** Script developed by RADACAD - edition: July 2021 
//************** set the variables below for your custom date table setting 
var _fromYear=2022 // set the start year of the date dimension. dates start from 1st of January of this year
var _toYear=2025   // set the end year of the date dimension. dates end at 31st of December of this year
var _startOfFiscalYear=7 // set the month number that is start of the financial year. example; if fiscal year start is July, value is 7
//************** 
var _today=TODAY()
return
ADDCOLUMNS(
    CALENDAR(
                DATE(_fromYear,1,1),
                DATE(_toYear,12,31)
),
"Year",YEAR([Date]),
"Start of Year",DATE( YEAR([Date]),1,1),
"End of Year",DATE( YEAR([Date]),12,31),

"Start of Month",DATE( YEAR([Date]), MONTH([Date]), 1),
"End of Month",EOMONTH([Date],0),

"Year Month Number",FORMAT([Date],"YYYY/MM"),
"Year Month Name",FORMAT([Date],"YYYY-MMM"),

"Day Name",FORMAT([Date],"DDDD"),
"Day Name Short",FORMAT([Date],"DDD"),

"Month Name",FORMAT([Date],"MMMM"),
"Month Name Short",FORMAT([Date],"MMM"),
"Start of Week", [Date]-WEEKDAY([Date])+1,
"End of Week",[Date]+7-WEEKDAY([Date]),

"Month",MONTH([Date]),
"Day",DAY([Date]),
"Day of Week",WEEKDAY([Date]),
"Day of Year",DATEDIFF(DATE( YEAR([Date]), 1, 1),[Date],DAY)+1,
"Week of Year",WEEKNUM([Date]),
"Days in Month",DATEDIFF(DATE( YEAR([Date]), MONTH([Date]), 1),EOMONTH([Date],0),DAY)+1
)
```



