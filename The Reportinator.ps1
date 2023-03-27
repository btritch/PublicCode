

.'Z:\RiskII\Tritch\Operation Loki\Functions.ps1'
.'Z:\RiskII\Tritch\Operation Loki\The Reportinator Reports.ps1'

$vcCCL_Prod_Conn = 'ODBC;DSN=CCL_PROD;UID=l00b3t;DBALIAS=CCL_PROD;DATABASE=bigsql;PORT=32051;HOSTNAME=11.48.219.145;CONNECTTIMEOUT=3600;'

[datetime] $dtToday = (get-date).ToString("MM/dd/yy") # '04/16/19' #
$vcToday = $dtToday.ToString('yyyyMMdd')

[datetime] $dtSatLastWeek = $dtToday.AddDays(-($dtToday.DayOfWeek.value__ + 1))
[int] $iDayOfWeekValue = $dtToday.DayOfWeek.value__

$iUselessBookColumns = 4

$z = 2 #used to run monthly reports early. 2 is default value, 1 is early run.

$arrReports = $null
$arrReports += 
(  
<#
('Non_Mon_Trans',        26, 2, 1, ($iUselessBookColumns +  1), 'W', 0),   
('Straight_Rollers',     6 , 2,$z, ($iUselessBookColumns +  2), 'M', 0),   
('NA_Waterfall',         26, 3, 1, ($iUselessBookColumns +  3), 'W', 0),   
('TBD1',                 1 , 2, 1, ($iUselessBookColumns +  4), 'W', 0),   
('TBD2',                 1 , 2,$z, ($iUselessBookColumns +  5), 'M', 0),#>   
('NA_Choice',            26, 2, 1, ($iUselessBookColumns +  6), 'W', 0),   
('NA_Cred_Qual',         6 , 2,$z, ($iUselessBookColumns +  7), 'M', 0),   
('Acct_Bal_FICO',        6 , 2,$z, ($iUselessBookColumns +  8), 'M', 0),   
('Prepaid_MIS',          6 , 2,$z, ($iUselessBookColumns +  9), 'M', 0), 
('Corp_Guarantor',       6 , 2,$z, ($iUselessBookColumns + 10), 'M', 0),  
('NA_Waterfall',         26, 2, 1, ($iUselessBookColumns + 11), 'W', 0),<#   
('NA_Choice_CCL',        26, 2, 1, ($iUselessBookColumns + 12), 'W', 0),#>   
('Daily_Auth_Tracking',  26, 5, 1, ($iUselessBookColumns + 13), 'W', 0),   
('Portfolio_Vintage',    6 , 2,$z, ($iUselessBookColumns + 14), 'M', 0), 
('NA_Vintage',           6 , 2,$z, ($iUselessBookColumns + 15), 'M', 0),
('FDS_Cash_Payments',    14, 2, 1, ($iUselessBookColumns + 16), 'D', 0),   
('FDS_CB_Refunds',       14, 2, 1, ($iUselessBookColumns + 17), 'W', 0),   
('Over_Limit',           6 , 2,$z, ($iUselessBookColumns + 18), 'M', 0),   
('Credit_Line_Change',   6 , 2,$z, ($iUselessBookColumns + 19), 'M', 0),   
('FICO_Score_Migration', 6 , 2,$z, ($iUselessBookColumns + 20), 'M', 0),
('Auth_Waterfall',       26, 5,$z, ($iUselessBookColumns + 21), 'W', 0),   
('Portfolio_Accounts',   6 , 2, 2, ($iUselessBookColumns + 22), 'M', 0),<#   
('Overlimit_Trans',      26, 2, 2, ($iUselessBookColumns + 23), 'W', 0),#>   
('Loss_Forecast',        6 , 2,$z, ($iUselessBookColumns + 24), 'M', 0), 
('NA_Line_Exposure',     6 , 2,$z, ($iUselessBookColumns + 25), 'M', 0),   
('Return_Checks',        6 , 2,$z, ($iUselessBookColumns + 26), 'M', 0),   
('FDS_Return_Checks',    4 , 2, 2, ($iUselessBookColumns + 27), 'W', 0),<#   
('FDS_OFAC',             1 , 2,$z, ($iUselessBookColumns + 28), 'M', 0),   
('Delq_Report',          6 , 2,$z, ($iUselessBookColumns + 29), 'M', 0),#>   
('Data_Validation',      1 , 2, 2, ($iUselessBookColumns + 30), 'D', 0),   
('POA_Certegy',          6 , 2,$z, ($iUselessBookColumns + 31), 'M', 0),
('POS_CLI',              6 , 2,$z, ($iUselessBookColumns + 32), 'M', 0),
('NA_VINTAGE_6',         6 , 2,$z, ($iUselessBookColumns + 33), 'M', 0),
('NA_IDA_Vintage_3',     6 , 2,$z, ($iUselessBookColumns + 34), 'M', 0),
('NA_IDA_Vintage_6',     13, 2,$z, ($iUselessBookColumns + 35), 'M', 0),<#
('NA_Referrals',         6 , 2,$z, ($iUselessBookColumns + 36), 'M', 0),
('NA_MTM_Comp',          6 , 2,$z, ($iUselessBookColumns + 37), 'M', 0),#>
('Ninty_Plus_Curve',     6 , 2,$z, ($iUselessBookColumns + 38), 'M', 0),<#
('IDA_Curve',            6 , 2,$z, ($iUselessBookColumns + 39), 'M', 0),#>
('NA_13_Month_Waterfall',6 , 2,$z, ($iUselessBookColumns + 40), 'M', 0),
('NA_Purchase_Activity', 6 , 2,$z, ($iUselessBookColumns + 41), 'M', 0),
('NA_Fraud',             26, 2, 1, ($iUselessBookColumns + 42), 'W', 0),
('NA_Referrals',         26, 2, 1, ($iUselessBookColumns + 43), 'W', 0),
('GCL_Waterfall',        6 , 2, 2, ($iUselessBookColumns + 44), 'M', 0),
('Corp_Acct_Net_Purch',  6,  2,$z, ($iUselessBookColumns + 45), 'M', 0)<#,
('', 1, 2, 2, ($iUselessBookColumns + 46), 'M', 0),
('', 1, 2, 2, ($iUselessBookColumns + 47), 'M', 0),
('', 1, 2, 2, ($iUselessBookColumns + 48), 'M', 0),
('', 1, 2, 2, ($iUselessBookColumns + 49), 'M', 0),
('', 1, 2, 2, ($iUselessBookColumns + 50), 'M', 0)#>
)

<#
0 - Report Name & Function (without fn) Name / 1 - Max Runs / 2 - Run Day / 3 - Run Week / 
4 - Column / 5 - Weekly/Daily/Monthly / 6 - Current # of Runs 
#>

$objExcel = New-Object -comobject Excel.Application #Enables the use of Excel files
$objExcel.Visible = $true #Excel is not visible (runs in the background)
$objExcel.DisplayAlerts = $false #Alerts from Excel are not displayed  

$vcTrackerName = 'Z:\RiskII\Risk Analytics\Report Tracker2.xlsx'

$wbTracker = $ObjExcel.Workbooks.Open($vcTrackerName, 2, $false) #Open Excel - Workbook Path & Name, Something, Read Only, Something, Password
$wsTracker = $wbTracker.worksheets.item('Tracker')

#
#$vcDateSQL = "SELECT MAX(PART_DATE) AS MAX_PART_DATE FROM CRD_MSTR.PDW_MMAST_EXTRACT_TEXT"
#
#$vcMaxPart_Date = fnInvokeSQLcmdDSN 'CCL_Prod' $vcDateSQL 0

$dtDates = $null

$vcDateSQL = 
"
    SELECT
        SUB1.STANDARD_DATE
        ,SUB1.DATE_KEY AS CUR_DATE
        ,SUB1.DAY_OF_WEEK AS CUR_DAY_OF_WEEK
        ,SUB1.PREV_DATE

        ,SUB1.YEAR AS CUR_YEAR
        ,SUB1.MONTH AS CUR_MONTH
        ,SUB1.MONTH_NAME AS CUR_MONTH_NAME

        ,SUB1.FISCAL_YEAR AS CUR_FISCAL_YEAR
        ,SUB1.FISCAL_MONTH AS CUR_FISCAL_MONTH
        ,SUB1.FISCAL_MONTH_NAME AS CUR_FISCAL_MONTH_NAME
        ,SUB1.FISCAL_WEEK_OF_MONTH AS CUR_FISCAL_WEEK_OF_MONTH

        ,FISCAL3.YEAR AS YEAR_PREV_WEEK
        ,FISCAL3.MONTH AS WEEK_PREV_WEEK
        ,FISCAL3.MONTH_NAME AS MONTH_NAME_PREV_WEEK

        ,FISCAL3.FISCAL_YEAR AS FISCAL_YEAR_PREV_WEEK
        ,FISCAL3.FISCAL_MONTH AS FISCAL_MONTH_PREV_WEEK
        ,FISCAL3.FISCAL_MONTH_NAME AS FISCAL_MONTH_NAME_PREV_WEEK
        ,FISCAL3.FISCAL_WEEK_OF_MONTH AS FISCAL_WEEK_OF_MONTH_PREV_WEEK

        ,FISCAL2.YEAR AS YEAR_PREV_MONTH
        ,FISCAL2.MONTH AS MONTH_PREV_MONTH
        ,FISCAL2.MONTH_NAME AS MONTH_NAME_PREV_MONTH

        ,FISCAL2.FISCAL_YEAR AS FISCAL_YEAR_PREV_MONTH
        ,FISCAL2.FISCAL_MONTH AS FISCAL_MONTH_PREV_MONTH
        ,FISCAL2.FISCAL_MONTH_NAME AS FISCAL_MONTH_NAME_PREV_MONTH

        ,SUB1.FIRST_DAY_OF_MONTH AS CUR_FIRST_DAY_OF_MONTH
        ,SUB1.LAST_DAY_OF_MONTH AS CUR_LAST_DAY_OF_MONTH
        ,SUB1.FIRST_DAY_OF_PREV_MONTH 
        ,SUB1.LAST_DAY_OF_PREV_MONTH

        ,SUB1.M01
        ,SUB1.M02
        ,SUB1.M03
        ,SUB1.M04
        ,SUB1.M05
        ,SUB1.M06
        ,SUB1.M07
        ,SUB1.M08
        ,SUB1.M09
        ,SUB1.M10
        ,SUB1.M11
        ,SUB1.M12
        ,SUB1.M13
        ,SUB1.M14
        ,SUB1.M15
        ,SUB1.M16
        ,SUB1.M17
        ,SUB1.M18
        ,SUB1.M19
        ,SUB1.M20
        ,SUB1.M21
        ,SUB1.M22
        ,SUB1.M23
        ,SUB1.M24

        ,FISCAL3.FISCAL_QUARTER AS FISCAL_QUARTER_PREV_WEEK
        ,FISCAL2.FISCAL_QUARTER AS FISCAL_QUARTER_PREV_MONTH

        ,SUB1.M25
        ,SUB1.M26
        ,SUB1.M27
        ,SUB1.M28
        ,SUB1.M29
        ,SUB1.M30
        ,SUB1.M31
        ,SUB1.M32
        ,SUB1.M33
        ,SUB1.M34
        ,SUB1.M35
        ,SUB1.M36
    FROM
    (
        SELECT
            STANDARD_DATE                                                                                   
            ,DATE_KEY
            ,DAY_OF_WEEK
            ,VARCHAR_FORMAT(FIRST_DAY_OF_MONTH,'YYYYMMDD') AS FIRST_DAY_OF_MONTH                            
            ,VARCHAR_FORMAT(LAST_DAY_OF_MONTH,'YYYYMMDD') AS LAST_DAY_OF_MONTH                              
            ,FISCAL_YEAR                                                                                    
            ,FISCAL_MONTH
            ,FISCAL_MONTH_NAME
            ,FISCAL_WEEK_OF_MONTH
            ,YEAR
            ,MONTH
            ,MONTH_NAME
            ,VARCHAR_FORMAT(ADD_DAYS(STANDARD_DATE, - DAY_OF_WEEK),'YYYYMMDD') AS EOW_PREV_WEEK
            ,VARCHAR_FORMAT(ADD_DAYS(STANDARD_DATE,-1),'YYYYMMDD') AS PREV_DATE

            ,VARCHAR_FORMAT(ADD_MONTHS(FIRST_DAY_OF_MONTH, -1),'YYYYMMDD') AS FIRST_DAY_OF_PREV_MONTH
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -1),'YYYYMMDD') AS LAST_DAY_OF_PREV_MONTH

            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -2),'YYYYMMDD') AS M01
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -3),'YYYYMMDD') AS M02
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -4),'YYYYMMDD') AS M03
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -5),'YYYYMMDD') AS M04
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -6),'YYYYMMDD') AS M05
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -7),'YYYYMMDD') AS M06
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -8),'YYYYMMDD') AS M07
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -9),'YYYYMMDD') AS M08
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -10),'YYYYMMDD') AS M09
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -11),'YYYYMMDD') AS M10
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -12),'YYYYMMDD') AS M11
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -13),'YYYYMMDD') AS M12
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -14),'YYYYMMDD') AS M13
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -15),'YYYYMMDD') AS M14
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -16),'YYYYMMDD') AS M15
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -17),'YYYYMMDD') AS M16
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -18),'YYYYMMDD') AS M17
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -19),'YYYYMMDD') AS M18
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -20),'YYYYMMDD') AS M19
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -21),'YYYYMMDD') AS M20
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -22),'YYYYMMDD') AS M21
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -23),'YYYYMMDD') AS M22
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -24),'YYYYMMDD') AS M23
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -25),'YYYYMMDD') AS M24
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -26),'YYYYMMDD') AS M25
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -27),'YYYYMMDD') AS M26
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -28),'YYYYMMDD') AS M27
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -29),'YYYYMMDD') AS M28
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -30),'YYYYMMDD') AS M29
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -31),'YYYYMMDD') AS M30
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -32),'YYYYMMDD') AS M31
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -33),'YYYYMMDD') AS M32
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -34),'YYYYMMDD') AS M33
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -35),'YYYYMMDD') AS M34
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -36),'YYYYMMDD') AS M35
            ,VARCHAR_FORMAT(ADD_MONTHS(LAST_DAY_OF_MONTH, -37),'YYYYMMDD') AS M36
        FROM
            REFERENCEDATA.DBO_DATE

        WHERE DATE_KEY BETWEEN '20180101' AND '$vcToday' --(SELECT CURRENT DATE FROM sysibm.sysdummy1)
    ) AS SUB1

    LEFT JOIN REFERENCEDATA.DBO_DATE AS FISCAL2
    ON FISCAL2.DATE_KEY = SUB1.LAST_DAY_OF_PREV_MONTH

    LEFT JOIN REFERENCEDATA.DBO_DATE AS FISCAL3
    ON FISCAL3.DATE_KEY = SUB1.EOW_PREV_WEEK

    ORDER BY
        SUB1.DATE_KEY DESC   
"

$dtDates = fnInvokeSQLcmdDSN 'CCL_Prod' $vcDateSQL 0

$vcCurrentMaxLine = $wsTracker.Cells.Item(2,2).text

$i = 0
while($vcCurrentMaxLine -ne $dtDates.cur_date[$i])
{
    $wsTracker.Cells.Item($i + 2,$i + 2).entireRow.insert() | Out-Null

    for($d = 0; $d -lt $dtDates.Columns.Count; $d++)
    {
        $wsTracker.Cells.Item($i + 2, 1) = $dtDates.standard_date[$i]
        $wsTracker.Cells.Item($i + 2, 2) = $dtDates.cur_date[$i]
        $wsTracker.Cells.Item($i + 2, 3) = $dtDates.cur_fiscal_week_of_month[$i]
        $wsTracker.Cells.Item($i + 2, 4) = $dtDates.cur_day_of_week[$i]
    }
    $i++   
}

$iBookLoopDown = 5
$iMaxRunBack = 190

for($iBookLoopDown = 0; $iBookLoopDown -lT $iMaxRunBack; $iBookLoopDown++)
{
    for($i = 0; $i -lt $arrReports.Count; $i++)
    {
        $bRunReport = $false
        $vcReportTargetTime = $false

        if($wsTracker.cells($iBookLoopDown + 2, $arrReports[$i][4]).text -ne 'Complete' -and $wsTracker.cells($iBookLoopDown + 2, $arrReports[$i][4]).text -ne 'ignore')
        {
            if($arrReports[$i][5] -eq 'M' -and 
            $arrReports[$i][3] -eq $dtDates.cur_fiscal_week_of_month[$iBookLoopDown] -and 
            $arrReports[$i][6] -lt $arrReports[$i][1] -and
            $arrReports[$i][2] -eq $dtDates.cur_day_of_week[$iBookLoopDown])
            {
                $bRunReport = $true
                $vcReportTargetTime = $dtDates.year_prev_month[$iBookLoopDown].ToString() + ' ' + $dtDates.month_name_prev_month[$iBookLoopDown]
            }

            if($arrReports[$i][5] -eq 'W' -and 
            $arrReports[$i][6] -lt $arrReports[$i][1] -and
            $arrReports[$i][2] -eq $dtDates.cur_day_of_week[$iBookLoopDown])
            {
                $bRunReport = $true
                $vcReportTargetTime = $dtDates.fiscal_year_prev_week[$iBookLoopDown].ToString() + ' ' + $dtDates.fiscal_month_name_prev_week[$iBookLoopDown] + ' Wk' + $dtDates.fiscal_week_of_month_prev_week[$iBookLoopDown].tostring()
            }

            if($arrReports[$i][5] -eq 'D' -and 
            $arrReports[$i][6] -lt $arrReports[$i][1])
            {
                $bRunReport = $true
                $vcReportTargetTime = $dtDates.prev_date[$iBookLoopDown].ToString()
            }
        
            if($bRunReport -eq $true)
            {
                (get-date).ToString() + ' Running ' + $arrReports[$i][0] + ' for ' + $vcReportTargetTime
                $vcRunResult = &('fn' + $arrReports[$i][0]) $arrReports[$i][0] $vcReportTargetTime $dtDates.Rows.Item($iBookLoopDown) $vcCCL_Prod_Conn
                $arrReports[$i][6] = $arrReports[$i][6] + 1
                $arrReports[$i][0] + ' for ' + $vcReportTargetTime + ' is complete'
                $wsTracker.cells($iBookLoopDown + 2, $arrReports[$i][4]) = 'Complete'
                $wbTracker.Save()
            }
        }
        else
        {
            $arrReports[$i][6] = $arrReports[$i][6] + 1
        }
    }   
}

$wbTracker.Save()
#$wbTracker.Close()
#$objExcel.Quit()

# Test Zone #####################################################################################################################################################
<#
$vcReport = 'Non_Mon'
$vcReport
$arrReportSettings = $arrReports | where {$_ -eq $vcReport}


#0 - Report Name / 1 - Max Runs / 2 - Run Day / 3 - Run Week / 4 - Column / 5 - Weekly/Daily/Monthly / 6 - Current # of Runs

function fnNon_Mon
{
    param
    (
        $Num,
        $Num2
    )
    'Non_M o n' + ' ' + $num + $num2[13]
}

function fnStraight_Rollers
{
    'Straight Rollers'
}

function fnNA_Waterfall
{
    'NA_Waterfall'
}

$arrFunctions2 = $null

$arrFunctions2 += ('Non_Mon', 'Straight_Rollers', 'NA_Waterfall')

&('fn' + $arrFunctions2[0]) '111111' $dtDates.Rows.Item(0) 'try again'
#>