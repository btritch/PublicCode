function fnFDS_Return_Checks
{
    param
    (
        $vcReportName,
        $vcReportDate,
        $dtDatesData,
        $vcConn
    )

    .'Z:\RiskII\Tritch\Operation Loki\Functions.ps1'

    <#
    $vcReportName = 'FDS_Return_Checks'
    $vcReportDate = $vcReportTargetTime
    $dtDatesData = $dtDates.Rows.Item($iBookLoopDown)
    #>

    $vcReportsFolder = 'Z:\RiskII\Risk Analytics\'
    $vcTemplate1 = $vcReportsFolder + 'Report - Templates\' + $vcReportName + ' - Template\' + $vcReportName + ' Template.xlsx'
    $vcOutput1 = $vcReportsFolder + 'Report - Output\' + $vcReportName + ' - Output\' + $vcReportName + ' ' + $vcReportDate + '.xlsx'

    # Workbooks ###################################################################################
    $wbReport1 = $ObjExcel.Workbooks.Open($vcTemplate1, 2, $false) #Open Excel - Workbook Path & Name, Something, Read Only, Something, Password
    
    # Worksheets ##################################################################################
    # Report Views #
    <#$wsReportView1 = $wbReport1.worksheets.item('')
    $wsReportView2 = $wbReport1.worksheets.item('')
    $wsReportView3 = $wbReport1.worksheets.item('')
    $wsReportView4 = $wbReport1.worksheets.item('')
    $wsReportView5 = $wbReport1.worksheets.item('')
    $wsReportView6 = $wbReport1.worksheets.item('')
    $wsReportView7 = $wbReport1.worksheets.item('')
    $wsReportView8 = $wbReport1.worksheets.item('')
    $wsReportView9 = $wbReport1.worksheets.item('')
    $wsReportView0 = $wbReport1.worksheets.item('')#>

    #Data Sheets #
    $wsReportData1 = $wbReport1.worksheets.item('By Acct')
    $wsReportData2 = $wbReport1.worksheets.item('By Tran + Tran Info')
    $wsReportData3 = $wbReport1.worksheets.item('By SSN')
    $wsReportData4 = $wbReport1.worksheets.item('By Acct&SSN + Acct Info')
    <#$wsReportData5 = $wbReport1.worksheets.item('Data5')
    $wsReportData6 = $wbReport1.worksheets.item('Data6')
    $wsReportData7 = $wbReport1.worksheets.item('Data7')
    $wsReportData8 = $wbReport1.worksheets.item('Data8')
    $wsReportData9 = $wbReport1.worksheets.item('Data9')
    $wsReportData0 = $wbReport1.worksheets.item('Data0')#>
    
    # Code Sheets #
    $wsReportCode1 = $wbReport1.worksheets.item('Code1')
    $wsReportCode2 = $wbReport1.worksheets.item('Code2')
    $wsReportCode3 = $wbReport1.worksheets.item('Code3')
    $wsReportCode4 = $wbReport1.worksheets.item('Code4')
    <#$wsReportCode5 = $wbReport1.worksheets.item('Code5')
    $wsReportCode6 = $wbReport1.worksheets.item('Code6')
    $wsReportCode7 = $wbReport1.worksheets.item('Code7')
    $wsReportCode8 = $wbReport1.worksheets.item('Code8')
    $wsReportCode9 = $wbReport1.worksheets.item('Code9')
    $wsReportCode0 = $wbReport1.worksheets.item('Code0')#>
    
    # Various Variables ###########################################################################
    # Date as Date #
    $dtPrevDay = [datetime]::parseexact($dtDatesData[1], 'yyyyMMdd', $null)

    $dtPrevMonth = [datetime]::parseexact($dtDatesData[27], 'yyyyMMdd', $null)
    $dtPrevMonthM1 = [datetime]::parseexact($dtDatesData[28], 'yyyyMMdd', $null)
    $dtPrevMonthM2 = [datetime]::parseexact($dtDatesData[29], 'yyyyMMdd', $null)
    $dtPrevMonthM3 = [datetime]::parseexact($dtDatesData[30], 'yyyyMMdd', $null)
    $dtPrevMonthM4 = [datetime]::parseexact($dtDatesData[31], 'yyyyMMdd', $null)
    $dtPrevMonthM5 = [datetime]::parseexact($dtDatesData[32], 'yyyyMMdd', $null)
    $dtPrevMonthM6 = [datetime]::parseexact($dtDatesData[33], 'yyyyMMdd', $null)
    $dtPrevMonthM7 = [datetime]::parseexact($dtDatesData[34], 'yyyyMMdd', $null)
    $dtPrevMonthM8 = [datetime]::parseexact($dtDatesData[35], 'yyyyMMdd', $null)
    $dtPrevMonthM9 = [datetime]::parseexact($dtDatesData[36], 'yyyyMMdd', $null)
    $dtPrevMonthM10 = [datetime]::parseexact($dtDatesData[37], 'yyyyMMdd', $null)
    $dtPrevMonthM11 = [datetime]::parseexact($dtDatesData[38], 'yyyyMMdd', $null)
    $dtPrevMonthM12 = [datetime]::parseexact($dtDatesData[39], 'yyyyMMdd', $null)
    $dtPrevMonthM13 = [datetime]::parseexact($dtDatesData[40], 'yyyyMMdd', $null)
    $dtPrevMonthM14 = [datetime]::parseexact($dtDatesData[41], 'yyyyMMdd', $null)
    $dtPrevMonthM15 = [datetime]::parseexact($dtDatesData[42], 'yyyyMMdd', $null)
    $dtPrevMonthM16 = [datetime]::parseexact($dtDatesData[43], 'yyyyMMdd', $null)
    $dtPrevMonthM17 = [datetime]::parseexact($dtDatesData[44], 'yyyyMMdd', $null)
    $dtPrevMonthM18 = [datetime]::parseexact($dtDatesData[45], 'yyyyMMdd', $null)
    $dtPrevMonthM19 = [datetime]::parseexact($dtDatesData[46], 'yyyyMMdd', $null)
    $dtPrevMonthM20 = [datetime]::parseexact($dtDatesData[47], 'yyyyMMdd', $null)
    $dtPrevMonthM21 = [datetime]::parseexact($dtDatesData[48], 'yyyyMMdd', $null)
    $dtPrevMonthM22 = [datetime]::parseexact($dtDatesData[49], 'yyyyMMdd', $null)
    $dtPrevMonthM23 = [datetime]::parseexact($dtDatesData[50], 'yyyyMMdd', $null)
    $dtPrevMonthM24 = [datetime]::parseexact($dtDatesData[51], 'yyyyMMdd', $null)

    # Dates as String #
    $vcPrevMonth = $dtDatesData[27]
    $vcPrevMonthM1 = $dtDatesData[28]
    $vcPrevMonthM2 = $dtDatesData[29]
    $vcPrevMonthM3 = $dtDatesData[30]
    $vcPrevMonthM4 = $dtDatesData[31]
    $vcPrevMonthM5 = $dtDatesData[32]
    $vcPrevMonthM6 = $dtDatesData[33]
    $vcPrevMonthM7 = $dtDatesData[34]
    $vcPrevMonthM8 = $dtDatesData[35]
    $vcPrevMonthM9 = $dtDatesData[36]
    $vcPrevMonthM10 = $dtDatesData[37]
    $vcPrevMonthM11 = $dtDatesData[38]
    $vcPrevMonthM12 = $dtDatesData[39]
    $vcPrevMonthM13 = $dtDatesData[40]
    $vcPrevMonthM14 = $dtDatesData[41]
    $vcPrevMonthM15 = $dtDatesData[42]
    $vcPrevMonthM16 = $dtDatesData[43]
    $vcPrevMonthM17 = $dtDatesData[44]
    $vcPrevMonthM18 = $dtDatesData[45]
    $vcPrevMonthM19 = $dtDatesData[46]
    $vcPrevMonthM20 = $dtDatesData[47]
    $vcPrevMonthM21 = $dtDatesData[48]
    $vcPrevMonthM22 = $dtDatesData[49]
    $vcPrevMonthM23 = $dtDatesData[50]
    $vcPrevMonthM24 = $dtDatesData[51]
    
    # Worksheet Formats ###########################################################################

    $iPayAmt = '5000'
    [int] $iDaysBack = 45

    $vcStartDate = $dtPrevDay.adddays(-($iDaysBack + 2)).tostring('yyyyMMdd')
    $vcEndDate = $dtPrevDay.adddays(-2).tostring('yyyyMMdd')

    $vcStartDate2 = $dtPrevDay.adddays(-($iDaysBack * 2 + 2)).tostring('yyyyMMdd')
    $vcEndDate2 = $dtPrevDay.adddays(-1).tostring('yyyyMMdd')
    ###############################################################################################

    $vcSQLCode1 = 
"
    SELECT
        *
    FROM
    (
        SELECT
            CONCAT('L', TRANS.CUST_ID) AS CUST_ID
            ,COUNT(*) AS NUM_TRANS
            ,SUM(TRANS.TRANSACTION_AMOUNT) AS SUM_TRANSACTION_AMOUNT
        FROM
            CRD_TRANS.PDW_TRANSACTION AS TRANS

        LEFT JOIN CRD_MSTR.PDW_DMAST AS MMAST
        ON MMAST.CUST_ID = TRANS.CUST_ID
        AND MMAST.PART_DATE = '$vcEndDate2'

        WHERE TRANS.TRANSACTION_DATE BETWEEN '$vcStartDate' AND '$vcEndDate' 
        AND TRANS.TRANSACTION_CODE = '271'
        AND COALESCE(MMAST.CUST_FLG_3,'') = 'F'
        AND RTRN_CHECK_DT BETWEEN '$vcStartDate2' AND '$vcEndDate' 

        GROUP BY
            CONCAT('L', TRANS.CUST_ID) 
    ) AS SUB1
    WHERE SUB1.SUM_TRANSACTION_AMOUNT >= $iPayAmt
"

    $vcSQLCode2 = 
"
    SELECT
            CONCAT('L', SUB1.CUST_ID) AS CUST_ID
            ,SUB1.NUM_TRANS
            ,SUB1.SUM_TRANSACTION_AMOUNT
            ,TRANS2.TRANSACTION_DATE
            ,DATE(DAYS(CONCAT(CAST(INTEGER(2000000 + TRANS2.JULIAN_POST_DATE) / 1000 AS CHAR(4)),'-01-01')) + MOD(INTEGER(2000000 + TRANS2.JULIAN_POST_DATE),1000)-1) AS POST_DATE
            ,TRANS2.TRANSACTION_AMOUNT
            ,CONCAT('L', TRANS2.MRCH_ACCOUNT_NUMBER) AS MRCH_ACCOUNT_NUMBER
        FROM
        (
            SELECT
                TRANS.CUST_ID
                ,COUNT(*) AS NUM_TRANS
                ,SUM(TRANS.TRANSACTION_AMOUNT) AS SUM_TRANSACTION_AMOUNT
            FROM
                CRD_TRANS.PDW_TRANSACTION AS TRANS
        
            LEFT JOIN CRD_MSTR.PDW_DMAST AS MMAST
            ON MMAST.CUST_ID = TRANS.CUST_ID
            AND MMAST.PART_DATE = '$vcEndDate2'
        
            WHERE TRANS.TRANSACTION_DATE BETWEEN '$vcStartDate' AND '$vcEndDate'
            AND TRANS.TRANSACTION_CODE = '271'

            AND COALESCE(MMAST.CUST_FLG_3,'') = 'F'
            AND RTRN_CHECK_DT BETWEEN '$vcStartDate2' AND '$vcEndDate' 
        
            GROUP BY
                TRANS.CUST_ID
        ) AS SUB1
        
        LEFT JOIN CRD_TRANS.PDW_TRANSACTION AS TRANS2
        ON TRANS2.CUST_ID = SUB1.CUST_ID
        AND TRANS2.TRANSACTION_DATE BETWEEN '$vcStartDate' AND '$vcEndDate'
        AND TRANS2.TRANSACTION_CODE = '271'
         
        LEFT JOIN CRD_MSTR.PDW_DMAST AS MMAST2 
        ON MMAST2.CUST_ID = TRANS2.CUST_ID 
        AND MMAST2.PART_DATE = '$vcEndDate2'
        AND COALESCE(MMAST2.CUST_FLG_3,'') = 'F'
        AND RTRN_CHECK_DT BETWEEN '$vcStartDate2' AND '$vcEndDate' 

        WHERE SUB1.SUM_TRANSACTION_AMOUNT >= $iPayAmt 
"

    $vcSQLCode3 = 
"
    SELECT
        *
    FROM
    (
        SELECT
            MMAST1.PRI_SSN_NR
            ,COUNT(*) AS NUM_TRANS
            ,SUM(TRANS.TRANSACTION_AMOUNT) AS SUM_TRANSACTION_AMOUNT
        FROM
            CRD_TRANS.PDW_TRANSACTION AS TRANS
        
        LEFT JOIN CRD_MSTR.PDW_DMAST AS MMAST
        ON MMAST.CUST_ID = TRANS.CUST_ID
        AND MMAST.PART_DATE = '$vcEndDate2'

        LEFT JOIN CRD_MSTR.PDW_MMAST_EXTRACT_TEXT AS MMAST1
        ON MMAST1.CUST_ID = TRANS.CUST_ID
        AND MMAST1.PART_DATE = '$vcPrevMonth'
        
        WHERE TRANS.TRANSACTION_DATE BETWEEN '$vcStartDate' AND '$vcEndDate'
        AND TRANS.TRANSACTION_CODE = '271'
        AND COALESCE(MMAST.CUST_FLG_3,'') = 'F'
        AND MMAST.RTRN_CHECK_DT BETWEEN '$vcStartDate2' AND '$vcEndDate' 
        AND PRI_SSN_NR IS NOT NULL
        
        GROUP BY
            MMAST1.PRI_SSN_NR
    ) AS SUB1

    WHERE SUB1.SUM_TRANSACTION_AMOUNT >= $iPayAmt
"

    $vcSQLCode4 = 
"
    SELECT
        SUB2.CUST_ID
        ,SUB2.PRI_SSN_NR
        ,SUB2.PRI_CUST_NM
        ,SUB2.SEC_SSN_NR
        ,SUB2.SEC_CUST_NM
        ,SUB2.ADDR_LINE_1
        ,SUB2.ADDR_LINE_2
        ,SUB2.CITY
        ,SUB2.STATE
        ,SUB2.ZIP_CODE
        ,SUB2.PRI_TELE_NBR
        ,SUB2.SEC_TELE_NBR
        ,SUB1.*
        ,SUB2.TRANSACTION_DATE
        ,SUB2.TRANSACTION_AMOUNT
        ,SUB2.POST_DATE
    FROM
    (
        SELECT
            MMAST1.PRI_SSN_NR
            ,COUNT(*) AS NUM_TRANS
            ,SUM(TRANS.TRANSACTION_AMOUNT) AS SUM_TRANSACTION_AMOUNT
        FROM
            CRD_TRANS.PDW_TRANSACTION AS TRANS
        
        LEFT JOIN CRD_MSTR.PDW_DMAST AS MMAST
        ON MMAST.CUST_ID = TRANS.CUST_ID
        AND MMAST.PART_DATE = '$vcEndDate2'

        LEFT JOIN CRD_MSTR.PDW_MMAST_EXTRACT_TEXT AS MMAST1
        ON MMAST1.CUST_ID = TRANS.CUST_ID
        AND MMAST1.PART_DATE = '$vcPrevMonth'
        
        WHERE TRANS.TRANSACTION_DATE BETWEEN '$vcStartDate' AND '$vcEndDate'
        AND TRANS.TRANSACTION_CODE = '271'
        AND COALESCE(MMAST.CUST_FLG_3,'') = 'F'
        AND MMAST.RTRN_CHECK_DT BETWEEN '$vcStartDate2' AND '$vcEndDate' 
        AND PRI_SSN_NR IS NOT NULL

        GROUP BY
            MMAST1.PRI_SSN_NR
    ) AS SUB1

    LEFT JOIN 
    (
        SELECT
            MMAST.CUST_ID
            ,MMAST1.PRI_SSN_NR
            ,MMAST1.PRI_CUST_NM
            ,MMAST1.SEC_SSN_NR
            ,MMAST1.SEC_CUST_NM
            ,MMAST1.ADDR_LINE_1
            ,MMAST1.ADDR_LINE_2
            ,MMAST1.CITY
            ,MMAST1.STATE
            ,MMAST1.ZIP_CODE
            ,MMAST1.PRI_TELE_NBR
            ,MMAST1.SEC_TELE_NBR
            ,TRANS.TRANSACTION_DATE
            ,TRANS.TRANSACTION_AMOUNT
            ,DATE(DAYS(CONCAT(CAST(INTEGER(2000000 + TRANS.JULIAN_POST_DATE) / 1000 AS CHAR(4)),'-01-01')) + MOD(INTEGER(2000000 + TRANS.JULIAN_POST_DATE),1000)-1) AS POST_DATE
        FROM
            CRD_TRANS.PDW_TRANSACTION AS TRANS
        
        LEFT JOIN CRD_MSTR.PDW_DMAST AS MMAST
        ON MMAST.CUST_ID = TRANS.CUST_ID
        AND MMAST.PART_DATE = '$vcEndDate2'

        LEFT JOIN CRD_MSTR.PDW_MMAST_EXTRACT_TEXT AS MMAST1
        ON MMAST1.CUST_ID = TRANS.CUST_ID
        AND MMAST1.PART_DATE = '$vcPrevMonth'
        
        WHERE TRANS.TRANSACTION_DATE BETWEEN '$vcStartDate' AND '$vcEndDate'
        AND TRANS.TRANSACTION_CODE = '271'
        AND COALESCE(MMAST.CUST_FLG_3,'') = 'F'
        AND MMAST.RTRN_CHECK_DT BETWEEN '$vcStartDate2' AND '$vcEndDate' 
        AND PRI_SSN_NR IS NOT NULL
    ) AS SUB2
    ON SUB2.PRI_SSN_NR = SUB1.PRI_SSN_NR

    WHERE SUB1.SUM_TRANSACTION_AMOUNT >= $iPayAmt
"

    # NEW CONNECTION CODE #########################################################################

    (get-date).ToString() + ' ' + $vcReportName + ' Data1'
    $vcSQLCode1_1 = $vcSQLCode1.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
    $wbReport1.Connections.Item('Data1').ODBCConnection.CommandText = $vcSQLCode1_1
    $wbReport1.Connections.Item('Data1').ODBCConnection.Connection = $vcConn
    $wbReport1.Connections.Item('Data1').ODBCConnection.BackgroundQuery = $false
    $wbReport1.Connections.Item('Data1').refresh()

    (get-date).ToString() + ' ' + $vcReportName + ' Data2'
    $vcSQLCode2_1 = $vcSQLCode2.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
    $wbReport1.Connections.Item('Data2').ODBCConnection.CommandText = $vcSQLCode2_1
    $wbReport1.Connections.Item('Data2').ODBCConnection.Connection = $vcConn
    $wbReport1.Connections.Item('Data2').ODBCConnection.BackgroundQuery = $false
    $wbReport1.Connections.Item('Data2').refresh()

    (get-date).ToString() + ' ' + $vcReportName + ' Data3'
    $vcSQLCode3_1 = $vcSQLCode3.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
    $wbReport1.Connections.Item('Data3').ODBCConnection.CommandText = $vcSQLCode3_1
    $wbReport1.Connections.Item('Data3').ODBCConnection.Connection = $vcConn
    $wbReport1.Connections.Item('Data3').ODBCConnection.BackgroundQuery = $false
    $wbReport1.Connections.Item('Data3').refresh()

    (get-date).ToString() + ' ' + $vcReportName + ' Data4'
    $vcSQLCode4_1 = $vcSQLCode4.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
    $wbReport1.Connections.Item('Data4').ODBCConnection.CommandText = $vcSQLCode4_1
    $wbReport1.Connections.Item('Data4').ODBCConnection.Connection = $vcConn
    $wbReport1.Connections.Item('Data4').ODBCConnection.BackgroundQuery = $false
    $wbReport1.Connections.Item('Data4').refresh()

    # CODE OUTPUT TO WORKBOOK #####################################################################
    $wsReportCode1.Columns('A').NumberFormat = "@"
    $wsReportCode1.Columns('A').columnwidth = 150
    $wsReportCode1.Columns('A').Clear() | Out-Null
    $arrSQLCode1 = $vcSQLCode1 -split "`r`n"
    for($c = 0; $c -lt $arrSQLCode1.Count; $c++)
    {
        $wsReportCode1.Cells.Item(($c+1), 1) = $arrSQLCode1[$c]
    }

    $wsReportCode2.Columns('A').NumberFormat = "@"
    $wsReportCode2.Columns('A').columnwidth = 150
    $wsReportCode2.Columns('A').Clear() | Out-Null
    $arrSQLCode2 = $vcSQLCode2 -split "`r`n"
    for($c = 0; $c -lt $arrSQLCode2.Count; $c++)
    {
        $wsReportCode2.Cells.Item(($c+1), 1) = $arrSQLCode2[$c]
    }
    
    $wsReportCode3.Columns('A').NumberFormat = "@"
    $wsReportCode3.Columns('A').columnwidth = 150
    $wsReportCode3.Columns('A').Clear() | Out-Null
    $arrSQLCode3 = $vcSQLCode3 -split "`r`n"
    for($c = 0; $c -lt $arrSQLCode3.Count; $c++)
    {
        $wsReportCode3.Cells.Item(($c+1), 1) = $arrSQLCode3[$c]
    }

    $wsReportCode4.Columns('A').NumberFormat = "@"
    $wsReportCode4.Columns('A').columnwidth = 150
    $wsReportCode4.Columns('A').Clear() | Out-Null
    $arrSQLCode4 = $vcSQLCode4 -split "`r`n"
    for($c = 0; $c -lt $arrSQLCode4.Count; $c++)
    {
        $wsReportCode4.Cells.Item(($c+1), 1) = $arrSQLCode4[$c]
    }
    
    ###############################################################################################
    $wbReport1.SaveAs($vcOutput1)
    $wbReport1.Close()

    return 'Complete'

}
# TEST ZONE ###############################################################################################################
<#
$vcSQLCode1_1 = $vcSQLCode1.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
$wbReport1.Connections.Item('Data1').ODBCConnection.CommandText = $vcSQLCode1_1
$wbReport1.Connections.Item('Data1').ODBCConnection.BackgroundQuery = $false
$wbReport1.Connections.Item('Data1').refresh()

$vcSQLCode2_1 = $vcSQLCode2.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
$wbReport1.Connections.Item('Data2').ODBCConnection.CommandText = $vcSQLCode2_1
$wbReport1.Connections.Item('Data2').ODBCConnection.BackgroundQuery = $false
$wbReport1.Connections.Item('Data2').refresh()

$vcSQLCode3_1 = $vcSQLCode3.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
$wbReport1.Connections.Item('Data3').ODBCConnection.CommandText = $vcSQLCode3_1
$wbReport1.Connections.Item('Data3').ODBCConnection.BackgroundQuery = $false
$wbReport1.Connections.Item('Data3').refresh()

$vcSQLCode4_1 = $vcSQLCode4.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")
$wbReport1.Connections.Item('Data4').ODBCConnection.CommandText = $vcSQLCode4_1
$wbReport1.Connections.Item('Data4').ODBCConnection.BackgroundQuery = $false
$wbReport1.Connections.Item('Data4').refresh()

$wsReportData1.ListObjects('Table1')

$test.ODBCConnection. = $wsReportData1.ListObjects.Item('Table1')

$test2 = $test.QueryTables.Add("ODBC;DSN=$vcDSN", $test.Range('A1'), $vcSQLCode1) 
$test2.refresh()

$Data1 = fnInvokeSQLcmdDSN 'CCL_Prod' $vcSQLCode1 0

$test4 =
"
    SELECT
        *
    FROM REFERENCEDATA.DBO_DATE
    WHERE FISCAL_YEAR = 2018
"

$test = $wbReport1.Connections.Item('Data1').refresh()

$test2 = $test$wbReport1.Connections.Item('Data1').ODBCConnection

$test2.RefreshPeriod

$wbReport1.Conne

$start = get-date
for($i = 0; $i -lt $Data1.Rows.Count; $i++)
{
    for($l = 0; $l -lt $Data1.Columns.Count; $l++)
    {
        $wsReportData1.Cells.Item($i+2, $l+1) = $Data1.Rows[$i][$l]
    }
}
$end = get-date


$i

$vcSQLCode1_1 = $vcSQLCode1.Replace("`n", " ")

$vcSQLCode1_1 = $vcSQLCode1.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ")

#>