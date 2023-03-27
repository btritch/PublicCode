
.'Z:\RiskII\Tritch\Operation Loki\fnSendToTop.ps1'

############################################################
# CHANGE THESE TO RUN THE THING THAT YOU WANT TO RUN... FOOL
[string] $vcMonth = 'September' #input the month to be ran. full month name
[string] $vcYear = '2019' #input year in the yyyy format
############################################################

$ReportList = 'Z:\RiskII\Risk Analytics\Report Package - Report List 3.xlsx'

$Excel = New-Object -ComObject "Excel.Application" 
$Excel.Visible = $true #Runs Excel in the background. 
$Excel.DisplayAlerts = $false #Supress alert messages. 

################################################################
# WORKBOOK AND WORKSHEET VARIABLES
$wbCompile = $Excel.Workbooks.open($ReportList) 

$wsCover = $wbCompile.worksheets.item('Cover')
$wsReportList = $wbCompile.worksheets.item('Report List')
$wsLast = $wbCompile.worksheets.item('The End')
################################################################

##################################
# VARIABLE INITIATION
[int] $vcCurFileWeek = 0

# THE VARIABLES NEEDED FOR TABLE OF CONTENTS
#####################################################################################################################################################
$iStartPage = 0
$iEndPage = 0
$iTableContents = 2

$iMaxRows = 37 #this is the number of rows for the table of contents. it is 1 more than the max number of items, the +1 is for the header
$iPageCount = 1 #set initial value

$iTargetRow = $iMaxRows + 1 #row target
$iPageSideInd = 0 #0 is left 4 is right
$iPageSideIndLegacy = 0
#####################################################################################################################################################

##################################

$wsCover.Range("A4") = "'" + $vcMonth + ' ' + $vcYear

$iReportLoop = 2
while($wsReportList.Cells($iReportLoop, 1).text -gt 0) #loop through all reports to be added to report
{
    (get-date).ToString() + ' Running ' + $wsReportList.Cells($iReportLoop, 3).text #text to output to console to show progress

    $vcFileToUseName = $null
    $vcFilePath = $wsReportList.Cells($iReportLoop, 13).text #get file path from workbook
    $afiles = Get-ChildItem $vcFilePath -File #add all items in folder to array

    for($i = 0; $i -lt $afiles.Length; $i++) #loop through all files in folder to find correct file to use
    {
        $vcFolderName = $vcFilePath.Substring(41, $vcFilePath.Length -42) #folder name without full path (path length is 42)

        $vcTimeFrame = $afiles[$i].Name.substring($vcFolderName.substring(0, $vcFolderName.length - 10).length + 2,
         $afiles[$i].Name.length - $vcFolderName.substring(0, $vcFolderName.length - 10).length - 2 - 5) #code to get year/month/week

        $vcFileYear = $vcTimeFrame.Substring(0,4) #the year of the file

        if($vcTimeFrame.Substring($vcTimeFrame.Length-3, 2) -eq 'Wk') #determine if the report is a weekly or monthly report
        { #weekly file
            $vcFileMonth = $vcTimeFrame.Substring(5,$vcTimeFrame.Length-9) #file month
            $vcFileWeek = $vcTimeFrame.Substring($vcTimeFrame.Length-1, 1) #file week
        }
        else 
        { #monthly file
            $vcFileMonth = $vcTimeFrame.Substring(5, $vcTimeFrame.Length-5) #file month
            $vcFileWeek = 0 #file week set to 0 for monthly report
        }

        if($vcFileYear -eq $vcYear -and $vcFileMonth -eq $vcMonth) #does year and month match?
        {
            if($vcFileWeek -ge $vcCurFileWeek) #compare file week # to current max week #
            {
                $vcCurFileWeek = $vcFileWeek # set current week to value of the file just compared
                $vcFileToUseName = $afiles[$i].Name #name of file to be used without path
                $vcFilesToUse = $afiles[$i].FullName #full name and file path of file to be used
            }
        }
    }
    $wsReportList.cells.item($iReportLoop, 12) = $vcFileToUseName #put the name of the file used into the report

    #################################################################################################################################################
    #OPEN REPORT WORKBOOK AND DO THE THINGS
    
    if($vcFileToUseName -ne $null)
    {
        $wbReport = $Excel.Workbooks.open($vcFilesToUse) 

        $wsLast.Cells.Item($iTableContents,(1 + $iPageSideInd)) = $wsReportList.Cells($iReportLoop, 3).text 
        $wsLast.Cells.Item($iTableContents,(1 + $iPageSideInd)).font.Bold = $true
        $wsLast.Cells.Item($iTableContents,(3 + $iPageSideInd)).font.Bold = $true
        $iReportPages = $iTableContents
        $iStartPage = $iEndPage + 1
    
        $iPageSideIndLegacy = $iPageSideInd

        #fnReturnToTop
        if($iTableContents -gt $iTargetRow)
        {
            if($iPageSideInd -eq 0)
            {
                $iPageSideInd = 4
                $iTableContents = $iTargetRow - $iMaxRows + 1
            }
            else
            {
                $iPageSideInd = 0
                $iPageCount++
                $iTargetRow = $iTargetRow + $iMaxRows
            }
        }
        else
        {
            $iTableContents++
        }

        for($i = 14; $i -lt 53)
        {
            if($wsReportList.cells($iReportLoop, $i).text -ne "")
            {
                $wsMoveMe = $wbReport.worksheets.item($wsReportList.cells($iReportLoop, $i).text) 
                $wsMoveMe.Cells().copy() | Out-Null
                $wsMoveMe.Range('A1').PasteSpecial(-4163) | Out-Null
                $wsMoveMe.name = $wsReportList.cells($iReportLoop, $i + 1).text
                $wsMoveMe.copy($wsLast)

                $wsReport = $wbCompile.Sheets($wbCompile.sheets.Count-1)

                #Header = Workbook Name Footer = [Center] Worksheet Name [Right] Page Page#
                $wsReport.PageSetup.CenterHeader = '&F'
                $wsReport.PageSetup.CenterFooter = '&A'
                $wsReport.PageSetup.RightFooter = 'Page &P'

                #Print Margins
                $wsReport.PageSetup.LeftMargin = $Excel.Application.InchesToPoints(0.25)
                $wsReport.PageSetup.RightMargin = $Excel.Application.InchesToPoints(0.25)
                $wsReport.PageSetup.TopMargin = $Excel.Application.InchesToPoints(0.4)
                $wsReport.PageSetup.BottomMargin = $Excel.Application.InchesToPoints(0.4)
                $wsReport.PageSetup.HeaderMargin = $Excel.Application.InchesToPoints(0.2)
                $wsReport.PageSetup.FooterMargin = $Excel.Application.InchesToPoints(0.2)

                $wsReport.PageSetup.FitToPagesWide = 1
                $wsReport.PageSetup.FitToPagesTall = 1

                $wsReport.PageSetup.Orientation = 2

                $wsReport.PageSetup.CenterHorizontally = $true

                $wsLast.Cells.Item($iTableContents,(2 + $iPageSideInd)) = $wsReportList.cells($iReportLoop, $i + 1).text
                $iEndPage++
                $wsLast.Cells.Item($iTableContents,(3 + $iPageSideInd)) = $iEndPage
            
                #fnReturnToTop
                if($iTableContents -eq $iTargetRow)
                {
                    if($iPageSideInd -eq 0)
                    {
                        $iPageSideInd = 4
                        $iTableContents = $iTargetRow - $iMaxRows + 1
                    }
                    else
                    {
                        $iPageSideInd = 0
                        $iPageCount++
                        $iTargetRow = $iTargetRow + $iMaxRows
                        $iTableContents++
                    }
                }
                else
                {
                    $iTableContents++
                }
            
                $wsMoveMe = $null
            }
            $i = $i + 2
        }

        if($iStartPage -eq $iEndPage)
        {
            $wsLast.Cells.Item($iReportPages,(3 + $iPageSideIndLegacy)) = $iStartPage.ToString() 
        }
        else
        {
            $wsLast.Cells.Item($iReportPages,(3 + $iPageSideIndLegacy)) = $iStartPage.ToString() + ' - ' + $iEndPage.ToString()
        }
        $iPageSideIndLegacy = $iPageSideInd

        $wbReport.Close()
        #################################################################################################################################################

        ############################
        # SET SHIT TO NULL OR 0
        $vcFileYear = $null
        $vcFileMonth = $null
        $vcFileWeek = $null
        $vcFileToUseName = $null 

        $vcCurFileWeek = 0
        ############################
    }
    $iReportLoop++
}

[string] $varLastUpdate = (get-date).ToString('yyyy.MM.dd')

$wsCover.Cells.Item(5,1) = 'Updated on ' + (get-date).ToString('M/d/yyyy')
$wsLast.move($wsReportList)
$wsReportList.Move($wsLast)

$wbCompile.SaveAs('Z:\RiskII\Risk Analytics\Report Package\Report Package ' + $vcYear + ' ' + $vcMonth + ' - Last Updated on ' + $varlastUpdate + '.xlsx')
