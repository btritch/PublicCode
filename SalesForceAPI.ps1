.'C:\Users\btritch\Desktop\Powershell\SalesForce API\GetSalesForceToken.ps1'
.'C:\Users\btritch\Desktop\Powershell\Utilities\Out-DataTable.ps1'

$userName = ''
$saleForcePW = ''

$tokenAPI = fn_GetSalesForceToken $userName $saleForcePW
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer $tokenAPI")

#Get Objects#######################################################################################
$objectsURL = "https://myg--uat.my.salesforce.com/services/data/v51.0/sobjects/"
$tables = Invoke-RestMethod $objectsURL -Method 'GET' -Headers $headers


for($table = 0; $table -lt $tables.sobjects.count; $table++)
{
    $tableName = $tables.sobjects[$table].name

#Get Fields########################################################################################
    $fieldsURL = "https://myg--uat.my.salesforce.com/services/data/v51.0/sobjects/$tableName/describe/"
    $tableFields = Invoke-RestMethod $fieldsURL -Method 'GET' -Headers $headers

    $tableFieldParam = $null
    if($tableFields.fields.Name.Contains("Id") -ne $true)
    {
        $tableId = $tableFields.fields[0].name
    }
    else
    {
        $tableId = "Id"
    }
    for($field = 0; $field -lt $tableFields.fields.Count; $field++)
    {
        $tableFieldParam += "+"
        if($field -ne 0){$tableFieldParam += ","}
        $tableFieldParam += $tableFields.fields[$field].name
    }

#SOQL###############################################################################################
    if($tableName.Substring($tableName.Length-5, 5) -ne "Event")
    {
        $dataURL= "https://myg--uat.my.salesforce.com/services/data/v51.0/query/?q=SELECT+COUNT($tableId)+FROM+$tableName"
        $dataCount = Invoke-RestMethod  $dataURL -Method 'GET' -Headers $headers
        $dataCount = $dataCount.Records.expr0

        write-host "$table - TableName $tableName contains $dataCount records."

        $resultsTable = $null
        $limitOffset = 0
        while($limitOffset -lt $dataCount)
        {
            $dataURL = "https://myg--uat.my.salesforce.com/services/data/v51.0/query/?q=" +
            "SELECT$tableFieldParam+FROM+$tableName+ORDER+BY+$tableId+LIMIT+500+OFFSET+$limitOffset"
            $dataResults = Invoke-RestMethod  $dataURL -Method 'GET' -Headers $headers
            
            $resultsTable += $dataResults.records | ConvertTo-DataTable

            $limitOffset+=500
        }

        if($dataCount -gt 0)
        {
            $exportPath = "C:\Users\btritch\Desktop\Powershell\SalesForce API\$tableName.csv"
            $resultsTable | export-csv -Path $exportPath -NoTypeInformation
        }
    }
    else
    {
        write-host "$table - TableName $tableName is an event table and is being skipped."
    }

}


#$dataResults.records[0]
