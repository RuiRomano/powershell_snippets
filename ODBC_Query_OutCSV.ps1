cls

$currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)

$outputFile = "$currentPath\output.csv"

$query = "SELECT TOP (1000) [ProductKey]
      ,[Product Code]
      ,[Product Name]
      ,[Manufacturer]
      ,[Brand]
      ,[Color]
      ,[Weight Unit Measure]
      ,[Weight]
      ,[Unit Cost]
      ,[Unit Price]
      ,[Subcategory Code]
      ,[Subcategory]
      ,[Category Code]
      ,[Category]
  FROM [Contoso 1M].[Data].[Product]"

$connectionString = "DSN=CONTOSO_ODBC;Uid=;Pwd=;"

$connection = new-object System.Data.Odbc.OdbcConnection
        
$connection.ConnectionString = $connectionString

$connection.Open()

$cmd = New-object System.Data.Odbc.OdbcCommand($query,$connection)

$dataset = New-Object System.Data.DataSet

(New-Object System.Data.Odbc.OdbcDataAdapter($cmd)).Fill($dataSet) | Out-Null

$dataset.Tables[0] | ConvertTo-Csv | Out-File $outputFile -Force

$cmd.Dispose()

$connection.Dispose()