cls

$query = "SELECT TOP (2000) [ProductKey]
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

Write-Host "Connection Open"

$connection.Open()

$cmd = New-object System.Data.Odbc.OdbcCommand($query,$connection)

Write-Host "Reading"

$reader = $cmd.ExecuteReader()

$rows = 0

while($reader.Read())
{  
    $rows++

    if($rows % 500 -eq 0)
    {
        Write-Host "Read $rows rows"

        $fieldCount=$reader.FieldCount

        Write-Host "Sample record:"
        
        for ($fieldOrdinal = 0; $fieldOrdinal -lt $fieldCount; $fieldOrdinal++)
        {
            $colName=$reader.GetName($fieldOrdinal)
            
            Write-Host "Column: $colName; Value: $($reader[$fieldOrdinal])"
        }
    }
}  

 Write-Host "Total rows: $rows "

$reader.Dispose()

$cmd.Dispose()

$connection.Dispose()