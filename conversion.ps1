clear-host

$csv = "C:\Users\ortiga\Desktop\table.csv" #archivo origen
$xlsx = "C:\Users\ortiga\Desktop\table2.xlsx" #archivo destino
$delimitador = "," #especificamos el delimitador

# creamos una nueva hoja vacia y la seleccionamos
$excel = New-Object -ComObject excel.application 
$documento = $excel.Workbooks.Add(1)
$hoja = $documento.worksheets.Item(1)




$conecta = ("TEXT;" + $csv)
$Conector = $hoja.QueryTables.add($conecta,$hoja.Range("A1"))
$query = $hoja.QueryTables.item($Conector.name)
$query.TextFileOtherDelimiter = $delimitador
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,1 * $hoja.Cells.Columns.Count
$query.AdjustColumnWidth = 1


$query.Refresh()
$query.Delete()

#guardamos el documento

$documento.SaveAs($xlsx,51)
$excel.Quit()

Get-Process -Name Excel | Stop-Process