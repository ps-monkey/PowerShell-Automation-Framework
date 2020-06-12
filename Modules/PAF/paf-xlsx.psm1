#Define PAFXLSX object type
class PAFXLSX {
[ValidateNotNullOrEmpty()][string] $Name
[HashTable] $Properties
[HashTable[]] $CustomCells
[HashTable[]] $CustomRows
[HashTable[]] $CustomColumns
[PSCustomObject] $Data
}
<#
$Cell1 = @{'CellAddress' = "A1"; 'Value' = "sasadsasd"; 'FontFamily' = "Arial"; 'FontSize' = 20; 'FontColor' = "#00FF00"; 'FontStyle' = "bold italic"; 'CellColor' = "#FF0000"; 'WrapText' = $true; }
$Cell2 = @{'CellAddress' = "A2"; 'Value' = "00000"; 'FontFamily' = "Verdana"; 'FontSize' = 10; 'FontColor' = "#0000FF"; 'FontStyle' = "bold";}
$CustomCells = @($Cell1,$Cell2)

$column1 = @{Index = 1; Width = 5}
$column2 = @{Index = 2; Width = 5}
CustomColumns = @($column1,$column2)

$Properties = @{'Autofit' = $true; 'AutoFilter' = $true; 'Bold-Header'= $true; 'Set-FreezeRow' = 2; 'Set-FreezeColumn' = 3 }

$Data = PAF-New-XLSX-Object -Name "test" -CustomCells $CustomCells -Properties $Properties -Data $VMReport
PAF-XLSX-Create -XLSXObject $Data -File $file
#>

Function PAF-New-XLSX-Object {
	param (
		[Parameter(Mandatory=$true)][string] $Name,
		[Parameter(Mandatory=$false)][HashTable] $Properties,
		[Parameter(Mandatory=$false)][HashTable[]] $CustomCells,
		[Parameter(Mandatory=$false)][HashTable[]] $CustomRows,
		[Parameter(Mandatory=$false)][HashTable[]] $CustomColumns,
		[Parameter(Mandatory=$true)][PSCustomObject] $Data
		)	
Return [PAFXLSX][Ordered]@{ Name = $Name; CustomCells = $CustomCells; Properties = $Properties; CustomRows = $CustomRows; CustomColumns = $CustomColumns; Data = $Data; }
}

Function PAF-XLSX-Create {
	param (
	[Parameter(Mandatory=$true)][PAFXLSX[]] $XLSXObject,
	[Parameter(Mandatory=$true)][string] $File
	)
If (Test-Path $File) { Remove-Item -Path $File -Force }

#Export XLSX
ForEach ($Worksheet in $XLSXObject) {
	If ($Worksheet.Data) {
		If ($Worksheet.Properties.ClearSheet) { $Worksheet.Data | Export-XLSX -Path $File -Worksheet $Worksheet.Name -ClearSheet }
		Else { $Worksheet.Data | Export-XLSX -Path $File -Worksheet $Worksheet.Name }
		}
	}
#Edit cells
$Excel = New-Excel -Path $File
ForEach ($WS in $XLSXObject) {
	$Worksheet = $Excel | Get-WorkSheet | ? { $_.name -eq $WS.Name }
	$Params = @{}
	If ($WS.Properties.AutoFilter) { $Params.Add("AutoFilter", $true) }
	If ($WS.Properties.'Bold-Header') { $Params.Add("Header", $true); $Params.Add("Bold", $true) }
	If ($WS.Properties.'Set-FreezeRow') { $WorkSheet | Set-FreezePane -Row $WS.Properties.'Set-FreezeRow' }
	If ($WS.Properties.'Set-FreezeColumn') { $WorkSheet | Set-FreezePane -Column $WS.Properties.'Set-FreezeColumn' }
	$WorkSheet | Format-Cell @Params
	
	ForEach ($CustomCell in $WS.CustomCells) { 	PAF-XLSX-EditCell -CustomCell $CustomCell -WorkSheet $WorkSheet }

	If ($WS.Properties.Autofit) { $WorkSheet | Format-Cell -Autofit }
	
	ForEach ($CustomRow in $WS.CustomRows) { PAF-XLSX-EditRow -CustomRows $CustomRows -WorkSheet $WorkSheet }
	ForEach ($CustomColumn in $WS.CustomColumns) { PAF-XLSX-EditColumn -CustomColumn $CustomColumn -WorkSheet $WorkSheet }
	}
$Excel | Close-Excel -Save
}

Function PAF-XLSX-EditCell {
	param (
	[Parameter(Mandatory=$true)][PSCustomObject] $CustomCell,
	[Parameter(Mandatory=$true)] $WorkSheet
	)

$Cell = $WorkSheet.Cells | ? {$_.Address -eq $CustomCell.CellAddress}

If ($CustomCell.Value) { $Cell.Value = $CustomCell.Value }
If ($CustomCell.Formula) { $Cell.Formula = $CustomCell.Formula }
If ($CustomCell.FontFamily) { $Cell.Style.Font.Name = $CustomCell.FontFamily }
If ($CustomCell.FontSize) { $Cell.Style.Font.Size = $CustomCell.FontSize }
If ($CustomCell.FontColor) { $Cell.Style.Font.Color.SetColor($CustomCell.FontColor) }
If ($CustomCell.FontStyle) {
	Switch -regex ($CustomCell.FontStyle) {
		"bold" { $Cell.Style.Font.Bold = $true }
		"italic" {$Cell.Style.Font.Italic = $true }
		"strike" { $Cell.Style.Font.Strike = $true }
		"underline" { $Cell.Style.Font.UnderLine = $true }
		}
	}
If ($CustomCell.CellColor) { 
	$Cell.Style.Fill.PatternType = "Solid"
	$Cell.Style.Fill.BackgroundColor.SetColor($CustomCell.CellColor)
	}
If ($CustomCell.WrapText) { $Cell.Style.WrapText = $true }
If ($CustomCell.TextRotation) { $Cell.Style.TextRotation = $CustomCell.TextRotation }


#VerticalAlignment
#HorizontalAlignment
}

Function PAF-XLSX-EditColumn {
	param (
	[Parameter(Mandatory=$true)][PSCustomObject] $CustomColumn,
	[Parameter(Mandatory=$true)] $WorkSheet
	)

$Column = $WorkSheet.Column($CustomColumn.Index)
If ($CustomColumn.Width) { $Column.Width = $CustomColumn.Width }
If ($CustomColumn.Collapsed) { $CustomColumn.Collapsed = $true }
If ($CustomColumn.Hidden) { $CustomColumn.Hidden = $true }
}

Function PAF-XLSX-EditRow {
	param (
	[Parameter(Mandatory=$true)][PSCustomObject] $CustomRow,
	[Parameter(Mandatory=$true)] $WorkSheet
	)

$Row = $WorkSheet.Column($CustomRow.Index)
If ($CustomRow.Width) { $Row.Width = $CustomRow.Width }
If ($CustomRow.Collapsed) { $Row.Collapsed = $true }
If ($CustomRow.Hidden) { $Row.Hidden = $true }
}
