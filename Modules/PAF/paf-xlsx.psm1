$BinaryPath = Join-Path $PSScriptRoot 'lib\epplus.dll'
If( -not ($Library = Add-Type -path $BinaryPath -PassThru -ErrorAction stop) ) { Throw "Failed to load EPPlus binary from $BinaryPath" }

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

#Functions below are taken from  PSExcel module https://github.com/RamblingCookieMonster/PSExcel
Function Export-XLSX {
[CmdletBinding(DefaultParameterSetName='Path')]
param(
	[parameter(ParameterSetName='Path', Position = 0, Mandatory=$true )][ValidateScript({$Parent = Split-Path $_ -Parent; If ( -not (Test-Path -Path $Parent -PathType Container) ) { Throw "SpecIfy a valid path.  Parent '$Parent' does not exist: $_" }; $True })] [string]$Path,
	[parameter(ParameterSetName='Excel', Position = 0, Mandatory=$true )][OfficeOpenXml.ExcelPackage]$Excel,
	[parameter(Position = 1, Mandatory=$true, ValueFromPipeline=$true, ValueFromRemainingArguments=$false)] $InputObject,
	[string[]]$Header,
	[string]$WorksheetName = "Worksheet1",
	[string[]]$PivotRows,
	[string[]]$PivotColumns,
	[string[]]$PivotValues,
	[OfficeOpenXml.Drawing.Chart.eChartType]$ChartType,
	[Switch]$Table,
	[OfficeOpenXml.Table.TableStyles]$TableStyle = [OfficeOpenXml.Table.TableStyles]"Medium2",
	[Switch]$AutoFit,
	[Switch]$AppEnd,
	[Switch]$Force,
	[Switch]$ClearSheet,
	[Switch]$ReplaceSheet,
	[Switch]$Passthru
	)
Begin {
	If ( $PSBoundParameters.ContainsKey('Path')) {
		If ( Test-Path $Path ) {
			If ($AppEnd) { Write-Verbose "'$Path' exists. AppEnding data" }
			ElseIf ($Force) {
				Try { Remove-Item -Path $Path -Force -Confirm:$False }
				Catch { Throw "'$Path' exists and could not be removed: $_" }
			}
			Else { Write-Verbose "'$Path' exists. Use -Force to overwrite. Attempting to add sheet to existing workbook" }
			}

		#Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
		$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
		}

	Write-Verbose "Export-XLSX '$($PSCmdlet.ParameterSetName)' PSBoundParameters = $($PSBoundParameters | Out-String)"

	$bound = $PSBoundParameters.keys -contains "InputObject"
	If (-not $bound) { [System.Collections.ArrayList]$AllData = @() }
	}
Process {
	#We write data by row, so need everything countable, not going to stream...
	If ($bound) { $AllData = $InputObject }
	Else { ForEach ($Object in $InputObject) { [void]$AllData.add($Object) } }
	}
End {
	#Deal with headers
	$ExistingHeader = @(($AllData | Select -first 1).PSObject.Properties | Select -ExpandProperty Name)
	
	$Columns = $ExistingHeader.count
	
	If ($Header) {
		If ($Header.count -ne $ExistingHeader.count) { Throw "Found '$columns' columns, provided $($header.count) headers.  You must provide a header For every column." }
		}
	Else { $Header = $ExistingHeader }

	#initialize stuff
	$RowIndex = 2
	Try {
		If ( $PSBoundParameters.ContainsKey('Path')) { $Excel = New-Object OfficeOpenXml.ExcelPackage($Path) -ErrorAction Stop }
		Else { $Path = $Excel.File.FullName }
	
		$Workbook = $Excel.Workbook
		If ($ReplaceSheet) {
			Try {
				Write-Verbose "Attempting to delete worksheet $WorksheetName"
				$Workbook.Worksheets.Delete($WorksheetName)
				}
			Catch {
				If ($_.Exception -notmatch 'Could not find worksheet to delete') {
					Write-Error "Error removing worksheet $WorksheetName"
					Throw $_
					}
				}
			}

		#If we have an excel or valid path, try to appEnd or clearsheet as needed
		If (($AppEnd -or $ClearSheet) -and ($PSBoundParameters.ContainsKey('Excel') -or (Test-Path $Path)) ) {
			$WorkSheet=$Excel.Workbook.Worksheets | Where-Object {$_.Name -like $WorkSheetName}
			If ($ClearSheet) { $WorkSheet.Cells[$WorkSheet.Dimension.Start.Row, $WorkSheet.Dimension.Start.Column, $WorkSheet.Dimension.End.Row, $WorkSheet.Dimension.End.Column].Clear() }
			If ($AppEnd) {
				$RealHeaderCount = $WorkSheet.Dimension.Columns
				If ($Header.count -ne $RealHeaderCount) {
					$Excel.Dispose()
					Throw "Found $RealHeaderCount existing headers, provided data has $($Header.count)."
					}
				$RowIndex = 1 + $Worksheet.Dimension.Rows
				}
			}
		Else { $WorkSheet = $Workbook.Worksheets.Add($WorkSheetName) }
		}
	Catch { Throw "Failed to initialize Excel, Workbook, or Worksheet. Try -ClearSheet Switch If worksheet already exists:`n`n_" }

	#Set those headers If we aren't appEnding
	If (-not $AppEnd) {
		For ($ColumnIndex = 1; $ColumnIndex -le $Header.count; $ColumnIndex++) {
			$WorkSheet.SetValue(1, $ColumnIndex, $Header[$ColumnIndex - 1])
			}
		}

	#Write the data...
	ForEach ($RowData in $AllData) {
		Write-Verbose "Working on object:`n$($RowData | Out-String)"
		For ($ColumnIndex = 1; $ColumnIndex -le $Header.count; $ColumnIndex++) {
			$Object = @($RowData.PSObject.Properties)[$ColumnIndex - 1]
			$Value = $Object.Value
			$WorkSheet.SetValue($RowIndex, $ColumnIndex, $Value)
	
			Try {
				#Nulls will error, catch them
				$ThisType = $Null
				$ThisType = $Value.GetType().FullName
				}
			Catch { Write-Verbose "Applying no style to null in row $RowIndex, column $ColumnIndex" }
	
			#Idea from Philip Thompson, thank you Philip!
			$StyleName = $Null
			$ExistingStyles = @($WorkBook.Styles.NamedStyles | Select -ExpandProperty Name)
			Switch -regex ($ThisType) {
				"double|decimal|single" { $StyleName = 'decimals'; $StyleFormat = "0.00" }
				"int\d\d$" { $StyleName = 'ints'; $StyleFormat = "0" }
				"datetime" { $StyleName = "dates"; $StyleFormat = "M/d/yyy h:mm" }
				"TimeSpan" { $WorkSheet.SetValue($RowIndex, $ColumnIndex, "$Value") }
				default { }#No default yet...
				}

			If ($StyleName) {
				If ($ExistingStyles -notcontains $StyleName) {
					$StyleSheet = $WorkBook.Styles.CreateNamedStyle($StyleName)
					$StyleSheet.Style.NumberFormat.Format = $StyleFormat
					}
				$WorkSheet.Cells.Item($RowIndex, $ColumnIndex).Stylename = $StyleName
				}
			}
		Write-Verbose "Wrote row $RowIndex"
		$RowIndex++
		}

	# Any pivot params specIfied?  add a pivot!
	If ($PSBoundParameters.Keys -match 'Pivot') {
		$Params = @{}
		If ($PivotRows){$Params.Add('PivotRows',$PivotRows)}
		If ($PivotColumns) {$Params.Add('PivotColumns',$PivotColumns)}
		If ($PivotValues)  {$Params.Add('PivotValues',$PivotValues)}
		If ($ChartType){$Params.Add('ChartType',$ChartType)}
		$Excel = Add-PivotTable @Params -Excel $Excel -WorkSheetName $WorksheetName -Passthru -ErrorAction stop
		}

	# Create table
	ElseIf ($Table) { $Excel = Add-Table -Excel $Excel -WorkSheetName $WorksheetName -TableStyle $TableStyle -Passthru }
	
	If ($AutoFit) { $WorkSheet.Cells[$WorkSheet.Dimension.Address].AutoFitColumns() }

	# This is an export command. Save whether we have a path or ExcelPackage input...
	$Excel.SaveAs($Path)
	
	If ($Passthru) { New-Excel -Path $Path }
	}
}

Function Add-Table {
[OutputType([OfficeOpenXml.ExcelPackage])]
[cmdletbinding(DefaultParameterSetName = 'Excel')]
param(
	[parameter(Position = 0,ParameterSetName = 'Excel', Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$false)][OfficeOpenXml.ExcelPackage]$Excel,
	[parameter(Position = 0,ParameterSetName = 'File', Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$false)][validatescript({Test-Path $_})][string]$Path,
	[parameter(Position = 1,Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][string]$WorkSheetName,
	[parameter(Mandatory=$false,ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][int]$StartRow,
	[parameter(Mandatory=$false,ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][int]$StartColumn,
	[parameter(Mandatory=$false,ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][int]$EndRow,
	[parameter(Mandatory=$false,ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][int]$EndColumn,
	[OfficeOpenXml.Table.TableStyles]$TableStyle,
	[string]$TableName,
	[Switch]$Passthru
	)
Process {
	Write-Verbose "PSBoundParameters: $($PSBoundParameters | Out-String)"
	$SourceWS = @{}
	If ($PSBoundParameters.ContainsKey( 'WorkSheetName') ) { $SourceWS.Add('Name',$WorkSheetName) }

	Try {
		If ($PSCmdlet.ParameterSetName -like 'File') { $Excel = New-Excel -Path $Path -ErrorAction Stop }
		$WorkSheets = @( $Excel | Get-Worksheet @SourceWS -ErrorAction Stop )
		}
	Catch	{ Throw "Could not get worksheets to search: $_" }

	If ($WorkSheets.Count -eq 0) { Throw "Something went wrong, we didn't find a worksheet" }

	ForEach ($SourceWorkSheet in $WorkSheets) {
		# Get the coordinates
		$dimension = $SourceWorkSheet.Dimension
		If (-not $StartRow) { $StartRow = $dimension.Start.Row }
		If (-not $StartColumn) { $StartColumn = $dimension.Start.Column }
		If (-not $EndRow) { $EndRow = $dimension.End.Row }
		If (-not $EndColumn) { $EndColumn = $dimension.End.Column }

		$Start = ConvertTo-ExcelCoordinate -Row $StartRow -Column $StartColumn
		$End = ConvertTo-ExcelCoordinate -Row $EndRow -Column $EndColumn
		$RangeCoordinates = "$Start`:$End"

		If (-not $TableName) { $TableWorksheetName = $SourceWorkSheet.Name }
		Else { $TableWorksheetName = $TableName }

		Write-Verbose "Adding table over data range '$RangeCoordinates' with name $TableWorksheetName"
		$Table = $SourceWorkSheet.Tables.Add($SourceWorkSheet.Cells[$RangeCoordinates], $TableWorksheetName)

		If ($TableStyle) {
			Write-Verbose "Adding $TableStyle table style"
			$Table.TableStyle = $TableStyle
			}

		If ($PSCmdlet.ParameterSetName -like 'File' -and -not $Passthru) {
			Write-Verbose "Saving '$($Excel.File)'"
			$Excel.save()
			$Excel.Dispose()
			}
		If ($Passthru) { $Excel }
		}
	}
}

Function New-Excel {
[OutputType([OfficeOpenXml.ExcelPackage])]
[cmdletbinding()]
param(
	[parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][validatescript({$Parent = Split-Path $_ -Parent;If ( -not (Test-Path -Path $Parent -PathType Container) ) { Throw "SpecIfy a valid path.  Parent '$Parent' does not exist: $_" }; $True})][string]$Path
	)
Process {
	If ($path) {
		#Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
		$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
		Write-Verbose "Creating excel object with path '$path'"
		New-Object OfficeOpenXml.ExcelPackage $Path
		}
	Else {
		Write-Verbose "Creating excel object with no specIfied path"
		New-Object OfficeOpenXml.ExcelPackage
		}
	}
}

Function Get-Worksheet {
[OutputType([OfficeOpenXml.ExcelWorksheet])]
[cmdletbinding(DefaultParameterSetName = "Workbook")]
param(
	[parameter(Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][string]$Name,
	[parameter( ParameterSetName = "Workbook", Mandatory=$true, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$false)][OfficeOpenXml.ExcelWorkbook]$Workbook,
	[parameter( ParameterSetName = "Excel", Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$false)][OfficeOpenXml.ExcelPackage]$Excel
	)
Process {
	$Output = Switch ($PSCmdlet.ParameterSetName) {
		"Workbook" { Write-Verbose "Processing Workbook"; $Workbook.Worksheets }
		"Excel" { Write-Verbose "Processing ExcelPackage"; $Excel.Workbook.Worksheets }
		}
	If ($Name) {
		$FilteredOutput = $Output | Where-Object {$_.Name -like $Name}
		If ($Name -notmatch '\*' -and -not $FilteredOutput) { Write-Error "$Name could not be found. Valid worksheets:`n$($Output | Select -ExpandProperty Name | Out-String)" }
		Else { $FilteredOutput }
		}
	Else { $Output }
	}
}

Function Format-Cell {
[OutputType([OfficeOpenXml.ExcelWorksheet])]
[cmdletbinding(DefaultParameterSetname = 'Range')]
param(
	[parameter( Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][OfficeOpenXml.ExcelWorksheet]$WorkSheet,
	[parameter( ParameterSetName = 'Range', Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][int]$StartRow,
	[parameter( ParameterSetName = 'Range', Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][int]$StartColumn,
	[parameter( ParameterSetName = 'Range', Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][int]$EndRow,
	[parameter( ParameterSetName = 'Range', Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)] [int]$EndColumn,
	[parameter( ParameterSetName = 'Header', Mandatory=$true, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][Switch]$Header,
	[boolean]$Bold,
	[boolean]$Italic,
	[boolean]$Underline,
	[int]$Size,
	[string]$Font,
	[System.Drawing.KnownColor]$Color,
	[System.Drawing.KnownColor]$BackgroundColor,
	[OfficeOpenXml.Style.ExcelFillStyle]$FillStyle,
	[boolean]$WrapText,
	[String]$NumberFormat,
	[boolean]$AutoFilter,
	[Switch]$Autofit,
	[double]$AutofitMinWidth,
	[double]$AutofitMaxWidth,
	[OfficeOpenXml.Style.ExcelVerticalAlignment]$VerticalAlignment,
	[OfficeOpenXml.Style.ExcelHorizontalAlignment]$HorizontalAlignment,
	[validateset('Left','Right','Top','Bottom','*')][string[]]$Border,
	[OfficeOpenXml.Style.ExcelBorderStyle]$BorderStyle,
	[System.Drawing.KnownColor]$BorderColor,
	[Switch]$Passthru
	)
Begin {
	If ($PSBoundParameters.ContainsKey('BorderColor')) {
		Try { $BorderColorConverted = [System.Drawing.Color]::FromKnownColor($BorderColor) }
		Catch { Throw "Failed to convert $($BorderColor) to a valid System.Drawing.Color: $_" }
		}

	If ($PSBoundParameters.ContainsKey('Color')) {
		Try { $ColorConverted = [System.Drawing.Color]::FromKnownColor($Color) }
		Catch { Throw "Failed to convert $($Color) to a valid System.Drawing.Color: $_" }
		}

	If ($PSBoundParameters.ContainsKey('BackgroundColor')) {
		Try {
			$BackgroundColorConverted = [System.Drawing.Color]::FromKnownColor($BackgroundColor)
			If (-not $PSBoundParameters.ContainsKey('FillStyle')) { $FillStyle = [OfficeOpenXml.Style.ExcelFillStyle]::Solid }
			}
		Catch { Throw "Failed to convert $($BackgroundColor) to a valid System.Drawing.Color: $_" }
		}
	}
Process {
	#Get the coordinates
	$dimension = $WorkSheet.Dimension

	If ($PSCmdlet.ParameterSetName -like 'Range') {
		If (-not $StartRow) { $StartRow = $dimension.Start.Row }
		If (-not $StartColumn) { $StartColumn = $dimension.Start.Column }
		If (-not $EndRow) { $EndRow = $dimension.End.Row }
		If (-not $EndColumn) { $EndColumn = $dimension.End.Column }
		}
	ElseIf ($PSCmdlet.ParameterSetName -like 'Header') {
		$StartRow = $dimension.Start.Row
		$StartColumn = $dimension.Start.Column
		$EndRow = $dimension.Start.Row
		$EndColumn = $dimension.End.Column
		}

	$Start = ConvertTo-ExcelCoordinate -Row $StartRow -Column $StartColumn
	$End = ConvertTo-ExcelCoordinate -Row $EndRow -Column $EndColumn
	$RangeCoordinates = "$Start`:$End"

	# Apply the Formatting
	$CellRange = $WorkSheet.Cells[$RangeCoordinates]

	Switch ($PSBoundParameters.Keys) {
		'Bold'{ $CellRange.Style.Font.Bold = $Bold  }
		'Italic'  { $CellRange.Style.Font.Italic = $Italic  }
		'Underline'   { $CellRange.Style.Font.UnderLine = $Underline}
		'Size'{ $CellRange.Style.Font.Size = $Size }
		'Font'{ $CellRange.Style.Font.Name = $Font }
		'Color'   { $CellRange.Style.Font.Color.SetColor($ColorConverted) }
		'BackgroundColor' { $CellRange.Style.Fill.PatternType = $FillStyle; $CellRange.Style.Fill.BackgroundColor.SetColor($BackgroundColorConverted) }
		'WrapText'{ $CellRange.Style.WrapText = $WrapText  }
		'VerticalAlignment'   { $CellRange.Style.VerticalAlignment = $VerticalAlignment }
		'HorizontalAlignment' { $CellRange.Style.HorizontalAlignment = $HorizontalAlignment }
		'AutoFilter'  { $CellRange.AutoFilter = $AutoFilter }
		'Autofit' {
				#Probably a cleaner way to call this...
				Try {
					If ($PSBoundParameters.ContainsKey('AutofitMaxWidth')) { $CellRange.AutoFitColumns($AutofitMinWidth, $AutofitMaxWidth) }
					ElseIf ($PSBoundParameters.ContainsKey('AutofitMinWidth')) { $CellRange.AutoFitColumns($AutofitMinWidth) }
					Else { $CellRange.AutoFitColumns() }
					}
				Catch { Write-Error $_ }
				}
		'Border' {
			If ($Border -eq '*') { $Border = 'Top', 'Bottom', 'Left', 'Right' }
			ForEach ($Side in @( $Border | Select -Unique ) ) {
				If (-not $BorderStyle) { $BorderStyle = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin }
				If (-not $BorderColorConverted) { $BorderColorConverted = [System.Drawing.Color]::Black }
				$CellRange.Style.Border.$Side.Style = $BorderStyle
				$CellRange.Style.Border.$Side.Color.SetColor( $BorderColorConverted )
				}
			}
		'NumberFormat' { $CellRange.Style.NumberFormat.Format = $NumberFormat }
		}
	If ($Passthru) { $WorkSheet }
	}
}

Function Set-FreezePane {
[cmdletbinding()]
param(
	[parameter( Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] [OfficeOpenXml.ExcelWorksheet]$WorkSheet,
	[int]$Row = 2,
	[int]$Column = 1,
	[Switch]$Passthru
	)
Process {
	$WorkSheet.View.FreezePanes($Row, $Column)
	If ($Passthru) { $WorkSheet }
	}
}

Function Close-Excel {
[cmdletbinding()]
param(
	[parameter( Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][OfficeOpenXml.ExcelPackage]$Excel,
	[Switch]$Save,
	[parameter( Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)][validatescript({$Parent = Split-Path $_ -Parent -ErrorAction SilentlyContinue; If ( -not (Test-Path -Path $Parent -PathType Container -ErrorAction SilentlyContinue) ) { Throw "SpecIfy a valid path.  Parent '$Parent' does not exist: $_" }; $True})][string]$Path
	)
Process {
	ForEach ($xl in $Excel) {
		Try {
			If ($Path) {
				Try { $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path) }#Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
				Catch { Write-Error "Could not resolve path For '$Path': $_"; Continue }

				write-verbose "Saving $($xl.File) as $($Path)"
				$xl.saveas($Path)
				}
			ElseIf ($Save) {
				write-verbose "Saving $($xl.File)"
				$xl.save()
				}
			}
		Catch { Write-Error "Error saving file.  Will not close this ExcelPackage: $_"; Continue }

		Try {
			write-verbose "Closing $($xl.File)"
			$xl.Dispose()
			$xl = $null
			}
		Catch { Write-Error $_; Continue }
		}
	}
}

#From http://stackoverflow.com/questions/297213/translate-a-column-index-into-an-excel-column-name
Function Get-ExcelColumn {
param([int]$ColumnIndex)

[string]$Chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
$ColumnIndex -= 1
[int]$Quotient = [math]::floor($ColumnIndex / 26)

If ($Quotient -gt 0) { ( Get-ExcelColumn -ColumnIndex $Quotient ) + $Chars[$ColumnIndex % 26] }
Else { $Chars[$ColumnIndex % 26] }
}

Function ConvertTo-ExcelCoordinate {
[OutputType([system.string])]
[cmdletbinding()]
param(
	[int]$Row,
	[int]$Column
	)

$ColumnIndex = Get-ExcelColumn $Column
"$ColumnIndex$Row"
}

