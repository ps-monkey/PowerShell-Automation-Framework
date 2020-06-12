add-type -AssemblyName System.Drawing

##############Load image binaries##############
. $("$PSScriptRoot\Images.ps1") -ea "Stop"

$global:CSS = ""
$global:JS = ""

$UserName = $env:UserName
#$LogonName = $env:UserDomain + "\" + $UserName
$LogonName =  "DOMAIN\" + $UserName
$UserFullName = Try { (Get-ADUser -Filter {SamAccountName -eq $UserName}).name } Catch {"Unknown"}

##############Load CSS colors##############
$MainColor = $global:Config.Properties.style_html.'MainColor'
$tableRowEven =  $global:Config.Properties.style_html.'tableRowEven'
$tableRowOdd = $global:Config.Properties.style_html.'tableRowOdd'
$h1 = $global:Config.Properties.style_html.'h1'
$h2 = $global:Config.Properties.style_html.'h2'
$h3 = $global:Config.Properties.style_html.'h3'
$h4 = $global:Config.Properties.style_html.'h4'

#Define PAFHTML object type
class PAFHTML {
[ValidateNotNullOrEmpty()][string]$Type
[HashTable]$Params
}

Function PAF-New-HTML-Object {
	param (
		[Parameter(Mandatory=$true)][ValidateSet("PAF-HTML-Chart-JS","PAF-HTML-Chart-Rendered","PAF-HTML-CustomHTML","PAF-HTML-Hyperlink","PAF-HTML-Table-L0","PAF-HTML-Table-L1","PAF-HTML-Table-L2","PAF-HTML-LineBreak","PAF-HTML-Table","PAF-HTML-Buttons-Horizontal","PAF-HTML-Buttons-Vertical")][string] $Type,
		[Parameter(Mandatory=$false)][HashTable] $Params
		#[Parameter(Mandatory=$false)] $Data
		)

Return [PAFHTML][Ordered]@{ Type = $Type; Params = $Params; }
}

Function PAF-HTML-Create {
	param ( [Parameter(Mandatory=$true)][AllowEmptyString()][string[]] $Body )


#Add image binaries
PAF-HTML-Add-IMG-CSS -HTML ($Body -join "`n")
	
$HTML = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN""  ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">
<html xmlns=""http://www.w3.org/1999/xhtml"">
<head>
<title>$($global:CustomerObj.ReportName)</title>
<script>
$global:JS
</script>
<style>
*  {
	-moz-box-sizing: border-box;
	-o-box-sizing: border-box;
	-webkit-box-sizing: border-box;
	box-sizing: border-box;
}
html { overflow-y:scroll; }
body { font-family: Verdana; background: #FFFFFF; font-size: 8pt;  }
table { width:100%; }
p { margin:0; }
h1, .h1 { font-size: 24pt; font-weight: bold; margin:0; display: inline-block; vertical-align: bottom; }
h2, .h2 { font-size: 18pt; font-weight: bold; margin:0; display: inline-block; vertical-align: bottom; }
h3, .h2 { font-size: 16pt; font-weight: bold; margin:0; display: inline-block; vertical-align: bottom; }
h4, .h3 { font-size: 12pt; font-weight: bold; margin:0; display: inline-block; vertical-align: bottom; }
.RightAlighend { text-align: right; width:10%; }
.Opposite { display: flex; justify-content: space-between; }
.exp_coll { font-size: 8pt; font-weight: bold; color:#0000EE; vertical-align: bottom; }
.exp_coll_e { font-size: 8pt; font-weight: bold; color:#551A8B; vertical-align: bottom; }
.Error {font-weight: bold; color:red;}
.Warning {font-weight: bold; color:orange;}
.Question {font-weight: bold; color:gray}
.Info {font-weight: bold; color:blue}
.Ok {font-weight: bold; color:green}
.Logo_Img { background-image:url('data:image/png;base64,$Logo_Img_src'); width:89px; height:30px; no-repeat; display: inline-block; }
.thick_hr { height:3px;color: $MainColor; background-color:$MainColor; border: none; }
$global:CSS
</style>
<base target='_blank'>
</head>
<body>
<div style='text-align: center;'><h1 style='color: $MainColor;'>$($global:CustomerObj.CustomerName): $($global:CustomerObj.ReportName)</h1></div>
<div class='Opposite'><span>Generated on: <b>$(Get-Date -format 'dd/MM/yyyy HH:mm')</b></span><span>by $($UserFullName) (<b>$($LogonName)</b>)</span></div>
<hr class='thick_hr'>
$Body
<hr class='thick_hr'>
<div class='Opposite'><span style='font-size:10pt'>2019 | Atos &copy; For internal use</span><span class='Logo_Img'> &nbsp </span></div>
</body>
</html>"

Remove-Variable CSS -Scope "Global" -ea SilentlyContinue
Remove-Variable JS -Scope "Global" -ea SilentlyContinue
#$global:CSS = ""
#$global:JS = ""

#remove img variables
Remove-Variable Error_img -Scope "Global" -ea SilentlyContinue
Remove-Variable Error_img_src -Scope "Global" -ea SilentlyContinue

Remove-Variable Warning_img -Scope "Global" -ea SilentlyContinue
Remove-Variable Warning_img_src -Scope "Global" -ea SilentlyContinue

Remove-Variable Question_img -Scope "Global" -ea SilentlyContinue
Remove-Variable Question_img_src -Scope "Global" -ea SilentlyContinue

Remove-Variable Info_img -Scope "Global" -ea SilentlyContinue
Remove-Variable Info_img_src -Scope "Global" -ea SilentlyContinue

Remove-Variable Ok_img -Scope "Global" -ea SilentlyContinue
Remove-Variable Ok_img_src -Scope "Global" -ea SilentlyContinue


#Remove-Variable CustomerObj -Scope "Global" -ea SilentlyContinue
Return $HTML
}

Function PAF-HTML-Headers {
	param (
		[Parameter(Mandatory=$true)][string] $Code,
		[Parameter(Mandatory=$false)][Switch] $CSS,
		[Parameter(Mandatory=$false)][Switch] $JS,
		[Parameter(Mandatory=$false)][Switch] $Top
		)

If ($CSS) {
	If (!$($global:CSS.contains($Code))) {
		If ($Top) { $global:CSS = $Code + "`r`n" + $global:CSS}
		Else { $global:CSS += "`r`n" + $Code} }
	}
If ($JS) {
	If (!$($global:JS.contains($Code))) {
		If ($Top) { $global:JS = $Code + "`r`n" + $global:JS }
		Else { $global:JS += "`r`n" + $Code }
		}
	}
}

Function PAF-HTML-Chart-JS {
	param ( 
		[Parameter(Mandatory=$true)][ValidateSet("bar","line","pie","horizontalBar")][string] $Type,
		[Parameter(Mandatory=$true)][double[]] $Data,
		[Parameter(Mandatory=$true)][string[]] $Labels, 
		[Parameter(Mandatory=$false)][string[]] $Titles= @("",""), #@(xTitle,yTitle)
		[Parameter(Mandatory=$false)][string] $Name= "",
		[Parameter(Mandatory=$false)][string] $chartID = ([guid]::NewGuid()).Guid -replace "-",
		[Parameter(Mandatory=$false)][string[]] $Size = @("30%","30%"), #(x,y) %'s and absolute values are allowed
		[Parameter(Mandatory=$false)][int] $xAxisMaxTicks = 7,
		[Parameter(Mandatory=$false)][string] $Threshold = ""
		)

$js_code = (Get-Content -Path "$PSScriptRoot\chart.js") -join "`n"
PAF-HTML-Headers -JS -Code $js_code

$js_code = "function colorSchema(index) {
var colors = ['$MainColor','#7337A3','#ACF80F','#27B9BA','#A63050','#1A863B','#F1A426','#0652D1','#8274BC','#C7ED6B','#4F62C1','#38BFA5','#ADCDD2','#26240A','#0BD1D9','#A8C6DD','#DB0674','#4D4DE2','#7B90DB','#E4D91A','#9713C6','#181F0A','#B4CA62','#5D8D11','#FD7E28','#F493AC','#0FC11A','#E4AF28','#1B8804','#C350AF','#8105CB','#0A4886']
return colors[index]
}
function drawGraph(type, chartID, name, xAxisData, yAxisData, xAxisLabel, yAxisLabel, limit, maxTicksLimit) {
var ctx = document.getElementById(chartID);
var yAxisData = eval(yAxisData)
var xAxisData = eval(xAxisData)

if (xAxisLabel == '') { var xAxisLabel = {} }
else { var xAxisLabel = {fontStyle: 'bold', display: true, labelString: xAxisLabel} }

if (yAxisLabel == '') { var yAxisLabel = {} }
else { var yAxisLabel = {fontStyle: 'bold', display: true, labelString: yAxisLabel} }

var max = Math.max.apply(null, yAxisData)
var min = Math.min.apply(null, yAxisData);

if (max > limit) {var yAxisLabel_max = max} else {var yAxisLabel_max = limit}
if (min < 0) {var yAxisLabel_min = min} else {var yAxisLabel_min = 0}

var options = {
	elements: { point:{radius: 0} },
	legend: { display: false },
	animation: { duration: 0 },
	scales: { 
		yAxes: [{ ticks: {suggestedMin: yAxisLabel_min}, scaleLabel: yAxisLabel, position: 'left'},{ ticks: {suggestedMin: yAxisLabel_min}, position: 'right' }],
		xAxes: [{ ticks: {suggestedMin: yAxisLabel_min, maxTicksLimit: maxTicksLimit, autoSkip: true, maxRotation: 0,minRotation: 0}, scaleLabel: xAxisLabel }]
		}
	}

if (type == 'bar' || type == 'horizontalBar') { delete options.scales.xAxes[0].ticks.maxTicksLimit }

if (type == 'horizontalBar' || type == 'pie') { options.scales.yAxes = [{ ticks: {suggestedMin: 0}, scaleLabel: yAxisLabel, position: 'left'}] }
else { options.scales.yAxes[1].ticks.suggestedMax = yAxisLabel_max; }

if (name !='') { options.title = {display: true, text: name} }

if (limit !='' && type !='pie') { options.annotation = { annotations: [{ type: 'line', mode: 'horizontal', scaleID: 'y-axis-0', value: limit, borderColor: '#ff0000', borderWidth: 2 }] } }

if (type == 'pie') { 
	var coloR = [];
	for (var i = 0; i < yAxisData.length; i++) { coloR.push(colorSchema(i)) }
	options.scales.xAxes[0].display= false
	options.scales.yAxes[0].display= false
	options.legend.display= true
	}
else { var coloR = '$MainColor' }

var data = { labels: xAxisData, datasets: [{label: name, lineTension: 0, fill: true, backgroundColor: coloR, data: yAxisData}] }
var config = {type: type, options: options, data: data}
var myChart = new Chart(ctx, config);
}"
PAF-HTML-Headers -JS -Code $js_code

$xAxisData = ([guid]::NewGuid()).Guid -replace "-" -replace "^[0-9\s]+"
$yAxisData = ([guid]::NewGuid()).Guid -replace "-" -replace "^[0-9\s]+"

$css_code = ".Graph { display:inline-block; }"
PAF-HTML-Headers -CSS -Code $css_code

Return "<div class='Graph' style='width:$($Size[0]);height:$($Size[1])'><canvas id='$chartID'></canvas></div>
<script>
var $xAxisData = ['$($Labels -join ''',''')']
var $yAxisData = [$($Data -join ',')]
drawGraph('$Type','$chartID','$Name', '$xAxisData', '$yAxisData', '$($Titles[0])' ,'$($Titles[1])','$Threshold','$xAxisMaxTicks')
</script>`n"
}

Function PAF-HTML-Chart-Rendered {
	param ( 
		[Parameter(Mandatory=$true)][ValidateSet("bar","line","pie","horizontalBar")][string] $Type,
		[Parameter(Mandatory=$true)][double[]] $Data,
		[Parameter(Mandatory=$true)][string[]] $Labels, 
		[Parameter(Mandatory=$false)][string[]] $Titles = @("",""), #@(xTitle,yTitle)
		[Parameter(Mandatory=$false)][string] $Name,
		[Parameter(Mandatory=$false)][int[]] $Size = @(300,400), #only absolute values are allowed
		[Parameter(Mandatory=$false)][int] $xAxisMaxTicks,
		[Parameter(Mandatory=$false)][string] $Threshold = ""
		)

$css_code = ".Graph {display:inline-block;}"
PAF-HTML-Headers -CSS -Code $css_code

$color = "#636363"
$Interval = 1
If ($xAxisMaxTicks) { $xInterval = [math]::Ceiling($Data.count/$xAxisMaxTicks) }

$yMax = ($Data | measure -Maximum).Maximum
If ([int]$Threshold -gt $yMax) { $yMax = $Threshold }

$order = [Math]::Pow(10,$($($yMax.ToString()).Length-1))
If ($yMax/$order -eq 1) { $yInterval = $order/10 }
ElseIf ($yMax/$order -le 2) {
	$R = 0
	$yMax = ([math]::DivRem([math]::Ceiling($yMax/$order*10),2,[ref]$R) + $r)*2*$order/10
	$yInterval = $order/10 * 2
	}
ElseIf ($yMax/$order -le 5) {
	If ([math]::Round($yMax/$order) -eq [math]::Ceiling($yMax/$order)) { $yMax = [math]::Ceiling($yMax/$order)*$order }
	Else { $yMax = $yMax = [math]::Floor($yMax/$order)*$order + $order/2 }
	$yInterval = $order/2
	}
ElseIf ($yMax/$order -lt 10) {
	$yMax = [math]::Ceiling($yMax/$order)*$order
	$yInterval = $order
	}

Switch -regex ($Type) {
"bar" { $ChartType = "column" }
"line" { $ChartType = "area" }
"pie" { $ChartType = "pie" }
"horizontalBar" { $ChartType = "bar" }
}

$font =  New-Object system.drawing.font("Arial",10,[system.drawing.fontstyle]::regular)
$tempfile = "$PSScriptRoot\Chart.png"
$Chart = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Chart
$Chart.Size = "$($Size[0]),$($Size[1])"
If ($Name) {
	$Chart.Titles.Add($Name) | out-null
	$Chart.Titles[0].Font = New-Object system.drawing.font("Arial",10,[system.drawing.fontstyle]::bold)
	$Chart.Titles[0].ForeColor = $color
	}
$ChartArea = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.ChartArea
$ChartArea.AxisX.Title = $Titles[0]

$ChartArea.AxisX.Interval = $xInterval
$ChartArea.AxisX.LabelStyle.Enabled = $true
#$ChartArea.AxisX.LabelStyle.Angle = 90
$ChartArea.AxisX.Minimum = 1;
$ChartArea.AxisX.Maximum = $labels.count
$ChartArea.AxisX.MajorGrid.Enabled = $true
$ChartArea.AxisX.MajorGrid.LineColor = "#e7e7e7"

$ChartArea.AxisY.Title = $Titles[1]
$ChartArea.AxisY.Maximum = $yMax
$ChartArea.AxisY.Interval = $yInterval
$ChartArea.AxisY.MajorGrid.Enabled = $true
$ChartArea.AxisY.MajorGrid.LineColor = "#cecece"

$ChartArea.AxisY2.Enabled =  "True"
$ChartArea.AxisY2.Maximum = $yMax
$ChartArea.AxisY2.Interval = $yInterval
$ChartArea.AxisY2.MajorGrid.Enabled = $false

$Chart.ChartAreas.Add($ChartArea)

ForEach ($Axe in $Chart.ChartAreas[0].Axes) {
	$Axe.Titlefont = New-Object system.drawing.font("Arial",10,[system.drawing.fontstyle]::bold)
	$Axe.TitleForeColor = $color
	$Axe.LineColor = $color
	$Axe.LabelStyle.ForeColor = $color
	$Axe.LabelStyle.Font = $font
	$Axe.MajorTickMark.LineColor = "#e7e7e7"
	}

$Chart.Series.Add($Name) | out-null
$Chart.Series[0].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::$ChartType
$Chart.Series[0].Points.DataBindXY($Labels,$Data)
$Chart.Series[0].Color = "$MainColor"


If ($Threshold -and $($type -ne "horizontalBar" -and $type -ne "pie")) {
	$Thresholds = @()
	$Data | % { $Thresholds += [double]$Threshold }
	$Chart.Series.Add('Limit') | out-null
	$Chart.Series['Limit'].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
	$Chart.Series['Limit'].Color = "red" 
	$Chart.Series['Limit'].Points.DataBindXY($Labels,$Thresholds)
	}

#save to base64
If ((Test-Path $tempfile)) { Remove-Item $tempfile -Force }
$Chart.SaveImage( $tempfile  ,"png")

$base64 = [Convert]::ToBase64String((Get-Content $tempfile -Encoding Byte))
Remove-Item $tempfile -Force

$css_code = "background-image:url('data:image/png;base64,$base64'); width: $($Size[0])px; height:$($Size[1])px; no-repeat; display: inline-block"

Return "<div class='Graph' style='width:$($Size[0]);height:$($Size[1])'><span style = ""$css_code""> &nbsp </span></div>"
}

Function PAF-HTML-EmbededIMG {
	param (
		[Parameter(Mandatory=$true)][string] $SourcePath,
		[Parameter(Mandatory=$false)][int[]] $Size #only absolute values are allowed
		)

$base64 = [Convert]::ToBase64String((Get-Content $SourcePath -Encoding Byte))

If (!$Size) {
	$png = New-Object System.Drawing.Bitmap $SourcePath
	$Size = @($png.Width,$png.Height)
	}
$css_code = "background-image:url('data:image/png;base64,$base64'); width: $($Size[0])px; height:$($Size[1])px; no-repeat; display: inline-block"

Return "<span style = ""$css_code""> &nbsp </span>"


}

Function PAF-HTML-CustomHTML {
	param ( 
		[Parameter(Mandatory=$false)][string] $Data,
		[Parameter(Mandatory=$false)][string] $CustomJS,
		[Parameter(Mandatory=$false)][string] $CustomCSS
		)
If ($CustomJS) { PAF-HTML-Headers -JS -Code $CustomJS }
If ($CustomCSS) { PAF-HTML-Headers -CSS -Code $CustomCSS }

Return "$Data"
}

Function PAF-HTML-Add-IMG-CSS {
	param (
		[Parameter(Mandatory=$false)][string] $HTML
		)
$css_code = ""
If ($HTML -match "class='Error_img") {
	$css_code += "`n.Error_img, .Error_img_s, .Error_img_i { background-image:url('data:image/png;base64,$Error_img_src'); width:16px; height:16px; no-repeat; display: inline-block; }"
	}
If ($HTML -match "class='Warning_img") {
	$css_code += "`n.Warning_img, .Warning_img_s, .Warning_img_i { background-image:url('data:image/png;base64,$Warning_img_src'); width:16px; height:16px; no-repeat; display: inline-block; }"
	}
If ($HTML -match "class='Question_img") {
	$css_code += "`n.Question_img, .Question_img_s, .Question_img_i { background-image:url('data:image/png;base64,$Question_img_src'); width:16px; height:16px; no-repeat; display: inline-block; }"
	}
If ($HTML -match "class='Info_img") {
	$css_code += "`n.Info_img, .Info_img_s, .Info_img_i { background-image:url('data:image/png;base64,$Info_img_src'); width:16px; height:16px;  no-repeat; display: inline-block;}"
	}
If ($HTML -match "class='Ok_img") {
	$css_code += "`n.Ok_img, .Ok_img_s, .Ok_img_i { background-image:url('data:image/png;base64,$Ok_img_src'); width:16px; height:16px; no-repeat; display: inline-block; }"
	}
$css_code += "`n.Error_img_i, .Warning_img_i, .Question_img_i, .Info_img_i, .Ok_img_i {display:none;}"

PAF-HTML-Headers -CSS -Code $css_code
}

Function PAF-HTML-Hyperlink {
	param ( 
		[Parameter(Mandatory=$false)][string] $Data,
		[Parameter(Mandatory=$false)][string] $Text,
		[Parameter(Mandatory=$false)][string] $Function,
		[Parameter(Mandatory=$false)][string] $CustomCSS
		)

If ($Function -match "ShowHideE") {
	$Data = "#"
	$Text = "collapse"
	$Function = $Function -replace "ShowHideE","ShowHide"
	$Function = " class='exp_coll_e' onClick=""$Function"""
	}
ElseIf ($Function -match "ShowHide") {
	$Data = "#"
	$Text = "expand"
	$Function = " class='exp_coll' onClick=""$Function"""
	}
Else { $Function = " onClick=""$Function""" }
If ($Data) { If (!$Text) {$Text = $Data } }
If ($CustomCSS) { $CustomCSS = " style='$CustomCSS'" }
Return "<a$Function href=""$Data""$CustomCSS>$Text</a>"
}

Function PAF-HTML-Table-L0 {
	param (
		[Parameter(Mandatory=$true)][string] $Data,
		[Parameter(Mandatory=$true)][string] $Title,
		[Parameter(Mandatory=$false)][Switch] $ColoredTitle,
		[Parameter(Mandatory=$false)][Switch] $Expanded,
		#[Parameter(Mandatory=$false)][ValidateSet("Vertical","Horizontal")][string] $Direction = "Vertical", ####TO DO!!!!
		[Parameter(Mandatory=$false)][Switch] $ShowState
		)

$css_code = ".Table-L0, .Table-L0 tr, .Table-L0 td {border: 0px; padding: 5px; }
.Table-L0-Title { text-align: center; }"
PAF-HTML-Headers -CSS -Code $css_code

$js_code = "function ShowHide(TableID,el) {
if (document.getElementById(TableID).style.display == 'none') {
	if (document.getElementById(TableID).tagName == 'TR') { var display = 'table-row'  }
	else { var display = 'table' }
	document.getElementById(TableID).style.display = display
	el.textContent  = 'collapse'
	el.style.color  = '#551A8B'
	}
else {
	document.getElementById(TableID).style.display = 'none'
	el.textContent  = 'expand'
	el.style.color  = '#0000EE'
	}
}"
PAF-HTML-Headers -JS -Code $js_code

If ($ShowState) {
	$css_code = ".Overall_Error_img, .Overall_Warning_img, .Overall_Question_img, .Overall_Info_img, .Overall_Ok_img {width:32px; height:32px; no-repeat; display: inline-block; }
.Overall_Error_img { background-image:url('data:image/png;base64,$Overall_Error_img_src'); }
.Overall_Warning_img { background-image:url('data:image/png;base64,$Overall_Warning_img_src'); }
.Overall_Question_img { background-image:url('data:image/png;base64,$Overall_Question_img_src'); }
.Overall_Info_img { background-image:url('data:image/png;base64,$Overall_Info_img_src'); }
.Overall_Ok_img { background-image:url('data:image/png;base64,$Overall_Ok_img_src'); }"
	PAF-HTML-Headers -CSS -Code $css_code
	
	$Table_L0_State = If ($Data -match "Error_img'|Error_img_i") { $Overall_Error_img } ElseIf ($Data -match "Warning_img'|Warning_img_i") {  $Overall_Warning_img } ElseIf ($Data -match "Question_img'|Question_img_i") { $Overall_Question_img } ElseIf ($Data -match "Info_img'|Info_img_i") { $Overall_Info_img } Else { $Overall_Ok_img }
	} 
Else { $Table_L0_State = "" }

$ID = ([guid]::NewGuid()).Guid -replace "-" -replace "^[0-9\s]+"

If ($Expanded) {
	$ShowHide = "ShowHideE('$ID',this);return false;"
	$display = 'table'
	}
Else {
	$ShowHide = "ShowHide('$ID',this);return false;"
	$display = 'none'
	}

$style = ""
If ($ColoredTitle) { $style  = " style='color: $MainColor;'"}

Return "`n<table class=""Table-L0"">
<tr><td class='Table-L0-Title'>$Table_L0_State <h2$style>$Title</h2> $(PAF-HTML-Hyperlink -Function $ShowHide)</td></tr>
<tr><td><table id='$ID' style='display:$display;'>
<tr><td>$Data</td></tr></table>
</td></tr>
</table>`n"
}

Function PAF-HTML-Table-L1 {
	param ( 
		[Parameter(Mandatory=$true)][string] $Data,
		[Parameter(Mandatory=$true)][string] $Title,
		[Parameter(Mandatory=$false)][Switch] $ColoredTitle,
		[Parameter(Mandatory=$false)][string] $URL,
		[Parameter(Mandatory=$false)][Switch] $Expanded,
		#[Parameter(Mandatory=$false)][ValidateSet("Vertical","Horizontal")][string] $Direction = "Vertical", ####TO DO!!!!
		[Parameter(Mandatory=$false)][Switch] $ShowState
		)

$css_code = ".Table-L1 { border: 1px solid $MainColor; margin-bottom: 10px; }
.Table-L1 .Opposite, .Table-L1 .Highlight { font-weight: bold; }"

PAF-HTML-Headers -CSS -Code $css_code

$js_code = "function ShowHide(TableID,el) {
if (document.getElementById(TableID).style.display == 'none') {
	if (document.getElementById(TableID).tagName == 'TR') { var display = 'table-row'  }
	else { var display = 'table' }
	document.getElementById(TableID).style.display = display
	el.textContent  = 'collapse'
	el.style.color  = '#551A8B'
	}
else {
	document.getElementById(TableID).style.display = 'none'
	el.textContent  = 'expand'
	el.style.color  = '#0000EE'
	}
}"
PAF-HTML-Headers -JS -Code $js_code

$ID = ([guid]::NewGuid()).Guid -replace "-"
If ($Expanded) {
	$ShowHide = "ShowHideE('$ID',this);return false;"
	$display = 'table-row'
	}
Else {
	$ShowHide = "ShowHide('$ID',this);return false;"
	$display = 'none'
	}

#Draw states count bar	
If ($ShowState) {
	$e_count = If ($Data -match "$Error_img|$Error_img_i") { "<b>$([regex]::Matches($Data,""$Error_img|$Error_img_i"").count) </b>x $($Error_img_s) $($Error_msg)" } Else {""}
	$w_count = If ($Data -match "$Warning_img|$Warning_img_i") { "<b>$([regex]::Matches($Data,""$Warning_img|$Warning_img_i"").count) </b>x $($Warning_img_s) $($Warning_msg)," } Else {""}
	$q_count = If ($Data -match "$Question_img|$Question_img_i") { "<b>$([regex]::Matches($Data,""$Question_img|$Question_img_i"").count) </b>x $($Question_img_s) $($Question_msg)," } Else {""}
	$i_count = If ($Data -match "$Info_img|$Info_img_i") { "<b>$([regex]::Matches($Data,""$Info_img|$Info_img_i"").count) </b>x $($Info_img_s) $($Info_msg)," } Else {""}
	
	$Table_L1_State_Text = If ($e_count -or $w_count -or $q_count -or $i_count) { "$($i_count) $($q_count) $($w_count) $($e_count)" } Else { $Ok_msg }
	$Table_L1_State = "`n<tr><td>$($Title):</td><td class='RightAlighend' style='white-space:nowrap;'>$Table_L1_State_Text</td></tr>"
	}


If ($URL) { $URL = "($URL)"}
$style = ""
If ($ColoredTitle) { $style  = " style='color: $MainColor;'"}

Return "`n<table class='Table-L1'>
<tr><td colspan=""2"" style='text-align:left'><h3$style>$Title $URL</h3></td></tr>$Table_L1_State
<tr><td colspan=""2""><div class=""Opposite""><span>Detailed view:</span><span>$(PAF-HTML-Hyperlink -Function $ShowHide)</span></div></td></tr>
<tr id='$ID' style='display:$display;'><td colspan=""2"">$Data</td></tr></table>`n"
}

Function PAF-HTML-Table-L2 {
	param ( 
		[Parameter(Mandatory=$true)][AllowEmptyString()][string] $Data,
		[Parameter(Mandatory=$true)][string] $Title,
		[Parameter(Mandatory=$false)][Switch] $ColoredTitle,
		[Parameter(Mandatory=$false)][string] $Error,
		#[Parameter(Mandatory=$false)][ValidateSet("Vertical","Horizontal")][string] $Direction = "Vertical", ####TO DO!!!!
		[Parameter(Mandatory=$false)][Switch] $ShowState,
		[Parameter(Mandatory=$false)][Switch] $Expandable,
		[Parameter(Mandatory=$false)][Switch] $Expanded,
		[Parameter(Mandatory=$false)][AllowEmptyString()][string] $EmptyObjectState,
		[Parameter(Mandatory=$false)][AllowEmptyString()][string] $EmptyObjectMessage
		)

$css_code = ".Table-L2 { display: flex; justify-content: space-between;  }
.Table-L2:hover { background: #BDCBE9; }"
PAF-HTML-Headers -CSS -Code $css_code

$js_code = "function ShowHide(TableID,el) {
if (document.getElementById(TableID).style.display == 'none') {
	if (document.getElementById(TableID).tagName == 'TR') { var display = 'table-row'  }
	else { var display = 'table' }
	document.getElementById(TableID).style.display = display
	el.textContent  = 'collapse'
	el.style.color  = '#551A8B'
	}
else {
	document.getElementById(TableID).style.display = 'none'
	el.textContent  = 'expand'
	el.style.color  = '#0000EE'
	}
}"
PAF-HTML-Headers -JS -Code $js_code

$ID = ([guid]::NewGuid()).Guid -replace "-"

If ($Expanded) { $ShowHide = "ShowHideE('$ID',this);return false;" }
Else { $ShowHide = "ShowHide('$ID',this);return false;" }

#exit with error
If ($Error) { Return "`n<table$TableWidth class='Table-Status'><tr><th>$Error $Error_img_i</th></tr></table>" }

If ($Data) {
	$UnknownState = $false
	
	ForEach ($line in $($Data -split "</tr> <tr>") | Select-Object -Skip 1) { If ($line -notmatch "Error_img|Warning_img|Ok_img|Info_img") { $UnknownState = $true } }
	$Table_L2_State = PAF-HTML-Table-Status-Get-State -StateList $Data -UnknownState $UnknownState
	
	If ($Expandable) {
		$Result_Header = "`n<div$TableWidth class='Table-L2'><span class='Opposite'><span style='width:75px;text-align:right'>$Table_L2_State </span><span> $($Title):</span></span><span>$(PAF-HTML-Hyperlink -Function $ShowHide)</span></div><hr>"
		If ($Expanded) { $Result_Header += "`n<div id='$ID' style='display:table; color:#551A8B; width:100%'>$Data`n</div>" }
		Else { $Result_Header += "`n<div id='$ID' style='display:none; width:100%'>$Data`n</div>" }
		}
	Else { $Result_Header = "`n<div$TableWidth style='display: flex; justify-content: space-between;'><span class='Opposite'><span style='width:75px;text-align:right'>$Table_L2_State </span><span> $($Title):</span></span></div><hr>`n$Data" }
	}
Else {
	If (!$EmptyObjectState) { $EmptyObjectState = $Ok_img }
	$Result_Header = "`n<div class='Opposite'><span class='Table-L2'><span style='width:75px;text-align:right'>$($EmptyObjectState -replace ""'>"", ""_s'>"") </span><span> $($Title):</span></span><span>$(PAF-HTML-Hyperlink -Function $ShowHide)</span></div><hr>"
	#Threat all empty object or state as OK
	If (!$EmptyObjectMessage) { $EmptyObjectMessage = "No issues found" }
	$Result_Header += "`n<table$TableWidth class='Table-Status' id='$ID' style='display:none;'><tr><th>$($EmptyObjectState -replace ""'>"", ""_i'>"") $EmptyObjectMessage</th></tr></table>"
	}

Return "$Result_Header"
}

Function PAF-HTML-LineBreak {
	param (
		[Parameter(Mandatory=$false)][Switch] $Drawline,
		[Parameter(Mandatory=$false)][Switch] $Thick
		)

$HTML = "<br>"
If ($Drawline) {
	If ($Thick) { $HTML += "<hr class='thick_hr'>" } 
	Else { $HTML += "<hr>" }
	}
Return "$HTML`n"
}

Function PAF-HTML-Table-Status-Get-State {
	param ( 
		[Parameter(Mandatory=$true)][string[]] $StateList,
		[Parameter(Mandatory=$true)][bool] $UnknownState
		)

$StatesList = ""	
#Check for error/warning/etc states
If ($StateList -match "Info_img|Question_img|Warning_img|Error_img") {
	If ($StateList -match "Info_img'|Info_img_i") { $StatesList += $Info_img_s}
	If ($StateList -match "Question_img'|Question_img_i") { $StatesList += $Question_img_s}
	If ($StateList -match "Warning_img'|Warning_img_i") { $StatesList += $Warning_img_s}
	If ($StateList -match "Error_img'|Error_img_i") { $StatesList += $Error_img_s }
	Return $StatesList
	}
Else {
	#Check if there are unknown states discovered
	If ($UnknownState) { Return $Question_img_s }
	Else { Return $Ok_img_s } 
	}
}

Function PAF-HTML-Table {
	param (
		[Parameter(Mandatory=$true)][PSCustomObject[]] $Data,
		[Parameter(Mandatory=$false)][string] $ID = ([guid]::NewGuid()).Guid -replace "-",
		[Parameter(Mandatory=$false)][string] $Title,
		[Parameter(Mandatory=$false)][Switch] $ColoredTitle,
		[Parameter(Mandatory=$false)][string] $Width = "", #"XXX px" or XXX %"
		[Parameter(Mandatory=$false)][string] $SortBy = $false,
		[Parameter(Mandatory=$false)][string[]] $CustomSortOrder, ###define type
		[Parameter(Mandatory=$false)][Switch] $Descending,
		[Parameter(Mandatory=$false)][ValidateCount(2,2)][int[]] $Thresholds = @(),
		[Parameter(Mandatory=$false)][string[]] $ThresholdPropertyNames,
		[Parameter(Mandatory=$false)][ValidateSet("Below","Above")][string] $Watermark = "Above",
		[Parameter(Mandatory=$false)][Switch] $ShowLegend,
		[Parameter(Mandatory=$false)][Switch] $Searchable,
		[Parameter(Mandatory=$false)][Switch] $HideStatusText,
		[Parameter(Mandatory=$false)][Hashtable] $ErrorsHash,
		[Parameter(Mandatory=$false)][Switch] $FormatedTable
		)
#Madatory CSS 
$css_code = ".Table-Regular { border: 0px; padding: 5px; text-align: center; }
.Table-Regular th { font-weight: bold; background-color: $MainColor; color: #FFFFFF; font-size: 12pt; padding: 5px; white-space: nowrap; }
.Table-Regular tr:nth-child(2n+1) { background: $tableRowOdd; }
.Table-Regular tr:nth-child(2n) { background: $tableRowEven; }
.Table-Regular td { width:10%; padding: 5px; }"
PAF-HTML-Headers -CSS -Code $css_code

If ($Searchable) { 
	$css_code = ".Search { background-repeat: no-repeat; padding-left: 18px; border: 1px solid $MainColor; margin-bottom: 5px; width: 100%; background-image:url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABy0lEQVQ4T5WS28tpURTFx3KPJ+SBRLlL5JHwn0vJmzxQEiFyzW2HFLmt01jnbKfd5+GcVbvdXnPu3xxzzCmklBIAHo8HLpcLZrMZTqcT3u83XC4XAoEAvF4vHA4H034cQQB/HI/H2G63MJlMhiTynU4n4vE4fD7fT8D9fpe9Xg+HwwF6MqtarVbsdjvs93sF5Xcmk4Hb7TZAhKZpstVqqSRKzefzhgS2NBwOIYSA3+9HMpk0qBTtdluyCmUWi8WvfXa7XWw2G6WiUCjAZrN98kS9Xpc0MBaLIRwOfwXcbjc0Gg0Vq1QqBkNFrVaTdDydTivHvx16U61WVahUKim1+hHNZlOez2flcC6X+wpYLBbo9/uq93K5bGxhPp9LPZhIJBAMBg2Q4/GITqeD1+sFj8eDbDYLs9n8V8H1elVToA9shW3QDxq2XC4xGo3UeBnjfSgUMo6Ri8QFGgwGCsJkVuOhZH2x+H4+n0ilUqqIfq82kcmapmG9XqvlIUgHUDZXmvtA6VRCFZFIRH1/APyBQVbR27Hb7b+ThMBkMlGPxWIxQAyAryP4c8m2uJGr1UrJZyFu5T8DyCGEKqbT6cef/wIQQsuogtOJRqP4BZ0+DW+X2f7AAAAAAElFTkSuQmCC');}
.Table-Regular th a, .Table-Regular td a { width: 100%; }
.Table-Regular th a.sort-by { padding-right: 18px; position: relative; }
.Table-Regular a:link, .Table-Regular a:visited, .Table-Regular a:hover, .Table-Regular a:active { color: #FFFFFF; text-decoration: none; }
a.sort-by:before, a.sort-by:after {border: 4px solid transparent; content: ''; display: block; height: 0; right: 5px; top: 50%; position: absolute; width: 0; }
a.sort-by:before { border-bottom-color: #FFFFFF; margin-top: -9px; }
a.sort-by:after { border-top-color: #FFFFFF; margin-top: 1px; }"

	$js_code = "var TableLastSortedColumn = -1;

function sortTable(el) {
var sortColumn = el.parentElement.cellIndex;
while ((el = el.parentElement) && el.nodeName.toUpperCase() !== 'TABLE');

var table = document.getElementById(el.id);
var tbody = table.getElementsByTagName('tbody')[0];
var rows = tbody.getElementsByTagName('tr');
var arrayOfRows = new Array();
var type = 'T' 
for(var i=0, len=rows.length; i<len; i++) {
	arrayOfRows[i] = new Object;
	arrayOfRows[i].oldIndex = i;
	var celltext = rows[i].getElementsByTagName('td')[sortColumn].innerText;
	if (Number(celltext.replace(/ /g, ''))) { type = 'N' }
	var re = type=='N' ? /[^\.\-\+\d]/g : /[^a-zA-Z0-9]/g;
		arrayOfRows[i].value = celltext.replace(re,"""").substr(0,25).toLowerCase();
	}
if (sortColumn == TableLastSortedColumn) { arrayOfRows.reverse(); }
else {
	TableLastSortedColumn = sortColumn;
	switch(type) {
		case 'N' : arrayOfRows.sort(CompareRowOfNumbers); break;
		case 'T'  : arrayOfRows.sort(CompareRowOfText);
		}
	}
var newTableBody = document.createElement('tbody');
for(var i=0, len=arrayOfRows.length; i<len; i++) { newTableBody.appendChild(rows[arrayOfRows[i].oldIndex].cloneNode(true)); }
table.replaceChild(newTableBody,tbody);
}

function CompareRowOfText(a,b) {
var aval = a.value;
var bval = b.value;
return( aval == bval ? 0 : (aval > bval ? 1 : -1) );
}

function CompareRowOfNumbers(a,b) {
var aval = /\d/.test(a.value) ? parseFloat(a.value) : 0;
var bval = /\d/.test(b.value) ? parseFloat(b.value) : 0;
return( aval == bval ? 0 : (aval > bval ? 1 : -1) );
}

function search(el) {
while ((el = el.parentElement) && el.nodeName.toLowerCase() !== 'table');
var table = document.getElementById(el.id);
var tr = table.getElementsByTagName('tr');
var txtValue = new Array(tr[0].getElementsByTagName('th').length);
var filter = new Array(tr[0].getElementsByTagName('th').length);

for (i = 0 ; i < tr[0].getElementsByTagName('th').length; i++) {
	filter[i] = tr[1].getElementsByTagName('th')[i].getElementsByTagName('input')[0].value.toUpperCase()
	}

	for (i = 2; i < tr.length; i++) {
		for (j = 0 ; j < tr[0].getElementsByTagName('th').length; j++) {
			td = tr[i].getElementsByTagName('td')[j];
			txtValue[j] = td.innerText.toUpperCase().indexOf(filter[j]);
			}
		var display = 'show'
		for (k = 0; k < txtValue.length; k++) { if (txtValue[k] == -1) { display = 'hide' } }
		
		if (display == 'show') { tr[i].style.display = '' }
		else { tr[i].style.display = 'none' }
	}
}"
	PAF-HTML-Headers -JS -Code $js_code
	PAF-HTML-Headers -CSS -Code $css_code
	}
		
$WarningThreshold = [int]$Thresholds[0]
$ErrorThreshold = [int]$Thresholds[1]

If ($Title) {
	$style = ""
	If ($ColoredTitle) { $style  = " style='color: $MainColor;'"}
	$Title = "<h4$style>$Title</h4>"
	}


If (!$FormatedTable) {

	If ($Width) { $TableWidth = " style='width:$($Size[0]);height:$($Size[1])'"}

	#Add first property with threshold status to all objects
	If ($ThresholdPropertyNames.count) {
		ForEach ($obj in $Data) {
			$hashtable = [ordered]@{}
			$hashtable["chkResult"] = "ok"
			ForEach ($ThresholdPropertyName in $ThresholdPropertyNames) {
				If ($Watermark -eq "Above") {
					If ($obj.$ThresholdPropertyName -ge $ErrorThreshold) { $hashtable["chkResult"] =  "error" } 
					ElseIf ($($obj.$ThresholdPropertyName -ge $WarningThreshold) -and $($hashtable["chkResult"] -eq "ok")) { $hashtable["chkResult"] =  "warning" }
					}
				Else {
					If ($obj.$ThresholdPropertyName -le $ErrorThreshold) { $hashtable["chkResult"] =  "error" }
					ElseIf ($($obj.$ThresholdPropertyName -le $WarningThreshold) -and $($hashtable["chkResult"] -eq "ok")) { $hashtable["chkResult"] =  "warning" }
					}
				}
			ForEach ($property in $obj.PSObject.properties.name) { $hashtable[$property] = $obj.$property }
			$Data = $Data -ne $obj
			$Data += New-Object -TypeName PSObject -Property $Hashtable
			}
		}

	$Params = @{}

	If ($Searchable) {
		$Props = $Data[0].PSObject.Properties.Name | ? {$_ -ne "chkResult"}
		$InputSearch = ""
		$Props  | % { $InputSearch += "<th><input type='text' class='Search' onkeyup='search(this)' placeholder='Search...'></th>" }
		ForEach ($Prop in $Props) {
			$new = "<a href='#' class='sort-by' onClick=""sortTable(this);return false;"">$Prop</a>"
			$Data = $Data | Add-Member -MemberType AliasProperty -Name $new -Value $Prop -PassThru
			$Data = $Data | Select-Object -Property * -ExcludeProperty $Prop
			}
		If ($SortBy -ne "$false") {
			$SortBy = "<a href='#' class='sort-by' onClick=""sortTable(this);return false;"">$SortBy</a>"
			If (!$CustomSortOrder) { $Params.Add("Property", $SortBy) }
			Else {
				$SortBy = '$_.''' + $SortBy +"'"
				$Switch = "Switch($SortBy){ "
				For($i = 0; $i -lt $CustomSortOrder.count; $i++) { $Switch = $Switch + "'" + $CustomSortOrder[$i] + "' {$i}; "  }
				$Switch = $Switch + "}"
				$Params.Add("Property" ,@{ e = { Invoke-Expression $Switch } })
				}
			}
		Else { If ($ThresholdPropertyNames.count) { $Params.Add("Property", "<a href='#' class='sort-by' onClick=""sortTable(this);return false;"">$($ThresholdPropertyNames[0])</a>") } }
		}
	Else {
		If ($SortBy -ne "$false") {
			If (!$CustomSortOrder) { $Params.Add("Property", $SortBy) }
			Else {
				$SortBy = '$_.''' + $SortBy +"'"
				$Switch = "Switch($SortBy){ "
				For($i = 0; $i -lt $CustomSortOrder.count; $i++) { $Switch = $Switch + "'" + $CustomSortOrder[$i] + "' {$i}; "  }
				$Switch = $Switch + "}"
				$Params.Add("Property" ,@{ e = { Invoke-Expression $Switch } })
				}
			}
		Else { If ($ThresholdPropertyNames.count) { $Params.Add("Property", $ThresholdPropertyNames[0]) } }
		}

	If ($Descending) { $Params.Add("Descending", $true) }
	
	If ($Params.count) { $Data = $Data | Sort-Object @Params | ConvertTo-Html -Fragment }
	Else { $Data = $Data | ConvertTo-Html -Fragment }

	$Data = $Data -replace "&lt;", "<" -replace "&gt;", ">" -replace "&#39;", "'"

	
	$Data = $Data -replace "<table>", "$Title`n<table$TableWidth class='Table-Regular' id='$ID'>" -replace "<th>chkResult</th>",""
	$Data = $Data -replace "</th></tr>","</th></tr><tr>$InputSearch</tr>"
	$Data = $Data -replace "</colgroup>", "</colgroup> <thead>" -replace "placeholder='Search...'></th></tr>", "placeholder='Search...'></th></tr></thead> <tbody>" -replace "</table>", "</tbody></table>"
	$Data = $Data -replace "<tr><td>error</td>","<tr class='Error'>" -replace "<tr><td>warning</td>","<tr class='Warning'>" -replace "<tr><td>ok</td>","<tr>"
	
	If (!$HideStatusText) {ForEach ($key in $ErrorsHash.Keys) { $Data = $Data -replace ">$key", ">$($ErrorsHash.$key) $key" -replace ",$key", " $($ErrorsHash.$key) $key" } }
	Else { ForEach ($key in $ErrorsHash.Keys) { $Data = $Data -replace ">$key", ">$($ErrorsHash.$key)" -replace ",$key", " $($ErrorsHash.$key)" } }
	
	If ($ShowLegend) {	
		If ($Thresholds.count) {
			$comp = "&gt;"
			If ($Watermark -eq "Below") { $comp = "&lt;" }
			$Data = $Data -replace "$Title`n<table$TableWidth class='Table-Regular' id='$ID'>","<div style='text-align: right;'><span style='font-weight: bold;'>Thresholds:</span> $Warning_img_s Warning  $comp<span style='font-weight: bold; color:orange'> $($Thresholds[0])</span> , $Error_img_s Error $comp<span style='font-weight: bold; color:red'> $($Thresholds[1])</span> </div>$Title`n<table$TableWidth class='Table-Regular' id='$ID'>"
			}
		Else {
			$css_code = ".Legend { text-align: center; width:1%; border-color: #00000; border-collapse: collapse; margin-right: 7px; white-space: nowrap; }
.Legend th { font-weight: bold; background-color: $MainColor; color: #00000; padding: 5px; border: 1px solid black;  }
.Legend td { padding: 0px;  border: 1px solid black;}"
			PAF-HTML-Headers -CSS -Code $css_code
			
			$t = "<table class='Legend'><tr><th>Status message</th><th>Status icon</th></tr>"
			ForEach ($key in $ErrorsHash.Keys) { 
				$t += "<tr><td>$key</td><td>$($ErrorsHash.$Key)</td></tr>" 
				}
			$t += "</table>"
			$Data = $Data -replace "$Title`n<table$TableWidth class='Table-Regular' id='$ID'>","<div class='Opposite'><span></span>$t</div>$Title`n<table$TableWidth class='Table-Regular' id='$ID'>"
			}
		}

	}


Else {
	$Result_Table = "<table><tr><td>"
	$Result_Table += $Data
	$Result_Table += "</td></tr></table>"
	$Data = $Result_Table
	$Data = $Data -replace "&lt;", "<" -replace "&gt;", ">" -replace "&#39;", "'"
	If (!$HideStatusText) { ForEach ($key in $ErrorsHash.Keys) { $Data = $Data -replace ">$key", ">$($ErrorsHash.$key) $key" -replace " $key|,$key", " $($ErrorsHash.$key) $key" } }
	Else { ForEach ($key in $ErrorsHash.Keys) { $Data = $Data -replace ">$key", ">$($ErrorsHash.$key)" -replace " $key|,$key", " $($ErrorsHash.$key)" } }
	}



Return $Data
}

Function PAF-HTML-Buttons-Horizontal {
	param (
		[Parameter(Mandatory=$true)][PSCustomObject[]] $Data,
		[Parameter(Mandatory=$false)][string] $Title,
		[Parameter(Mandatory=$false)][Switch] $ColoredTitle,
		[Parameter(Mandatory=$false)][int] $Size = 250
		)

$css_code = ".button { font-family: Verdana; border: 3px solid $MainColor; background-color: #FFFFFF; color: $MainColor; padding: 2px; text-align: center; text-decoration: none; display: inline-block; font-size: 12pt; cursor: pointer; font-weight: bold; width: $($Size)px; }
.active, .button:hover { background-color: $MainColor; color: white; }
.dataSource { text-align:center; display: none; }"
PAF-HTML-Headers -CSS -Code $css_code

$js_code = "function showButton(id,tables) {
var tables = eval(tables)
var button = document.getElementById(id)
var buttonsArray = (button.parentElement.parentElement.parentElement).getElementsByTagName('button');
ids = []
for (i = 0; i < buttonsArray.length; i++) { ids.push(buttonsArray[i]) }
for (i = 0; i < buttonsArray.length; i++) { buttonsArray[i].className = buttonsArray[i].className.replace(' active', '') }
button.className += ' active';
for (i = 0; i < tables.length; i++) { document.getElementById(tables[i]).style.display = 'none'  }
document.getElementById(tables[ids.indexOf(button)]).style.display = 'table'
}"
PAF-HTML-Headers -JS -Code $js_code

$tablesArr = ([guid]::NewGuid()).Guid -replace "-" -replace "^[0-9\s]+"
$TableIDs = @()
$ButtonIDs = @()
$tables = ""

$HTML = "<table style='text-align:center'>`n<tr><td>"
$width = [int](100/$Data.count)

If ($Title) {
	$style = ""
	If ($ColoredTitle) { $style  = " style='color: $MainColor;'"}
	$HTML += "`n<h2$style>$Title</h2>"
	}

For ($i = 0; $i -lt $Data.count; $i++) { 
	$ButtonIDs += ([guid]::NewGuid()).Guid -replace "-"
	$TableIDs += ([guid]::NewGuid()).Guid -replace "-"
	If ($Data[$i].Table.Type -like "PAF-*") { $table = PAF-HTML-Render -HTMLObject $Data[$i].Table }
	Else { $table = $Data[$i].Table }
	$tables += "`n<table class='dataSource' id='$($TableIDs[$i])'>`n<tr><td>$table</td></tr>`n</table>"
}

$HTML += "`n<table><tr>"
For ($i = 0; $i -lt $Data.count; $i++) {
	$HTML += "`n<td width='$width%'><button id='$($ButtonIDs[$i])' class='button' onClick=""showButton('$($ButtonIDs[$i])','$tablesArr')"">$($Data[$i].ButtonText)</button></td>"
	}

Return "$HTML
</tr></table>`n<hr>`n</td></tr>`n<tr><td>
$tables`n</td></tr>`n</table>`n<script>
var $tablesArr = ['$($TableIDs -join ""','"")']
document.getElementById('$($ButtonIDs[0])').className += ' active';
document.getElementById($tablesArr[0]).style.display = 'table'
</script>"
}

Function PAF-HTML-Buttons-Vertical {
	param (
		[Parameter(Mandatory=$true)][PSCustomObject[]] $Data,
		[Parameter(Mandatory=$false)][string] $Title,
		[Parameter(Mandatory=$false)][Switch] $ColoredTitle,
		[Parameter(Mandatory=$false)][int] $Size = 250
		)

$css_code = ".button { font-family: Verdana; border: 3px solid $MainColor; background-color: #FFFFFF; color: $MainColor; padding: 2px; text-align: center; text-decoration: none; display: inline-block; font-size: 12pt; cursor: pointer; font-weight: bold; width: $($Size)px; }
.active, .button:hover { background-color: $MainColor; color: white; }
.dataSource { text-align:center; display: none; }"
PAF-HTML-Headers -CSS -Code $css_code

$js_code = "function showButton(id,tables) {
var tables = eval(tables)
var button = document.getElementById(id)
var buttonsArray = (button.parentElement.parentElement.parentElement).getElementsByTagName('button');
ids = []
for (i = 0; i < buttonsArray.length; i++) { ids.push(buttonsArray[i]) }
for (i = 0; i < buttonsArray.length; i++) { buttonsArray[i].className = buttonsArray[i].className.replace(' active', '') }
button.className += ' active';
for (i = 0; i < tables.length; i++) { document.getElementById(tables[i]).style.display = 'none'  }
document.getElementById(tables[ids.indexOf(button)]).style.display = 'table'
}"
PAF-HTML-Headers -JS -Code $js_code

$tablesArr = ([guid]::NewGuid()).Guid -replace "-" -replace "^[0-9\s]+"
$TableIDs = @()
$ButtonIDs = @()
$tables = ""

$HTML = "<table style='text-align:center'>"
If ($Title) { 
	$style = ""
	If ($ColoredTitle) { $style  = " style='color: $MainColor;'"}
	$HTML += "`n<tr><td colspan='2'><h2$style>$Title</h2></td></tr>"
	}
$rowspan = "rowspan='$($Data.count)'"

For ($i = 0; $i -lt $Data.count; $i++) { 
	$ButtonIDs += ([guid]::NewGuid()).Guid -replace "-"
	$TableIDs += ([guid]::NewGuid()).Guid -replace "-"
	If ($Data[$i].Table.Type -like "PAF-*") { $table = PAF-HTML-Render -HTMLObject $Data[$i].Table }
	Else { $table = $Data[$i].Table }
	$tables += "`n<table class='dataSource' id='$($TableIDs[$i])'>`n<tr><td>$table</td></tr>`n</table>"
	}

For ($i = 0; $i -lt $Data.count; $i++) {
	If ($i -eq 0) { $HTML += "`n<tr style='text-align:left'><td><button id='$($ButtonIDs[$i])' class='button' onClick=""showButton('$($ButtonIDs[$i])','$tablesArr')"">$($Data[$i].ButtonText)</button></td>`n<td $rowspan style='width:100%'>$tables</td></tr>" }
	Else { $HTML += "`n<tr style='text-align:left'><td><button id='$($ButtonIDs[$i])' class='button' onClick=""showButton('$($ButtonIDs[$i])','$tablesArr')"">$($Data[$i].ButtonText)</button></td></tr>" } 
	}

Return "$HTML`n</table>`n<script>
var $tablesArr = ['$($TableIDs -join ""','"")']
document.getElementById('$($ButtonIDs[0])').className += ' active';
document.getElementById($tablesArr[0]).style.display = 'table'
</script>"
}

####RENDER HTM Objects
Function PAF-HTML-Render {
	param ( [PAFHTML[]] $HTMLObject )
	
$HTML = ""
ForEach ($obj in $HTMLObject) {
	$Data = $obj.Params.data
    $Params = $obj.Params
    $keys = @($Params.Keys)
    $keys | % { If ($_ -eq "Data") { $Params.Remove($_) } }

	If ($Data.Type -like "PAF-*") { $HTML += &$($obj.Type) -Data $(PAF-HTML-Render -HTMLObject $Data) @Params  }
	Else { 
        If ($data) { $HTML += &$($obj.Type) -Data $Data @Params }
		Else { $HTML += &$($obj.Type) @Params }
		}
	}
Return $HTML
}
