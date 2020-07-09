#################################################################################################################################################################################################
#Check if run as administrator
If (!(([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
	Write-host "Must be run with administrative permissions.`nPlease use ""Run As Administrator"" context menu to start PowerShell" -foregroundcolor "red"
	Return "Error"
}

#Clear $error variable
$error.Clear()

#Enable TLS12
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Define Trust and Accept all certificates policy
If ( -not ("TrustAllCertsPolicy" -as [type])) {
Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
	public bool CheckValidationResult(
	ServicePoint srvPoint, X509Certificate certificate,
	WebRequest request, int certificateProblem) {
	return true;
	}
}
"@
}
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'

#Decrypt function
Function Decrypt {
	param ( [Parameter(Mandatory=$true)][string] $Path )
If (Get-Content -Path $Path) { Return Get-Content -Path $Path | Unprotect-CmsMessage | ConvertFrom-Json }
Else { Return "" }
}

#Define modules location
$PAFModulesFolder = "$PAFFolder\Modules"

#Load status images
. $("$PAFModulesFolder\PAF\Images.ps1") -ea "Stop"

#Load configuration
Try { $global:Config = Decrypt -Path "config.pafc" }
Catch { Write-host "Configuration cannot be loaded!" -foregroundcolor "red" }

$DefaultProperties = Get-Content -Path "$PAFFolder\defaults.pafp" | ConvertFrom-Json

#Load Properties
#load html styles
If (!$global:Config.Properties.style_html) { $global:Config.Properties | Add-Member noteproperty "style_html" -value $DefaultProperties.style_html }

#load Modules list
$Modules = $global:Config.Properties.Modules

#Load environment
$global:environment = If ($global:Config.envFileLocation) { Decrypt -Path $global:Config.envFileLocation } Else { "" }

#Define resources from the environment
$Resources = ($global:environment.PSObject.Properties | ? {($_.Value).count -gt 1}).Name

#Add secured passwords to environment
ForEach ($Resource in $Resources) {
	ForEach ($System in $global:environment.$Resource | ? {!$_.Label}) {
		#Add secure credentials
		If ($System.Password) {
			$SecurePassword = ConvertTo-SecureString $System.Password -AsPlainText -Force
			$System | Add-Member noteproperty "SecurePassword" -value $SecurePassword
			$Credentials = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $System.UserName, $SecurePassword
			$System | Add-Member noteproperty "Credentials" -value $Credentials
			}
		}
	}

#Create Custom objects
$global:CustomerObj = $global:Config.Customer
$EmailObj = $global:Config.Mail

#Load default modules
Write-host "Loading default modules" -foregroundcolor "yellow"
Try { If (-Not (Get-Module).Name.Contains("ActiveDirectory")) { Import-Module "ActiveDirectory" -ea "Stop" } }
Catch { Write-host "Cannot load Active Directory modules"  -foregroundcolor "red" }

#Set Report file name and output path
$ReportFileName = $($global:CustomerObj.'NameTemplate' -creplace "%HH%", "$(Get-Date -format HH)" -creplace "%hh%", "$(Get-Date -format hh)" -creplace "%mm%", "$(Get-Date -format mm)" -creplace "%ss%", "$(Get-Date -format ss)" -creplace "%dd%", "$(Get-Date -format dd)" -creplace "%MM%", "$(Get-Date -format MM)" -creplace "%MMMM%", "$(Get-Date -format MMMM)" -creplace "%yy%", "$(Get-Date -format yy)" -creplace "%yyyy%", "$(Get-Date -format yyyy)")
$ReportFileFolder = $global:Config.ReportLocation -creplace "%HH%", "$(Get-Date -format HH)" -creplace "%hh%", "$(Get-Date -format hh)" -creplace "%mm%", "$(Get-Date -format mm)" -creplace "%ss%", "$(Get-Date -format ss)" -creplace "%dd%", "$(Get-Date -format dd)" -creplace "%MM%", "$(Get-Date -format MM)" -creplace "%MMMM%", "$(Get-Date -format MMMM)" -creplace "%yy%", "$(Get-Date -format yy)" -creplace "%yyyy%", "$(Get-Date -format yyyy)"
$ReportFilePath = $ReportFileFolder + "\" + $ReportFileName
 
#Define default variables
$UserName = $env:UserName
$LogonName = $env:UserDomain + "\" + $UserName
$UserFullName = Try { (Get-ADUser -Filter {SamAccountName -eq $UserName}).name } Catch {"Unknown"}

Write-host "Loading modules for defined in configuration" -foregroundcolor "yellow"
#Load HTML module if filetype is HTM(L)
If ($ReportFileName -match "\.html$|\.htm$") {
	[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
	[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms.DataVisualization')
	If (-Not (Get-Module).Name.Contains("paf-html")) { Import-Module $($PAFModulesFolder + "\PAF\paf-html.psm1") -WarningAction SilentlyContinue -ea "Stop" }
	}

#Load PSExcel module if filetype is XLSX
If ($ReportFileName -match "\.xlsx$") {
	If (-Not (Get-Module).Name.Contains("paf-xlsx")) { Import-Module $($PAFModulesFolder + "\PAF\paf-xlsx.psm1") -ea "Stop" -WarningAction SilentlyContinue }
	}

ForEach ($Module in $Modules) {
	Try { If (-Not (Get-Module).Name.Contains($Module.Path -replace ".*\\")) { Import-Module $Module.Path -ea "Stop" -WarningAction SilentlyContinue } }
	Catch { Write-host "Cannot load PS Module: $($Module.Path -replace '.*\\')"  -foregroundcolor "red" }
	}

Function PAF-SaveFile {
	param ( 
		[Parameter(Mandatory=$true)] $Content,
		[Parameter(Mandatory=$false)][string] $File = $ReportFilePath,
		[Parameter(Mandatory=$false)][string] $Delimiter = ",",
		[Parameter(Mandatory=$false)][Switch] $Raw
		)
$Path = Split-Path -Path $File
If (!(Test-Path $Path)) { New-Item -ItemType Directory -Force -Path $Path | out-null }

If (!$Raw) {
	Switch -regex ($File) {
		"\.xlsx$" {
            If ($Content) { PAF-XLSX-Create -XLSXObject $Content -File $File }
            break
            }
		"\.html$|\.htm$" {
			$global:CSS = ""
			$global:JS = ""
			If ($Content) {
                $HTMLBody = PAF-HTML-Render -HTMLObject $Content
			    PAF-HTML-Create -Body $HTMLBody | Out-File $File -Force;
                }
			break 
			}
		"\.csv$" { $Content | Export-Csv -NoTypeInformation -Delimiter $Delimiter -Force; break  }
		default { $Content | Out-File $File -Force ; break }
		}
	}
Else { $Content | Out-File $File -Force }

Remove-Module "paf-core" -ea SilentlyContinue  #Removed for demo purposes
Remove-Variable Config -Scope "Global" -ea SilentlyContinue
Remove-Variable CustomerObj -Scope "Global" -ea SilentlyContinue
Remove-Variable environment -Scope "Global" -ea SilentlyContinue
}

Function PAF-SendEmail {
	param ( 
		[Parameter(Mandatory=$false)][ValidateSet("SMTP Server","vRealize Orchestrator")][string] $Transport = $EmailObj.'Transport',
		[Parameter(Mandatory=$false)][string] $SmtpServer = $EmailObj.'SMTPServer',
		[Parameter(Mandatory=$false)][string] $vROSMTPServer = $EmailObj.'vROSMTPServer',
		[Parameter(Mandatory=$false)][string] $From = $EmailObj.'e-mailFrom',
		[Parameter(Mandatory=$false)][string] $To = $EmailObj.'e-mailTo',
		[Parameter(Mandatory=$false)][string] $Subject = $EmailObj.'Subject',
		[Parameter(Mandatory=$false)][string] $Body = $EmailObj.'BodyText',
		[Parameter(Mandatory=$false)][string] $Attachment = $ReportFilePath,
		[Parameter(Mandatory=$false)] $AttachReport = $EmailObj.'AttachReport',
		[Parameter(Mandatory=$false)][Switch] $AsPlainText
		)

If (!$Body) { $Body = " " }

If ($EmailObj.SendByEmail -eq "yes") { Write-host "Sending email..." -foregroundcolor "yellow" }

If ($Transport -eq "SMTP Server") {
	$MailSettings = @{
		SmtpServer = $SmtpServer
		From = $From
		To = $To.split(";").split(",")
		Subject = $Subject
		Body = $Body
		BodyAsHtml = If ($AsPlainText) { $false } Else { $true }
	}
	If ($AttachReport) { $MailSettings.Add("Attachments", $Attachment) }
	Send-MailMessage @MailSettings
	}
If ($Transport -eq "vRealize Orchestrator") {
	If ($AttachReport) {
		If (!$Attachment) { $Attachment = $ReportFilePath }
		$AttachmentFileName = Split-Path -Path $Attachment -Leaf
		
		Add-Type -Path $($PAFModulesFolder + "\WinSCP\WinSCPnet.dll")
		$Options = New-Object WinSCP.SessionOptions -Property @{
			Protocol = [WinSCP.Protocol]::Sftp
			HostName = $EmailObj.vROHostName
			UserName = $EmailObj.vROSSHUser
			Password = $EmailObj.vROSSHPassword
			GiveUpSecurityAndAcceptAnySshHostKey = $true
			}

		$TransferOptions = New-Object WinSCP.TransferOptions
		$TransferOptions.FilePermissions =  New-Object WinSCP.FilePermissions
		$TransferOptions.FilePermissions.Octal = "644"
		
		$RemotePath = "/etc/vco/$AttachmentFileName"
		$WinSCPSession = New-Object WinSCP.Session
		$WinSCPSession.DisableVersionCheck = $true
		$WinSCPSession.Open($Options)
		$WinSCPSession.PutFiles($Attachment, $RemotePath,$false,$TransferOptions) | out-null
		}
	
	If (-Not (Get-Module).Name.Contains("PowervRO")) { Import-Module ($PAFModulesFolder + "\PowervRO") -ea "Stop"}
	$SecurePassword = ConvertTo-SecureString $EmailObj.vROPassword -AsPlainText -Force
	Connect-vROServer -Server $EmailObj.vROHostName -Username $EmailObj.vROUser -Password $SecurePassword -IgnoreCertRequirements -SslProtocol "tls"  -ea "Stop" | out-null
	$id = (Get-vROWorkflow -Name 'Send email with attachment').ID
	If (!$id) {
		$path  = $PAFModulesFolder + "\PowervRO\email.workflow"
		$Category = (Get-vROCategory -CategoryType "WorkflowCategory" | ? {$_.Name -eq "Mail"})[0]
		Import-vROWorkflow -CategoryId $Category.ID -File $path -Confirm:$false
		Sleep 10
		$id = (Get-vROWorkflow -Name 'Send email with attachment').ID
		}
	$emails = $To -replace ";",","
	$smtpPort = "25"
	$Params = @()
	$Params += New-vROParameterDefinition -Name "fromAddress" -Value $From -Type String -Scope LOCAL
	$Params += New-vROParameterDefinition -Name "smtpHost" -Value $vROSMTPServer -Type String -Scope LOCAL
	$Params += New-vROParameterDefinition -Name "smtpPort" -Value $smtpPort -Type String -Scope LOCAL
	$Params += New-vROParameterDefinition -Name "subject" -Value $Subject -Type String -Scope LOCAL	
	$Params += New-vROParameterDefinition -Name "attachment" -Value $AttachmentFileName -Type String -Scope LOCAL
	$Params += New-vROParameterDefinition -Name "emailTo" -Value $emails -Type String -Scope LOCAL
	$Params += New-vROParameterDefinition -Name "body" -Value $Body -Type String -Scope LOCAL
	Sleep 5
	Invoke-vROWorkflow -id $id -Parameters $Params | out-null

	Disconnect-vROServer -Confirm:$false
	Sleep 30
	$WinSCPSession.RemoveFiles($RemotePath) | out-null
	$WinSCPSession.Dispose()
	}
}

Function PAF-Test-Port {
    param (
		[Parameter(Mandatory=$true)][string] $Address,
		[Parameter(Mandatory=$true)][int] $Port
		)
$tcpClient = New-Object Net.Sockets.TcpClient
Try { $tcpClient.Connect("$Address", $Port); $true }
Catch { $false }
Finally { $tcpClient.Dispose() }
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
	[Parameter(Mandatory=$true)][int]$Row,
	[Parameter(Mandatory=$true)][int]$Column
	)

$ColumnIndex = Get-ExcelColumn $Column
"$ColumnIndex$Row"
}
