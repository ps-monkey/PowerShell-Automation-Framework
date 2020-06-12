param ($IP)
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

Function Test-Port {
	param (
		$Address,
		$Port
		)
$TCPtimeout = 1000
$TCPObject = New-Object system.Net.Sockets.TcpClient
$connect = $TCPObject.BeginConnect($Address,$Port,$null,$null)
$wait = $connect.AsyncWaitHandle.WaitOne($TCPtimeout,$false)
If (!$wait) {  
	$TCPObject.Close()
	Return $false
	}
Else {
	$error.Clear()
	Try { $TCPObject.EndConnect($connect) | out-null } Catch { Return $false }
	If ($error[0]) { Return $false }
	$TCPObject.Close()
	Return $true
	}
}

Function AAFEnv-DetectSystemType {
	param ([string]$IP)

$Types = New-Object System.Collections.ArrayList
$HTTPS_query = Try { (Invoke-WebRequest -Uri "https://$($IP)" -TimeoutSec 2).Content } Catch {}
#Exclude EMC ESRS
If (Test-Port -Address $IP -Port 9443) {
	Try { If ((Invoke-WebRequest -Uri "https://$($IP):9443" -TimeoutSec 2).Content -match "esrs") { Return $false} } Catch {}
	}
		
#Exclude VCE Vision
If ($HTTPS_query -cmatch "VCE Vision") { Return $false }

Try { If ((Invoke-WebRequest -Uri "https://$($IP)/vsphere-client" -TimeoutSec 2).StatusCode -eq 200) { Return "vCenter" } } Catch {}
If ($HTTPS_query -cmatch "NSX") { Return "NSX" }
If ($(Test-Port -Address $IP -Port 8281) -and $(Test-Port -Address $IP -Port 8283)) { Return "vRO" }
If ($(Test-Port -Address $IP -Port 443) -and $(Test-Port -Address $IP -Port 3389) -and $(Test-Path $("filesystem::\\$($IP)\CMFiles$"))) { Return "VCM" }
Try {
	If (Test-Port -Address $IP -Port 7080) {
		If ((Invoke-WebRequest -Uri "http://$($IP):7080" -TimeoutSec 2).Content -cmatch "Hyperic") { Return "Hyperic" }
		}
	}
Catch {}

If ($HTTPS_query -cmatch "vRealize Operations Manager") { Return "vROPs" }
If ($HTTPS_query -cmatch "vCloud Director") { Return "vCD" }
If ($HTTPS_query -cmatch "Cisco UCS Manager") { Return "UCS-Fabric" }
If ($HTTPS_query -cmatch "Cisco Integrated Management") { Return "UCS-Standalone" }

If (Test-Port -Address $IP -Port 443) {
	Try { If ((Invoke-RestMethod -Uri "https://$($IP)/redfish/v1" -TimeoutSec 2).AccountService.'@odata.id' -match "idrac") { Return "Dell" } } Catch {}
}

#bull
#bull-sequana

Try { If ((Invoke-WebRequest -Uri "https://$($IP)/smsflex/VPlexConsole.html" -TimeoutSec 2).StatusCode -eq 200) { Return "VPLEX" } } Catch {}
If ($HTTPS_query -cmatch "window.open.*start.html") { Return "VNX" }
If ($HTTPS_query -cmatch "EMC Unisphere") { Return "Unity" }

#"JuniperSwitch"
Try { If ((Invoke-WebRequest -Uri "https://$($IP)/ddem/login" -TimeoutSec 2).StatusCode -eq 200) { Return "DataDomain" } } Catch {}
If ($HTTPS_query -cmatch "EMC Avamar") { Return "Avamar" }
If (Test-Port -Address $IP -Port 3389) { Return "WindowsServer" }
If (Test-Port -Address $IP -Port 8443) { 
	Try { If ((Invoke-WebRequest -Uri "https://$($IP):8443/core/orionSplashScreen.do" -TimeoutSec 2).StatusCode -eq 200) { Return "ePO" } } Catch {}
	} 

If (Test-Port -Address $IP -Port 22) { 
	$Types.Add("LinuxServer") | out-null 
	$Types.Add("CiscoMDS") | out-null 
	$Types.Add("CiscoNexus") | out-null 
	}
#If (Test-Port -Address $IP -Port 22) { $Types.Add("CiscoMDS")  | out-null }
#If (Test-Port -Address $IP -Port 22) { $Types.Add("CiscoNexus")  | out-null }

If (!$Types) { Return $false }
Return $Types
}

$DiscoveredSystemTypes = New-Object System.Collections.ArrayList

If (Test-Connection -Computername $IP -Count 1 -Quiet) {
	$DNSNAme = Try { ([System.Net.Dns]::gethostentry($IP)).HostName } Catch {""}
	$DiscoveredSystemTypes = AAFEnv-DetectSystemType -IP $IP
	$DiscoveredSystem = [pscustomobject][ordered]@{'IP'= $IP; 'DNS'= $DNSNAme; 'System Type' = $DiscoveredSystemTypes}
	If ($DiscoveredSystem.'System Type') { $DiscoveredSystem }
	} 

