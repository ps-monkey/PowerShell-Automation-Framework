Function Invoke-Parallel {
    <#
    .SYNOPSIS
        Function to control parallel processing using runspaces

    .DESCRIPTION
        Function to control parallel processing using runspaces

            Note that each runspace will not have access to variables and commands loaded in your session or in other runspaces by default.  
            This behaviour can be changed with parameters.

    .PARAMETER ScriptFile
        File to run against all input objects.  Must include parameter to take in the input object, or use $args.  Optionally, include parameter to take in parameter.  Example: C:\script.ps1

    .PARAMETER ScriptBlock
        Scriptblock to run against all computers.

        You may use $Using:<Variable> language in PowerShell 3 and later.
        
            The parameter block is added for you, allowing behaviour similar to foreach-object:
                Refer to the input object as $_.
                Refer to the parameter parameter as $parameter

    .PARAMETER InputObject
        Run script against these specified objects.

    .PARAMETER Parameter
        This object is passed to every script block.  You can use it to pass information to the script block; for example, the path to a logging folder
        
            Reference this object as $parameter if using the scriptblock parameterset.

    .PARAMETER ImportVariables
        If specified, get user session variables and add them to the initial session state

    .PARAMETER ImportModules
        If specified, get loaded modules and pssnapins, add them to the initial session state

    .PARAMETER Throttle
        Maximum number of threads to run at a single time.

    .PARAMETER SleepTimer
        Milliseconds to sleep after checking for completed runspaces and in a few other spots.  I would not recommend dropping below 200 or increasing above 500

    .PARAMETER RunspaceTimeout
        Maximum time in seconds a single thread can run.  If execution of your code takes longer than this, it is disposed.  Default: 0 (seconds)

        WARNING:  Using this parameter requires that maxQueue be set to throttle (it will be by default) for accurate timing.  Details here:
        http://gallery.technet.microsoft.com/Run-Parallel-Parallel-377fd430

    .PARAMETER NoCloseOnTimeout
		Do not dispose of timed out tasks or attempt to close the runspace if threads have timed out. This will prevent the script from hanging in certain situations where threads become non-responsive, at the expense of leaking memory within the PowerShell host.

    .PARAMETER MaxQueue
        Maximum number of powershell instances to add to runspace pool.  If this is higher than $throttle, $timeout will be inaccurate
        
        If this is equal or less than throttle, there will be a performance impact

        The default value is $throttle times 3, if $runspaceTimeout is not specified
        The default value is $throttle, if $runspaceTimeout is specified

    .PARAMETER LogFile
        Path to a file where we can log results, including run time for each thread, whether it completes, completes with errors, or times out.

	.PARAMETER Quiet
		Disable progress bar.

    .EXAMPLE
        Each example uses Test-ForPacs.ps1 which includes the following code:
            param($computer)

            if(test-connection $computer -count 1 -quiet -BufferSize 16){
                $object = [pscustomobject] @{
                    Computer=$computer;
                    Available=1;
                    Kodak=$(
                        if((test-path "\\$computer\c$\users\public\desktop\Kodak Direct View Pacs.url") -or (test-path "\\$computer\c$\documents and settings\all users

        \desktop\Kodak Direct View Pacs.url") ){"1"}else{"0"}
                    )
                }
            }
            else{
                $object = [pscustomobject] @{
                    Computer=$computer;
                    Available=0;
                    Kodak="NA"
                }
            }

            $object

    .EXAMPLE
        Invoke-Parallel -scriptfile C:\public\Test-ForPacs.ps1 -inputobject $(get-content C:\pcs.txt) -runspaceTimeout 10 -throttle 10

            Pulls list of PCs from C:\pcs.txt,
            Runs Test-ForPacs against each
            If any query takes longer than 10 seconds, it is disposed
            Only run 10 threads at a time

    .EXAMPLE
        Invoke-Parallel -scriptfile C:\public\Test-ForPacs.ps1 -inputobject c-is-ts-91, c-is-ts-95

            Runs against c-is-ts-91, c-is-ts-95 (-computername)
            Runs Test-ForPacs against each

    .EXAMPLE
        $stuff = [pscustomobject] @{
            ContentFile = "windows\system32\drivers\etc\hosts"
            Logfile = "C:\temp\log.txt"
        }
    
        $computers | Invoke-Parallel -parameter $stuff {
            $contentFile = join-path "\\$_\c$" $parameter.contentfile
            Get-Content $contentFile |
                set-content $parameter.logfile
        }

        This example uses the parameter argument.  This parameter is a single object.  To pass multiple items into the script block, we create a custom object (using a PowerShell v3 language) with properties we want to pass in.

        Inside the script block, $parameter is used to reference this parameter object.  This example sets a content file, gets content from that file, and sets it to a predefined log file.

    .EXAMPLE
        $test = 5
        1..2 | Invoke-Parallel -ImportVariables {$_ * $test}

        Add variables from the current session to the session state.  Without -ImportVariables $Test would not be accessible

    .EXAMPLE
        $test = 5
        1..2 | Invoke-Parallel -ImportVariables {$_ * $Using:test}

        Reference a variable from the current session with the $Using:<Variable> syntax.  Requires PowerShell 3 or later.

    .FUNCTIONALITY
        PowerShell Language

    .NOTES
        Credit to Boe Prox for the base runspace code and $Using implementation
            http://learn-powershell.net/2012/05/10/speedy-network-information-query-using-powershell/
            http://gallery.technet.microsoft.com/scriptcenter/Speedy-Network-Information-5b1406fb#content
            https://github.com/proxb/PoshRSJob/

        Credit to T Bryce Yehl for the Quiet and NoCloseOnTimeout implementations

        Credit to Sergei Vorobev for the many ideas and contributions that have improved functionality, reliability, and ease of use

    .LINK
        https://github.com/RamblingCookieMonster/Invoke-Parallel
    #>
    [cmdletbinding(DefaultParameterSetName='ScriptBlock')]
    Param (   
        [Parameter(Mandatory=$false,position=0,ParameterSetName='ScriptBlock')]
            [System.Management.Automation.ScriptBlock]$ScriptBlock,

        [Parameter(Mandatory=$false,ParameterSetName='ScriptFile')]
        [ValidateScript({test-path $_ -pathtype leaf})]
            $ScriptFile,

        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Alias('CN','__Server','IPAddress','Server','ComputerName')]    
            [PSObject]$InputObject,
            [PSObject]$Parameter,
            [switch]$ImportVariables,
            [switch]$ImportModules,
            [int]$Throttle = 20,
            [int]$SleepTimer = 200,
            [int]$RunspaceTimeout = 0,
			[switch]$NoCloseOnTimeout = $false,
            [int]$MaxQueue,
        [validatescript({Test-Path (Split-Path $_ -parent)})]
            [string]$LogFile = "C:\temp\log.log",
			[switch] $Quiet = $false
    )
    
    Begin {
                
        #No max queue specified?  Estimate one.
        #We use the script scope to resolve an odd PowerShell 2 issue where MaxQueue isn't seen later in the function
        if( -not $PSBoundParameters.ContainsKey('MaxQueue') )
        {
            if($RunspaceTimeout -ne 0){ $script:MaxQueue = $Throttle }
            else{ $script:MaxQueue = $Throttle * 3 }
        }
        else
        {
            $script:MaxQueue = $MaxQueue
        }

        #Write-Verbose "Throttle: '$throttle' SleepTimer '$sleepTimer' runSpaceTimeout '$runspaceTimeout' maxQueue '$maxQueue' logFile '$logFile'"

        #If they want to import variables or modules, create a clean runspace, get loaded items, use those to exclude items
        if ($ImportVariables -or $ImportModules)
        {
            $StandardUserEnv = [powershell]::Create().addscript({

                #Get modules and snapins in this clean runspace
                $Modules = Get-Module | Select -ExpandProperty Name
                $Snapins = Get-PSSnapin | Select -ExpandProperty Name

                #Get variables in this clean runspace
                #Called last to get vars like $? into session
                $Variables = Get-Variable | Select -ExpandProperty Name
                
                #Return a hashtable where we can access each.
                @{
                    Variables = $Variables
                    Modules = $Modules
                    Snapins = $Snapins
                }
            }).invoke()[0]
            
            if ($ImportVariables) {
                #Exclude common parameters, bound parameters, and automatic variables
                Function _temp {[cmdletbinding()] param() }
                $VariablesToExclude = @( (Get-Command _temp | Select -ExpandProperty parameters).Keys + $PSBoundParameters.Keys + $StandardUserEnv.Variables )
                #Write-Verbose "Excluding variables $( ($VariablesToExclude | sort ) -join ", ")"

                # we don't use 'Get-Variable -Exclude', because it uses regexps. 
                # One of the veriables that we pass is '$?'. 
                # There could be other variables with such problems.
                # Scope 2 required if we move to a real module
                $UserVariables = @( Get-Variable | Where { -not ($VariablesToExclude -contains $_.Name) } ) 
                #Write-Verbose "Found variables to import: $( ($UserVariables | Select -expandproperty Name | Sort ) -join ", " | Out-String).`n"

            }

            if ($ImportModules) 
            {
                $UserModules = @( Get-Module | Where {$StandardUserEnv.Modules -notcontains $_.Name -and (Test-Path $_.Path -ErrorAction SilentlyContinue)} | Select -ExpandProperty Path )
                $UserSnapins = @( Get-PSSnapin | Select -ExpandProperty Name | Where {$StandardUserEnv.Snapins -notcontains $_ } ) 
            }
        }

        #region functions
            
            Function Get-RunspaceData {
                [cmdletbinding()]
                param( [switch]$Wait )

                #loop through runspaces
                #if $wait is specified, keep looping until all complete
                Do {

                    #set more to false for tracking completion
                    $more = $false

                    #Progress bar if we have inputobject count (bound parameter)
                    #if (-not $Quiet) {
					#	Write-Progress  -Activity "Running Query" -Status "Starting threads"`
					#		-CurrentOperation "$startedCount threads defined - $totalCount input objects - $script:completedCount input objects processed"`
					#		-PercentComplete $( Try { $script:completedCount / $totalCount * 100 } Catch {0} )
					#}
					$ProgressBar.Value = $( Try { $script:completedCount / $totalCount * 100 } Catch {0} )

                    #run through each runspace.           
                    Foreach($runspace in $runspaces) {
                    
                        #get the duration - inaccurate
                        $currentdate = Get-Date
                        $runtime = $currentdate - $runspace.startTime
                        $runMin = [math]::Round( $runtime.totalminutes ,2 )

                        #set up log object
                        $log = "" | select Date, Action, Runtime, Status, Details
                        $log.Action = "Removing:'$($runspace.object)'"
                        $log.Date = $currentdate
                        $log.Runtime = "$runMin minutes"

                        #If runspace completed, end invoke, dispose, recycle, counter++
                        If ($runspace.Runspace.isCompleted) {
                            
                            $script:completedCount++
                        
                            #check if there were errors
                            if($runspace.powershell.Streams.Error.Count -gt 0) {
                                
                                #set the logging info and move the file to completed
                                $log.status = "CompletedWithErrors"
                                #Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                                foreach($ErrorRecord in $runspace.powershell.Streams.Error) {
                                    Write-Error -ErrorRecord $ErrorRecord
                                }
                            }
                            else {
                                
                                #add logging details and cleanup
                                $log.status = "Completed"
                                #Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                            }

                            #everything is logged, clean up the runspace
                            $runspace.powershell.EndInvoke($runspace.Runspace)
                            $runspace.powershell.dispose()
                            $runspace.Runspace = $null
                            $runspace.powershell = $null

                        }

                        #If runtime exceeds max, dispose the runspace
                        ElseIf ( $runspaceTimeout -ne 0 -and $runtime.totalseconds -gt $runspaceTimeout) {
                            
                            $script:completedCount++
                            $timedOutTasks = $true
                            
							#add logging details and cleanup
                            $log.status = "TimedOut"
                            #Write-Verbose ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1]
                            #Write-Error "Runspace timed out at $($runtime.totalseconds) seconds for the object:`n$($runspace.object | out-string)"

                            #Depending on how it hangs, we could still get stuck here as dispose calls a synchronous method on the powershell instance
                            if (!$noCloseOnTimeout) { $runspace.powershell.dispose() }
                            $runspace.Runspace = $null
                            $runspace.powershell = $null
                            $completedCount++

                        }
                   
                        #If runspace isn't null set more to true  
                        ElseIf ($runspace.Runspace -ne $null ) {
                            $log = $null
                            $more = $true
                        }

                        #log the results if a log file was indicated
                        #if($logFile -and $log){
                        #    ($log | ConvertTo-Csv -Delimiter ";" -NoTypeInformation)[1] | out-file $LogFile -append
                        #}
                    }

                    #Clean out unused runspace jobs
                    $temphash = $runspaces.clone()
                    $temphash | Where { $_.runspace -eq $Null } | ForEach {
                        $Runspaces.remove($_)
                    }

                    #sleep for a bit if we will loop again
                    if($PSBoundParameters['Wait']){ Start-Sleep -milliseconds $SleepTimer }

                #Loop again only if -wait parameter and there are more runspaces to process
                } while ($more -and $PSBoundParameters['Wait'])
                
            #End of runspace function
            }

        #endregion functions
        
        #region Init

            if($PSCmdlet.ParameterSetName -eq 'ScriptFile')
            {
                $ScriptBlock = [scriptblock]::Create( $(Get-Content $ScriptFile | out-string) )
            }
            elseif($PSCmdlet.ParameterSetName -eq 'ScriptBlock')
            {
                #Start building parameter names for the param block
                [string[]]$ParamsToAdd = '$_'
                if( $PSBoundParameters.ContainsKey('Parameter') )
                {
                    $ParamsToAdd += '$Parameter'
                }

                $UsingVariableData = $Null
                

                # This code enables $Using support through the AST.
                # This is entirely from  Boe Prox, and his https://github.com/proxb/PoshRSJob module; all credit to Boe!
                
                if($PSVersionTable.PSVersion.Major -gt 2)
                {
                    #Extract using references
                    $UsingVariables = $ScriptBlock.ast.FindAll({$args[0] -is [System.Management.Automation.Language.UsingExpressionAst]},$True)    

                    If ($UsingVariables)
                    {
                        $List = New-Object 'System.Collections.Generic.List`1[System.Management.Automation.Language.VariableExpressionAst]'
                        ForEach ($Ast in $UsingVariables)
                        {
                            [void]$list.Add($Ast.SubExpression)
                        }

                        $UsingVar = $UsingVariables | Group SubExpression | ForEach {$_.Group | Select -First 1}
        
                        #Extract the name, value, and create replacements for each
                        $UsingVariableData = ForEach ($Var in $UsingVar) {
                            Try
                            {
                                $Value = Get-Variable -Name $Var.SubExpression.VariablePath.UserPath -ErrorAction Stop
                                [pscustomobject]@{
                                    Name = $Var.SubExpression.Extent.Text
                                    Value = $Value.Value
                                    NewName = ('$__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
                                    NewVarName = ('__using_{0}' -f $Var.SubExpression.VariablePath.UserPath)
                                }
                            }
                            Catch
                            {
                                Write-Error "$($Var.SubExpression.Extent.Text) is not a valid Using: variable!"
                            }
                        }
                        $ParamsToAdd += $UsingVariableData | Select -ExpandProperty NewName -Unique

                        $NewParams = $UsingVariableData.NewName -join ', '
                        $Tuple = [Tuple]::Create($list, $NewParams)
                        $bindingFlags = [Reflection.BindingFlags]"Default,NonPublic,Instance"
                        $GetWithInputHandlingForInvokeCommandImpl = ($ScriptBlock.ast.gettype().GetMethod('GetWithInputHandlingForInvokeCommandImpl',$bindingFlags))
        
                        $StringScriptBlock = $GetWithInputHandlingForInvokeCommandImpl.Invoke($ScriptBlock.ast,@($Tuple))

                        $ScriptBlock = [scriptblock]::Create($StringScriptBlock)

                        Write-Verbose $StringScriptBlock
                    }
                }
                
                $ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock("param($($ParamsToAdd -Join ", "))`r`n" + $Scriptblock.ToString())
            }
            else
            {
                Throw "Must provide ScriptBlock or ScriptFile"; Break
            }

            #Write-Debug "`$ScriptBlock: $($ScriptBlock | Out-String)"
            #Write-Verbose "Creating runspace pool and session states"

            #If specified, add variables and modules/snapins to session state
            $sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            if ($ImportVariables)
            {
                if($UserVariables.count -gt 0)
                {
                    foreach($Variable in $UserVariables)
                    {
                        $sessionstate.Variables.Add( (New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $Variable.Name, $Variable.Value, $null) )
                    }
                }
            }
            if ($ImportModules)
            {
                if($UserModules.count -gt 0)
                {
                    foreach($ModulePath in $UserModules)
                    {
                        $sessionstate.ImportPSModule($ModulePath)
                    }
                }
                if($UserSnapins.count -gt 0)
                {
                    foreach($PSSnapin in $UserSnapins)
                    {
                        [void]$sessionstate.ImportPSSnapIn($PSSnapin, [ref]$null)
                    }
                }
            }

            #Create runspace pool
            $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
            $runspacepool.Open() 

            #Write-Verbose "Creating empty collection to hold runspace jobs"
            $Script:runspaces = New-Object System.Collections.ArrayList        
        
            #If inputObject is bound get a total count and set bound to true
            $bound = $PSBoundParameters.keys -contains "InputObject"
            if(-not $bound)
            {
                [System.Collections.ArrayList]$allObjects = @()
            }

            #Set up log file if specified
            if( $LogFile ){
                New-Item -ItemType file -path $logFile -force | Out-Null
                ("" | Select Date, Action, Runtime, Status, Details | ConvertTo-Csv -NoTypeInformation -Delimiter ";")[0] | Out-File $LogFile
            }

            #write initial log entry
            $log = "" | Select Date, Action, Runtime, Status, Details
                $log.Date = Get-Date
                $log.Action = "Batch processing started"
                $log.Runtime = $null
                $log.Status = "Started"
                $log.Details = $null
                if($logFile) {
                    ($log | convertto-csv -Delimiter ";" -NoTypeInformation)[1] | Out-File $LogFile -Append
                }

			$timedOutTasks = $false

        #endregion INIT
    }

    Process {

        #add piped objects to all objects or set all objects to bound input object parameter
        if($bound)
        {
            $allObjects = $InputObject
        }
        Else
        {
            [void]$allObjects.add( $InputObject )
        }
    }

    End {
        
        #Use Try/Finally to catch Ctrl+C and clean up.
        Try
        {
            #counts for progress
            $totalCount = $allObjects.count
            $script:completedCount = 0
            $startedCount = 0

            foreach($object in $allObjects){
        
                #region add scripts to runspace pool
                    
                    #Create the powershell instance, set verbose if needed, supply the scriptblock and parameters
                    $powershell = [powershell]::Create()
                    
                    if ($VerbosePreference -eq 'Continue')
                    {
                        [void]$PowerShell.AddScript({$VerbosePreference = 'Continue'})
                    }

                    [void]$PowerShell.AddScript($ScriptBlock).AddArgument($object)

                    if ($parameter)
                    {
                        [void]$PowerShell.AddArgument($parameter)
                    }

                    # $Using support from Boe Prox
                    if ($UsingVariableData)
                    {
                        Foreach($UsingVariable in $UsingVariableData) {
                            #Write-Verbose "Adding $($UsingVariable.Name) with value: $($UsingVariable.Value)"
                            [void]$PowerShell.AddArgument($UsingVariable.Value)
                        }
                    }

                    #Add the runspace into the powershell instance
                    $powershell.RunspacePool = $runspacepool
    
                    #Create a temporary collection for each runspace
                    $temp = "" | Select-Object PowerShell, StartTime, object, Runspace
                    $temp.PowerShell = $powershell
                    $temp.StartTime = Get-Date
                    $temp.object = $object
    
                    #Save the handle output when calling BeginInvoke() that will be used later to end the runspace
                    $temp.Runspace = $powershell.BeginInvoke()
                    $startedCount++

                    #Add the temp tracking info to $runspaces collection
                    #Write-Verbose ( "Adding {0} to collection at {1}" -f $temp.object, $temp.starttime.tostring() )
                    $runspaces.Add($temp) | Out-Null
            
                    #loop through existing runspaces one time
                    Get-RunspaceData

                    #If we have more running than max queue (used to control timeout accuracy)
                    #Script scope resolves odd PowerShell 2 issue
                    $firstRun = $true
                    while ($runspaces.count -ge $Script:MaxQueue) {

                        #give verbose output
                        if($firstRun){
                            #Write-Verbose "$($runspaces.count) items running - exceeded $Script:MaxQueue limit."
                        }
                        $firstRun = $false
                    
                        #run get-runspace data and sleep for a short while
                        Get-RunspaceData
                        Start-Sleep -Milliseconds $sleepTimer
                    
                    }

                #endregion add scripts to runspace pool
            }
                     
            #Write-Verbose ( "Finish processing the remaining runspace jobs: {0}" -f ( @($runspaces | Where {$_.Runspace -ne $Null}).Count) )
            Get-RunspaceData -wait

            if (-not $quiet) {
			    #Write-Progress -Activity "Running Query" -Status "Starting threads" -Completed
		    }
        }
        Finally
        {
            #Close the runspace pool, unless we specified no close on timeout and something timed out
            if ( ($timedOutTasks -eq $false) -or ( ($timedOutTasks -eq $true) -and ($noCloseOnTimeout -eq $false) ) ) {
	            #Write-Verbose "Closing the runspace pool"
			    $runspacepool.close()
            }

            #collect garbage
            [gc]::Collect()
        }       
    }
}

Function Get-IPrange {
<# 
  .SYNOPSIS  
    Get the IP addresses in a range 
  .EXAMPLE 
   Get-IPrange -start 192.168.8.2 -end 192.168.8.20 
  .EXAMPLE 
   Get-IPrange -ip 192.168.8.2 -mask 255.255.255.0 
  .EXAMPLE 
   Get-IPrange -ip 192.168.8.3 -cidr 24 
#> 
 
param (
	[string] $start, 
	[string] $end, 
	[string] $ip, 
	[string] $mask, 
	[int] $cidr 
) 
 
Function IP-toINT64 () { 
	param ($ip)
$octets = $ip.split(".")
Return [int64]([int64]$octets[0]*16777216 +[int64]$octets[1]*65536 +[int64]$octets[2]*256 +[int64]$octets[3])
}
Function INT64-toIP() {
	param ([int64]$int)
Return (([math]::truncate($int/16777216)).tostring()+"."+([math]::truncate(($int%16777216)/65536)).tostring()+"."+([math]::truncate(($int%65536)/256)).tostring()+"."+([math]::truncate($int%256)).tostring() )
}

If ($ip) {$ipaddr = [Net.IPAddress]::Parse($ip)}
If ($cidr) {$maskaddr = [Net.IPAddress]::Parse((INT64-toIP -int ([convert]::ToInt64(("1"*$cidr+"0"*(32-$cidr)),2)))) }
If ($mask) {$maskaddr = [Net.IPAddress]::Parse($mask)}
If ($ip) {$networkaddr = new-object net.ipaddress ($maskaddr.address -band $ipaddr.address)}
If ($ip) {$broadcastaddr = new-object net.ipaddress (([system.net.ipaddress]::parse("255.255.255.255").address -bxor $maskaddr.address -bor $networkaddr.address))}

If ($ip) {
	$startaddr = IP-toINT64 -ip $networkaddr.ipaddresstostring
	$endaddr = IP-toINT64 -ip $broadcastaddr.ipaddresstostring
	}
Else {
	$startaddr = IP-toINT64 -ip $start
	$endaddr = IP-toINT64 -ip $end
	}
	
For ($i = $startaddr; $i -le $endaddr; $i++) { INT64-toIP -int $i }
}

#Environment management functions
Function PAFCI-BuildTree { 
ForEach ($CIGroup in $global:PAFDefaultConfig.CIConfig.CIGroup) { $CIGroupCombobox.Items.Add($CIGroup) | out-null }
$CIGroupCombobox.SelectedIndex = 0
ForEach ($System in $Systems) { PAFCI-AddObject -Object $System -Type "CI"}
}

Function PAFCI-AddObject {
	param (
		[Parameter(Mandatory=$true)] $Object,
		[Parameter(Mandatory=$true)] $Type
		)
Switch ($Type) {
	"CIGroup" {
		$oldFont = $label.font
		$BoldFont = New-Object Drawing.Font($oldFont.FontFamily, $oldFont.Size, [Drawing.FontStyle]::Bold)

		$NewNode = New-Object System.Windows.Forms.TreeNode 
		$NewNode.Text =  $Object + "     "
		$NewNode.Name = $Object
		$NewNode.Tag = "root"
		$NewNode | Add-Member noteproperty "Description" -value ($global:PAFDefaultConfig.CIConfig | ?{$_.CIGroup -eq $Object}).Description
		$NewNode.NodeFont = $BoldFont
		$treeView1.Nodes.Add($NewNode) | out-null
		}
	"CI" {
		$newNode = new-object System.Windows.Forms.TreeNode  
		$newNode.Name = $Object.Properties.Name
		$newNode.Text = $Object.Properties.Name
		$newNode.Tag = "CI"
		$newNode | Add-Member noteproperty "CIGroup" -value $Object.CIGroup
		$newNode | Add-Member noteproperty "Type" $Object.Type
		$newNode | Add-Member noteproperty "Properties" -value $Object.Properties
		If (!$($treeView1.Nodes | ? { $_.Name -eq $Object.CIGroup })) { PAFCI-AddObject -Object $Object.CIGroup -Type "CIGroup" }
		($treeView1.Nodes | ? { $_.Name -eq $Object.CIGroup }).Nodes.Add($newNode) | out-null
		#Return $newNode
		}
	}
PAFCI-FormUpdated
}

Function PAFCI-DeleteObject {
	param (
		[Parameter(Mandatory=$false)] $Object,
		[Parameter(Mandatory=$true)] $Type,
		[Parameter(Mandatory=$false)][Switch] $FromTree 
		)
Switch ($Type) {
	"Treeview" {
		If ($treeview1.SelectedNode.Tag -eq "CI") { PAFCI-DeleteObject -Type "CI" -FromTree}
		Else { PAFCI-DeleteObject -Object $treeview1.SelectedNode.Name -Type "CIGroup" -FromTree }
		$treeview1.Select()
		}
	"CIGroup" {
		If (!$Object) {$Object = $CIGroupCombobox.SelectedItem}
		If ($Systems) {
			$Systems = $Systems -ne ($Systems | ? {$_.CIGroup -eq $Object})
			$Node = $treeView1.Nodes | ? { $_.Name -eq $Object}
			Try { $treeView1.Nodes.Remove($Node) | out-null }
			Catch {}
			If (!$FromTree) { $CIGroupCombobox.Items.Remove($Object) }
			If ($Systems) { $CIGroupCombobox.SelectedIndex = 0 } 
			}
		Else {
			$CIGroupCombobox.Items.Clear()
			$CITypeCombobox.Items.Clear()
			$CIGroupCombobox.Text = ""
			$CITypeCombobox.Text = ""
			$DescriptionTextBox.Text = ""
			$CIProperties_GridView.Rows.Clear()
			}
		}
	"CIType" {
		$t = ($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Types
		$y = ($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Types | ? {$_.Type -eq $CITypeCombobox.SelectedItem}
		($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Types = ($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Types | ? {$_ -ne $(($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Types | ? {$_.Type -eq $CITypeCombobox.SelectedItem})}
		
		If ($CITypeCombobox.Items.count) {
			#$Nodes = $($treeView1.Nodes | ? { $_.Name -eq $CIGroupCombobox.SelectedItem }) | ? {$_.type -eq "vSphere"}
			$Nodes = $($treeView1.Nodes | ? { $_.Name -eq $CIGroupCombobox.SelectedItem }).Nodes | ? {$_.Type -eq $CITypeCombobox.SelectedItem}
			ForEach ($Node in $Nodes) { $treeView1.Nodes.Remove($Node) | out-null }
			$CITypeCombobox.Items.Remove($CITypeCombobox.SelectedItem)
			
			If ($CITypeCombobox.Items.count) { $CITypeCombobox.SelectedIndex = 0 }
			Else {
				$CIProperties_GridView.Rows.Clear()
				$CITypeCombobox.Items.Clear()
				$CITypeCombobox.Text = ""
				}
			}
		}
	"CI" {
		If ($treeView1.SelectedNode.Tag -eq "CI") { 
			If ($treeview1.SelectedNode.Parent.Nodes.count -eq 1) { $treeview1.Nodes.Remove($treeview1.SelectedNode.Parent) }
			Else { $treeview1.Nodes.Remove($treeview1.SelectedNode) }
			$treeview1.Select()
			}
		}
	}
PAFCI-FormUpdated
}

Function PAFCI-NewObject {
	param ( [Parameter(Mandatory=$true)] $Type )
Switch ($Type) {
	"CIGroup" {		
		$global:PAFDefaultConfig.CIConfig += [PSCustomObject][Ordered]@{'CIGroup' = $NewCIGroupNameTextBox.Text; 'Description' = $NewCIGroupDescrTextBox.Text; 'Types' = @()}
		$CIGroupCombobox.Items.Add($NewCIGroupNameTextBox.Text)
		If ($CIGroupCombobox.Items.count -eq 1) { $CIGroupCombobox.SelectedIndex = 0 }
		$NewCIGroupForm.Close()
		}
	"CIType" {
		($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Types += [Ordered]@{'Type' = $NewCITypeTextBox.text; 'Properties' = @("Name")}
		$CITypeCombobox.Items.Add($NewCITypeTextBox.text)
		If ($CITypeCombobox.Items.count -eq 1) { $CITypeCombobox.SelectedIndex = 0 }
		$NewCITypeForm.Close()
		}
	}
PAFCI-FormUpdated
}

Function CIGroupCombobox_Changed {
$DescriptionTextBox.Text = ($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Description
$CITypeCombobox.Items.Clear()
ForEach ($Type in ($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Types.Type ) { $CITypeCombobox.Items.Add($Type) | out-null }

If ($CITypeCombobox.Items.count) { $CITypeCombobox.SelectedIndex = 0 }
Else {
	$CITypeCombobox.Text = ""
	$CIProperties_GridView.Rows.Clear()
	}
}

Function CITypeCombobox_Changed {
$CIProperties_GridView.Rows.Clear()

ForEach ($PropertyName in (($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroupCombobox.SelectedItem}).Types | ? {$_.Type -eq $CITypeCombobox.SelectedItem}).Properties) {
	$CIProperties_GridView.Rows.Add($PropertyName,"") | out-null
	}
}

Function PAFCI-UpdateSystem {
If ($treeView1.SelectedNode.Properties) {
	$NewProperties = [PSCustomObject][Ordered]@{}
	ForEach ($row in $CIProperties_GridView.Rows) {
		If ($row.Cells[0].Value) {
			If ($row.Cells[0].Value -like "*password*") { $value = $row.Tag }
			Else { $value = $row.Cells[1].Value }
			$NewProperties | Add-Member noteproperty $row.Cells[0].Value -value $value
			}
		}
	$treeView1.SelectedNode.Properties = $NewProperties
	$treeView1.SelectedNode.CIGroup = $CIGroupCombobox.SelectedItem
	$treeView1.SelectedNode.Type = $CITypeCombobox.SelectedItem
	PAFCI-FormUpdated
	}
}

Function PAFCI-NewSystem {
$Properties = [PSCustomObject][Ordered]@{}
ForEach ($row in $CIProperties_GridView.Rows) {
	If ($row.Cells[0].Value) { $Properties | Add-Member noteproperty $row.Cells[0].Value -value $row.Cells[1].Value }
	}
PAFCI-AddObject -Object @{CIGroup = $CIGroupCombobox.SelectedItem; Type = $CITypeCombobox.SelectedItem; Properties = $Properties} -Type "CI"
}

Function PAFCI-ShowSystem {
$CIGroupCombobox.SelectedItem = $treeView1.SelectedNode.CIGroup
$CITypeCombobox.SelectedItem = $treeView1.SelectedNode.Type

$CIProperties_GridView.Rows.Clear()
ForEach ($PropertyName in $treeView1.SelectedNode.Properties.PSObject.Properties.Name) {
	If ($PropertyName -like "*password*") { $value = If ($treeView1.SelectedNode.Properties.$PropertyName) { '*' * $treeView1.SelectedNode.Properties.$PropertyName.ToString().Length } Else {""} }
	Else { $value = $treeView1.SelectedNode.Properties.$PropertyName }
	$CIProperties_GridView.Rows.Add($PropertyName,$value) | out-null
	}
}

Function PAFCI-NewCITypeForm {
$NewCITypeForm = New-Object system.Windows.Forms.Form
$NewCITypeForm.ClientSize = '360,100'
$NewCITypeForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $IconStream).GetHIcon())
$NewCITypeForm.Text = "CI editor: New CI Type"
$NewCITypeForm.TopMost = $false
$NewCITypeForm.FormBorderStyle = "FixedSingle"
$NewCITypeForm.MaximizeBox = $false

$Label = New-Object system.Windows.Forms.Label
$Label.Text = "Selected CI Group"
$Label.AutoSize = $true
$Label.Width = 25
$Label.Height = 10
$Label.Location = New-Object System.Drawing.Point(10,15)
$NewCITypeForm.Controls.Add($Label)

$Label = New-Object system.Windows.Forms.Label
$Label.Text = "CI Type name"
$Label.AutoSize = $true
$Label.Width = 25
$Label.Height = 10
$Label.Location = New-Object System.Drawing.Point(10,45)
$NewCITypeForm.Controls.Add($Label)

$NewTypeTextBox = New-Object system.Windows.Forms.TextBox
$NewTypeTextBox.multiline = $false
$NewTypeTextBox.Width = 200
$NewTypeTextBox.Height = 20
$NewTypeTextBox.Enabled = $false
$NewTypeTextBox.Text = $CIGroupCombobox.SelectedItem
$NewTypeTextBox.Location = New-Object System.Drawing.Point(150,10)
$NewCITypeForm.Controls.Add($NewTypeTextBox)


$NewCITypeTextBox = New-Object system.Windows.Forms.TextBox
$NewCITypeTextBox.multiline = $false
$NewCITypeTextBox.Width = 200
$NewCITypeTextBox.Height = 20
$NewCITypeTextBox.Location = New-Object System.Drawing.Point(150,40)
$NewCITypeForm.Controls.Add($NewCITypeTextBox)

$OkButton = New-Object system.Windows.Forms.Button
$OkButton.Text = "Ok"
$OkButton.Width = 70
$OkButton.Height = 20
$OkButton.Location = New-Object System.Drawing.Point(200,70)
$OkButton.Add_Click({ PAFCI-NewObject -Type "CIType" })
$NewCITypeForm.Controls.Add($OkButton)

$CancelButton = New-Object system.Windows.Forms.Button
$CancelButton.Text = "Cancel"
$CancelButton.Width = 70
$CancelButton.Height = 20
$CancelButton.Location = New-Object System.Drawing.Point(280,70)
$CancelButton.Add_Click({ $NewCITypeForm.Close() })
$NewCITypeForm.Controls.Add($CancelButton)

$NewCITypeForm.ShowDialog() | out-null 
}

Function PAFCI-NewCIGroupForm {
$NewCIGroupForm = New-Object system.Windows.Forms.Form
$NewCIGroupForm.ClientSize = '360,100'
$NewCIGroupForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $IconStream).GetHIcon())
$NewCIGroupForm.Text = "CI editor: New CI Group"
$NewCIGroupForm.TopMost = $true
$NewCIGroupForm.FormBorderStyle = "FixedSingle"
$NewCIGroupForm.MaximizeBox = $false

$Label1 = New-Object system.Windows.Forms.Label
$Label1.Text = "CI Group name"
$Label1.AutoSize = $true
$Label1.Width = 25
$Label1.Height = 10
$Label1.Location = New-Object System.Drawing.Point(10,15)
$NewCIGroupForm.Controls.Add($Label1)

$Label2 = New-Object system.Windows.Forms.Label
$Label2.Text = "Description"
$Label2.AutoSize = $true
$Label2.Width = 25
$Label2.Height = 10
$Label2.Location = New-Object System.Drawing.Point(10,45)
$NewCIGroupForm.Controls.Add($Label2)

$NewCIGroupNameTextBox = New-Object system.Windows.Forms.TextBox
$NewCIGroupNameTextBox.multiline = $false
$NewCIGroupNameTextBox.Width = 200
$NewCIGroupNameTextBox.Height = 20
$NewCIGroupNameTextBox.Location = New-Object System.Drawing.Point(150,10)
$NewCIGroupForm.Controls.Add($NewCIGroupNameTextBox)

$NewCIGroupDescrTextBox = New-Object system.Windows.Forms.TextBox
$NewCIGroupDescrTextBox.multiline = $false
$NewCIGroupDescrTextBox.Width = 200
$NewCIGroupDescrTextBox.Height = 20
$NewCIGroupDescrTextBox.Location = New-Object System.Drawing.Point(150,40)
$NewCIGroupForm.Controls.Add($NewCIGroupDescrTextBox)

$OkButton = New-Object system.Windows.Forms.Button
$OkButton.Text = "Ok"
$OkButton.Width = 70
$OkButton.Height = 20
$OkButton.Location = New-Object System.Drawing.Point(200,70)
$OkButton.Add_Click({ PAFCI-NewObject -Type "CIGroup" })
$NewCIGroupForm.Controls.Add($OkButton)

$CancelButton = New-Object system.Windows.Forms.Button
$CancelButton.Text = "Cancel"
$CancelButton.Width = 70
$CancelButton.Height = 20
$CancelButton.Location = New-Object System.Drawing.Point(280,70)
$CancelButton.Add_Click({ $NewCIGroupForm.Close() })
$NewCIGroupForm.Controls.Add($CancelButton)

$NewCIGroupForm.ShowDialog() | out-null 
}

Function PAFCI-Save {
param ( $Path ) 
$env = [Ordered]@{}
ForEach ($CIGroup in $treeView1.Nodes) { 
	$CIGroup_obj = @()
	$CIGroup_obj += @{'Label' = $CIGroup.Description}
	ForEach ($Node in $CIGroup.nodes) {
		$CI_obj = [Ordered]@{}
		$CI_obj.Add("Type",$Node.Type)
		ForEach ($Property in $($Node.Properties).PSObject.Properties.Name) {
				If ($Node.Properties.$Property) { $CI_obj.Add($Property,$Node.Properties.$Property) }
				Else { $CI_obj.Add($Property,"") }
			}
		$CIGroup_obj += $CI_obj
		}	
	$env.Add($CIGroup.Name,$CIGroup_obj)
	}

Encrypt -Variable $env | Set-Content -Path $Path

$global:EnvFormUpdated = $false

[System.Windows.Forms.MessageBox]::Show(("Configuration saved"),"PAF Configuration") | out-null
}

Function PAFCI-ExpandCollapseNode {
If ($treeView1.SelectedNode.Tag -eq "root") { 
	If ($treeView1.SelectedNode.IsExpanded) { $treeView1.SelectedNode.Collapse() }
	Else { $treeView1.SelectedNode.Expand() }
	}
}

Function PAFCI-MoveUp {
$idx = $treeview1.SelectedNode.Index
$node = $treeview1.SelectedNode
If ($treeview1.SelectedNode.Tag -eq "CI") {
If ($treeview1.SelectedNode.Index -gt 0) {
	$parent = $node.Parent
	$parent.Nodes.Remove($node)
	$parent.Nodes.Insert($idx - 1 , $node)
	}
}
Else {
If ($treeview1.SelectedNode.Index -gt 0) {
	$CIGroupCombobox.Items.Remove($treeview1.SelectedNode.Name)
	$CIGroupCombobox.Items.Insert($idx - 1, $treeview1.SelectedNode.Name)
	$CIGroupCombobox.SelectedIndex = $idx - 1
	$treeview1.SelectedNode.Nodes.Remove($node)
	$treeview1.Nodes.Insert($idx - 1, $node)	
	}
}
$treeview1.SelectedNode = $node
$treeview1.Select()
PAFCI-FormUpdated
}

Function PAFCI-MoveDown {
$idx = $treeview1.SelectedNode.Index
$node = $treeview1.SelectedNode
If ($treeview1.SelectedNode.Tag -eq "CI") {
If ($treeview1.SelectedNode.Index -lt $treeview1.SelectedNode.Parent.Nodes.count) {
	$parent = $node.Parent
	$parent.Nodes.Remove($node)
	$parent.Nodes.Insert($idx + 1, $node)
	}
}
Else {
If ($treeview1.SelectedNode.Index -lt $treeview1.Nodes.count - 1) {
	$CIGroupCombobox.Items.Remove($treeview1.SelectedNode.Name)
	$CIGroupCombobox.Items.Insert($idx + 1, $treeview1.SelectedNode.Name)
	$CIGroupCombobox.SelectedIndex = $idx + 1
	$treeview1.SelectedNode.Nodes.Remove($node)
	$treeview1.Nodes.Insert($idx + 1 , $node)
	}
}
$treeview1.SelectedNode = $node
$treeview1.Select()
PAFCI-FormUpdated
}

Function PAFCI-Scan {
$IPRange = @()
[System.Windows.Forms.Cursor]::Current = 'WaitCursor'
$DiscoveredSystemsGridview.UseWaitCursor = $true
#Merge IP ranges
ForEach ($Row in $IPGridView.Rows.Index) {
	If ($IPGridView[0,$Row].Value -and $IPGridView[1,$Row].Value) { $IPRange += Get-IPrange -start $IPGridView[0,$Row].Value -end $IPGridView[1,$Row].Value }
	}

#$DiscoveredSystemsGridview.Rows.Clear()
$i = $DiscoveredSystemsGridview.Rows.GetRowCount([System.Windows.Forms.DataGridViewElementStates]::Visible) - 1

$DiscoveredSystems = Invoke-Parallel -ScriptFile $($PSScriptRoot + "\scan-IP.ps1") -InputObject $IPRange -Quiet
$ProgressBar.Value = 100
$DiscoveredSystemsGridview.UseWaitCursor = $false

ForEach ($DiscoveredSystem in $DiscoveredSystems | Sort-Object IP) {
	If ($DiscoveredSystem.'System Type'.count -gt 1) {
		$cell = New-Object System.Windows.Forms.DataGridViewComboBoxCell
		ForEach ($Type in $DiscoveredSystem.'System Type') { [void] $cell.Items.Add($Type) }
		$cell.Value  = $DiscoveredSystem.'System Type'[0]
		$DiscoveredSystemsGridview.Rows.Add($DiscoveredSystem.'IP',$DiscoveredSystem.'DNS') | out-null
		$DiscoveredSystemsGridview[3,$i] = $cell
		}
	Else {
		$DiscoveredSystemsGridview.Rows.Add($DiscoveredSystem.'IP',$DiscoveredSystem.'DNS',"", $DiscoveredSystem.'System Type') | out-null
		$DiscoveredSystemsGridview[3,$i].ReadOnly = $true
		}
	$i++
	}
}

Function PAFCI-ScanForm-Delete {
If ([System.Windows.Forms.MessageBox]::Show(("Are you sure you want to exclude selected systems?"),"PAF Configuration",4,48) -eq "Yes") {
	$DiscoveredSystemsGridview.SelectedRows | ForEach-Object {
		If ($DiscoveredSystemsGridview[0,$_.Index].Value) { $DiscoveredSystemsGridview.Rows.Remove($_) }
		}
	}
}

Function PAFCI-ScanForm-Export {
For ($i = 0; $i -lt $DiscoveredSystemsGridview.Rows.GetRowCount([System.Windows.Forms.DataGridViewElementStates]::Visible) - 1; $i++) {
	$name = If ($DiscoveredSystemsGridview[2,$i].Value) { $DiscoveredSystemsGridview[2,$i].Value } Else { If ($DiscoveredSystemsGridview[1,$i].Value) { $DiscoveredSystemsGridview[1,$i].Value } Else { $DiscoveredSystemsGridview[0,$i].Value } }
	$type = $DiscoveredSystemsGridview[3,$i].Value	
	PAFCI-AddObject -Object @{CIGroup = ($global:Config.Properties.CIConfig | ? {$_.Types -contains $type}).CIGroup; Type = $type; Properties = @($name,$DiscoveredSystemsGridview[0,$i].Value,"",$DiscoveredSystemsGridview[4,$i].Value,$DiscoveredSystemsGridview[5,$i].Value)} -Type "CI"
	}
[System.Windows.Forms.MessageBox]::Show(("Discovered system added to the environment`nPlease remember that some systems might require additional setting to be configured besides username and password!"),"PAF Configuration") | out-null
$DiscoveredSystemsGridview.Rows.Clear()
}

#####Environment search form here
Function PAFCI-ScanForm {
$DiscoveryForm = New-Object System.Windows.Forms.Form
$DiscoveryForm.Size = New-Object System.Drawing.Size(700,769)
$DiscoveryForm.Text = "Environment Editor tool: system discovery"
$DiscoveryForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $IconStream).GetHIcon())
$DiscoveryForm.StartPosition = "CenterScreen"
$DiscoveryForm.FormBorderStyle = "FixedSingle"
$DiscoveryForm.MaximizeBox = $false

$Label = New-Object system.Windows.Forms.Label
$Label.Text = "Enter IP address ranges:"
$Label.AutoSize = $true
$Label.Location = New-Object System.Drawing.Point(10,20)
$DiscoveryForm.Controls.Add($Label)

$IPGridView = New-Object System.Windows.Forms.DataGridView
$IPGridView.Location = New-Object System.Drawing.Size(10,50) 
$IPGridView.Size = New-Object System.Drawing.Size(250,113)
$IPGridView.AllowUserToOrderColumns = $True
$IPGridView.AutoSizeColumnsMode = 'fill'
$IPGridView.RowHeadersVisible = $false;
$IPGridView.ColumnCount = 2
$IPGridView.SelectionMode = 'FullRowSelect'
$IPGridView.Columns[0].Name = "Start IP"
$IPGridView.Columns[1].Name = "End IP"
$DiscoveryForm.Controls.Add($IPGridView)
$IPGridView.Rows.Add() | out-null
$IPGridView.Rows.Add() | out-null
$IPGridView.Rows.Add() | out-null

$ScanButton = New-Object System.Windows.Forms.Button
$ScanButton.Location = New-Object System.Drawing.Size(350,50) 
$ScanButton.Size = New-Object System.Drawing.Size(250,30)
$ScanButton.Text = "Scan"
$ScanButton.Add_Click({ PAFCI-Scan })
$DiscoveryForm.Controls.Add($ScanButton)

$DiscoveredSystemsGridview = New-Object System.Windows.Forms.DataGridView
$DiscoveredSystemsGridview.Location = New-Object System.Drawing.Size(0,185) 
$DiscoveredSystemsGridview.Size = New-Object System.Drawing.Size(685,465)
$DiscoveredSystemsGridview.AllowUserToOrderColumns = $True
$DiscoveredSystemsGridview.AutoSizeColumnsMode = 'fill'
$DiscoveredSystemsGridview.RowHeadersVisible = $false;
$DiscoveredSystemsGridview.ColumnCount = 6
$DiscoveredSystemsGridview.SelectionMode = 'FullRowSelect'
$DiscoveredSystemsGridview.Columns[0].Name = "IP"
$DiscoveredSystemsGridview.Columns[0].ReadOnly = $true
$DiscoveredSystemsGridview.Columns[1].Name = "DNS"
$DiscoveredSystemsGridview.Columns[1].ReadOnly = $true
$DiscoveredSystemsGridview.Columns[2].Name = "System name"
$DiscoveredSystemsGridview.Columns[3].Name = "System type"
$DiscoveredSystemsGridview.Columns[4].Name = "Username"
$DiscoveredSystemsGridview.Columns[5].Name = "Password"
$DiscoveryForm.Controls.Add($DiscoveredSystemsGridview)

$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Location = New-Object System.Drawing.Size(15,670) 
$ExportButton.Size = New-Object System.Drawing.Size(205,30)
$ExportButton.Text = "Add to environment"
$ExportButton.Add_Click({ PAFCI-ScanForm-Export })
$DiscoveryForm.Controls.Add($ExportButton)

$ScanDeleteButton = New-Object System.Windows.Forms.Button
$ScanDeleteButton.Location = New-Object System.Drawing.Size(235,670) 
$ScanDeleteButton.Size = New-Object System.Drawing.Size(205,30)
$ScanDeleteButton.Text = "Delete selected"
$ScanDeleteButton.Add_Click({ PAFCI-ScanForm-Delete })
$DiscoveryForm.Controls.Add($ScanDeleteButton)

$ScanCloseButton = New-Object System.Windows.Forms.Button
$ScanCloseButton.Location = New-Object System.Drawing.Size(460,670) 
$ScanCloseButton.Size = New-Object System.Drawing.Size(205,30)
$ScanCloseButton.Text = "Close"
$ScanCloseButton.Add_Click({ $DiscoveryForm.Close() })
$DiscoveryForm.Controls.Add($ScanCloseButton)



$ProgressBar = New-Object system.Windows.Forms.ProgressBar
$ProgressBar.width = 681
$ProgressBar.height = 10
$ProgressBar.location = New-Object System.Drawing.Point(2,720)
$DiscoveryForm.Controls.Add($ProgressBar)

[void] $DiscoveryForm.ShowDialog()
}

Function PAFCI-FormUpdated { $global:EnvFormUpdated = $true }

Function PAFCI-DeleteEmptyRow {
Try { 
	If (!$CIProperties_GridView.Rows[$CIProperties_GridView.CurrentCell.RowIndex].Cells[0].value -and !$CIProperties_GridView.Rows[$CIProperties_GridView.CurrentCell.RowIndex].Cells[1].value) {
		$CIProperties_GridView.Rows.RemoveAt($CIProperties_GridView.CurrentCell.RowIndex)
		} 
	}
Catch {}
}

Function PAFCI-HidePasswords { 
If ($CIProperties_GridView.Rows[$CIProperties_GridView.CurrentCell.RowIndex].Cells[0].value -like "*password*" ) {
	$passw = $CIProperties_GridView.Rows[$CIProperties_GridView.CurrentCell.RowIndex].Cells[1].Value
	$CIProperties_GridView.Rows[$CIProperties_GridView.CurrentCell.RowIndex].Tag = $passw;
	$CIProperties_GridView.Rows[$CIProperties_GridView.CurrentCell.RowIndex].Cells[1].Value = '*' * $CIProperties_GridView.Rows[$CIProperties_GridView.CurrentCell.RowIndex].Cells[1].Value.ToString().Length
	}

}

#####Editor form here
Function PAFCI-EditForm {
param ( $environment, $path ) 
#Load environment

$global:EnvFormUpdated = $false

$Systems = @()

ForEach ($CIGroup in $environment.PSObject.Properties.Name) {
	If ($global:PAFDefaultConfig.CIConfig.CIGroup -notcontains $CIGroup) {
		$global:PAFDefaultConfig.CIConfig += [PSCustomObject][Ordered]@{'CIGroup' = $CIGroup; 'Description' = $environment.$CIGroup.Label; 'Types' = @()}
		}
	
	ForEach ($CIType in $environment.$CIGroup | ? {!$_.Label}) {
		If (($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroup}).Types.Type -notcontains $CIType.Type) {	
			($global:PAFDefaultConfig.CIConfig | ? {$_.CIGroup -eq $CIGroup}).Types += [Ordered]@{'Type' = $CIType.Type; 'Properties' = $CIType.PSObject.Properties.Name | ? {$_ -ne "Type"}}
			}	
		$Systems += [PSCustomObject][Ordered]@{CIGroup = $CIGroup; Type = $CIType.Type; Properties = $CIType | Select * -ExcludeProperty "Type"}
		}
	}

#Minimize main form
#$ConfForm.WindowState = "Minimized"

$EnvForm = New-Object System.Windows.Forms.Form
$EnvForm.Text = $($Path -replace ".*\\") + " - CI editor"
$EnvForm.Name = "EnvForm" 
$EnvForm.DataBindings.DefaultDataSourceUpdateMode = 0
$EnvForm.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $IconStream).GetHIcon())
$EnvForm.ClientSize = New-Object System.Drawing.Size(660,565) 
$EnvForm.FormBorderStyle = "FixedSingle"
$EnvForm.MaximizeBox = $false
$EnvForm.StartPosition = "centerscreen"
$EnvForm.TopMost = $false
$EnvForm.AutoSize = $true
$EnvForm.KeyPreview = $true;
$EnvForm.Add_KeyDown( { If ($_.Control -and $_.KeyCode -eq "Q") { $EnvForm.Close() } If ($_.Control -and $_.KeyCode -eq "S") { PAFCI-Save -Path $Path } } )
$EnvForm.Add_FormClosing({ If ($global:EnvFormUpdated) { If ([System.Windows.Forms.MessageBox]::Show(("Do you want to save changes to $($Path -replace '.*\\')?`n`nIf you click ""No"" last changes will not be saved"),"PAF Configuration",4) -eq "Yes") { PAFConfig-Save } };  $ConfForm.WindowState = "Normal" })

$label = New-Object System.Windows.Forms.Label
$label.Name = "label1" 
$label.Location = New-Object System.Drawing.Size(12,5)
$Label.AutoSize = $true
$label.Text = "CI treeview"
$EnvForm.Controls.Add($label)

$treeView1 = New-Object System.Windows.Forms.TreeView
$treeView1.Size = New-Object System.Drawing.Size(250,485)
$treeView1.Name = "treeView1" 
$treeView1.Location = New-Object System.Drawing.Size(15,25)
$treeView1.DataBindings.DefaultDataSourceUpdateMode = 0 
$treeView1.TabIndex = 0
$treeView1.HideSelection = $false
$treeView1.Add_KeyDown( { If (($_.KeyCode -eq "Enter")) { PAFCI-ShowSystem; PAFCI-ExpandCollapseNode } If (($_.KeyCode -eq "Delete")) { PAFCI-DeleteObject -Type "Treeview" } })
$treeView1.Add_DoubleClick({ PAFCI-ShowSystem })
$EnvForm.Controls.Add($treeView1)

$SearchButton = New-Object system.Windows.Forms.Button
$SearchButton.Size = New-Object System.Drawing.Size(30,30)
$SearchButton.Location = New-Object System.Drawing.Point(270,25)
$SearchButton.Image = [System.Convert]::FromBase64String($SearchImg)
$SearchButton.Add_Click({ PAFCI-ScanForm })
$EnvForm.Controls.Add($SearchButton)

$UpButton = New-Object system.Windows.Forms.Button
$UpButton.Size = New-Object System.Drawing.Size(30,30)
$UpButton.Location = New-Object System.Drawing.Point(270,227)
$UpButton.Image = [System.Convert]::FromBase64String($UpImg)
$UpButton.Add_Click({ PAFCI-MoveUp })
$EnvForm.Controls.Add($UpButton)

$DeleteButton = New-Object system.Windows.Forms.Button
$DeleteButton.Size = New-Object System.Drawing.Size(30,30)
$DeleteButton.Location = New-Object System.Drawing.Point(270,262)
$DeleteButton.Image = [System.Convert]::FromBase64String($RecycleBinImg)
$DeleteButton.Add_Click({ PAFCI-DeleteObject -Type "Treeview" })
$EnvForm.Controls.Add($DeleteButton)

$DownButton = New-Object system.Windows.Forms.Button
$DownButton.Size = New-Object System.Drawing.Size(30,30)
$DownButton.Location = New-Object System.Drawing.Point(270,297)
$DownButton.Image = [System.Convert]::FromBase64String($DownImg)
$DownButton.Add_Click({ PAFCI-MoveDown })
$EnvForm.Controls.Add($DownButton)

$Groupbox1 = New-Object system.Windows.Forms.Groupbox
$Groupbox1.Width = 340
$Groupbox1.Height = 450
$Groupbox1.Text = "CI Details"
$Groupbox1.Location = New-Object System.Drawing.Point(305,20)
$EnvForm.Controls.Add($Groupbox1)

$DescriptionLabel = New-Object system.Windows.Forms.Label
$DescriptionLabel.Name = "DescriptionLabel" 
$DescriptionLabel.Location = New-Object System.Drawing.Size(10,20)
$DescriptionLabel.AutoSize = $true
$DescriptionLabel.Text = "Description"
$Groupbox1.Controls.Add($DescriptionLabel)

$DescriptionTextBox = New-Object system.Windows.Forms.TextBox
$DescriptionTextBox.Location = New-Object System.Drawing.Size(100,20)
$DescriptionTextBox.Width = 225
$DescriptionTextBox.Height = 20
$DescriptionTextBox.Name = "DescriptionTextBox"
$DescriptionTextBox.enabled = $false
$Groupbox1.Controls.Add($DescriptionTextBox)

$label = New-Object System.Windows.Forms.Label
$label.Name = "label2" 
$label.Location = New-Object System.Drawing.Size(10,50)
$Label.AutoSize = $true
$label.Text = "CI Group"
$Groupbox1.Controls.Add($label)

$CIGroupCombobox = New-Object System.Windows.Forms.ComboBox
$CIGroupCombobox.FormattingEnabled = $True
$CIGroupCombobox.Location = New-Object System.Drawing.Size(100,50)
$CIGroupCombobox.DropDownStyle = "DropDownList"
$CIGroupCombobox.Name = "CI Group"
$CIGroupCombobox.Size = New-Object System.Drawing.Size(140,20)
$CIGroupCombobox.add_SelectedIndexChanged({ CIGroupCombobox_Changed })
$Groupbox1.Controls.Add($CIGroupCombobox)

$CreateCIGroupButton = New-Object system.Windows.Forms.Button
$CreateCIGroupButton.Size = New-Object System.Drawing.Size(20,20)
$CreateCIGroupButton.Location = New-Object System.Drawing.Point(245,50)
$CreateCIGroupButton.Image = [System.Convert]::FromBase64String($AddImg)
$CreateCIGroupButton.Add_Click({ PAFCI-NewCIGroupForm })
$Groupbox1.Controls.Add($CreateCIGroupButton)

$DeleteCIGroupButton = New-Object system.Windows.Forms.Button
$DeleteCIGroupButton.Size = New-Object System.Drawing.Size(20,20)
$DeleteCIGroupButton.Location = New-Object System.Drawing.Point(270,50)
$DeleteCIGroupButton.Image = [System.Convert]::FromBase64String($DeleteImg)
$DeleteCIGroupButton.Add_Click({ PAFCI-DeleteObject -Type "CIGroup" })
$Groupbox1.Controls.Add($DeleteCIGroupButton)

$label = New-Object System.Windows.Forms.Label
$label.Name = "label3" 
$label.Location = New-Object System.Drawing.Size(10,80)
$Label.AutoSize = $true
$label.Text = "CI Type"
$label.Text = "CI Type"
$Groupbox1.Controls.Add($label)

$CITypeCombobox = New-Object System.Windows.Forms.ComboBox
$CITypeCombobox.FormattingEnabled = $false
$CITypeCombobox.Size = New-Object System.Drawing.Size(140,20)
$CITypeCombobox.Location = New-Object System.Drawing.Size(100,80)
$CITypeCombobox.DropDownStyle = "DropDownList"
$CITypeCombobox.Name = "CI Type"
$CITypeCombobox.add_SelectedIndexChanged({ CITypeCombobox_Changed })
$Groupbox1.Controls.Add($CITypeCombobox)

$CreateCITypeButton = New-Object system.Windows.Forms.Button
$CreateCITypeButton.Size = New-Object System.Drawing.Size(20,20)
$CreateCITypeButton.Location = New-Object System.Drawing.Point(245,80)
$CreateCITypeButton.Image = [System.Convert]::FromBase64String($AddImg)
$CreateCITypeButton.Add_Click({ PAFCI-NewCITypeForm })
$Groupbox1.Controls.Add($CreateCITypeButton)

$DeleteCITypeButton = New-Object system.Windows.Forms.Button
$DeleteCITypeButton.Size = New-Object System.Drawing.Size(20,20)
$DeleteCITypeButton.Location = New-Object System.Drawing.Point(270,80)
$DeleteCITypeButton.Image = [System.Convert]::FromBase64String($DeleteImg)
$DeleteCITypeButton.Add_Click({ PAFCI-DeleteObject -Type "CIType" })
$Groupbox1.Controls.Add($DeleteCITypeButton)

$CIProperties_GridView = New-Object system.Windows.Forms.DataGridView
$CIProperties_GridView.width = 320
$CIProperties_GridView.height = 290
$CIProperties_GridView.Location = New-Object System.Drawing.Size(10,110)
$CIProperties_GridView.ColumnCount = 2
$CIProperties_GridView.AutoSizeColumnsMode = 'fill'
$CIProperties_GridView.ColumnHeadersVisible = $true
$CIProperties_GridView.AllowUserToResizeRows = $false
$CIProperties_GridView.RowHeadersVisible = $false
$CIProperties_GridView.Columns[0].Name = "Property name"
$CIProperties_GridView.Columns[0].width = 120
$CIProperties_GridView.Columns[0].SortMode = 0
$CIProperties_GridView.Columns[1].Name = "Property value"
$CIProperties_GridView.Columns[1].SortMode = 0
$CIProperties_GridView.Add_KeyDown( { If (($_.KeyCode -eq "Delete")) { $CIProperties_GridView.Rows[$CIProperties_GridView.CurrentCell.RowIndex].Cells[$CIProperties_GridView.CurrentCell.ColumnIndex].value = "";  } })
$CIProperties_GridView.Add_CellValueChanged({ PAFCI-DeleteEmptyRow })
$CIProperties_GridView.Add_CellEndEdit({ PAFCI-HidePasswords; PAFCI-DeleteEmptyRow })

$Groupbox1.Controls.Add($CIProperties_GridView)

$NewSystemButton = New-Object system.Windows.Forms.Button
$NewSystemButton.Name = "NewSystemButton" 
$NewSystemButton.Text = "Add new system"
$NewSystemButton.Size = New-Object System.Drawing.Size(315,30)
$NewSystemButton.Location = New-Object System.Drawing.Point(10,410)
$NewSystemButton.Add_Click({ PAFCI-NewSystem })
$Groupbox1.Controls.Add($NewSystemButton)

$Button1 = New-Object system.Windows.Forms.Button
$Button1.Name = "Button1" 
$Button1.Text = "&Save system properties"
$Button1.Size = New-Object System.Drawing.Size(340,30)
$Button1.Location = New-Object System.Drawing.Point(305,480)
$Button1.Add_Click({ PAFCI-UpdateSystem }) #Select-Asset $treeView1.SelectedNode.Name
$EnvForm.Controls.Add($Button1)

$CloseEnvButton = New-Object system.Windows.Forms.Button
$CloseEnvButton.Name = "CloseEnvButton" 
$CloseEnvButton.Text = "Close"
$CloseEnvButton.Size = New-Object System.Drawing.Size(100,30)
$CloseEnvButton.Location = New-Object System.Drawing.Point(425,520)
$CloseEnvButton.Add_Click({ $EnvForm.Close() })
$EnvForm.Controls.Add($CloseEnvButton) 

$SaveButton = New-Object system.Windows.Forms.Button
$SaveButton.Name = "SaveButton" 
$SaveButton.Text = "&Save"
$SaveButton.Size = New-Object System.Drawing.Size(100,30)
$SaveButton.Location = New-Object System.Drawing.Point(545,520)
$SaveButton.Add_Click({ PAFCI-Save -Path $Path })
$EnvForm.Controls.Add($SaveButton) 
 
#Save the initial state of the form 
$InitialFormWindowState = $EnvForm.WindowState 

#Genereta systems tree
PAFCI-BuildTree

$global:EnvFormUpdated = $false

#Show the Form 
$EnvForm.ShowDialog() | out-null
}



#PAFCI-EditForm -Environment $ee -Path $Path