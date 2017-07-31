<#
.SYNOPSIS
	Insane Move - Copy sites to Office 365 in parallel.  ShareGate Insane Mode times ten!
.DESCRIPTION
	Copy SharePoint site collections to Office 365 in parallel.  CSV input list of source/destination URLs.  XML with general preferences.
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='CSV list of source and destination SharePoint site URLs to copy to Office 365.')]
	[string]$fileCSV,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Verify all Office 365 site collections.  Prep step before real migration.')]
	[Alias("v")]
	[switch]$verifyCloudSites = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Copy incremental changes only. http://help.share-gate.com/article/443-incremental-copy-copy-sharepoint-content')]
	[Alias("i")]
	[switch]$incremental = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Measure size of site collections in GB.')]
	[Alias("m")]
	[switch]$measure = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Lock sites read-only.')]
	[Alias("ro")]
	[switch]$readOnly = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Unlock sites read-write.')]
	[Alias("rw")]
	[switch]$readWrite = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Lock sites no access.')]
	[Alias("na")]
	[switch]$noAccess = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Update local User Profile Service with cloud personal URL.  Helps with Hybrid Onedrive audience rules.  Need to recompile audiences after running this.')]
	[Alias("ups")]
	[switch]$userProfileSetHybridURL = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Dry run replaces core "Copy-Site" with "NoCopy-Site" to execute all queueing but not transfer any data.')]
	[Alias("d")]
	[switch]$dryRun = $false,

	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Clean servers to preprae for next migration batch.')]
	[Alias("c")]
	[switch]$clean = $false
)

# Plugin
Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -Prefix M | Out-Null
Import-Module SharePointPnPPowerShellOnline -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module CredentialManager -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

# Config
$root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[xml]$settings = Get-Content "$root\InsaneMove.xml"
"$root\InsaneMove.xml"
$maxWorker = $settings.settings.maxWorker

Function VerifyPSRemoting() {
	"<VerifyPSRemoting>"
	$ssp = Get-WSManCredSSP
	if ($ssp[0] -match "not configured to allow delegating") {
		# Enable remote PowerShell over CredSSP authentication
		Enable-WSManCredSSP -DelegateComputer * -Role Client -Force
		Restart-Service WinRM
	}
}

Function ReadIISPW {
	"<ReadIISPW>"
	# Read IIS password for current logged in user
	$pass = $null
	Write-Host "===== Read IIS PW ===== $(Get-Date)" -Fore Yellow

	# Current user (ex: Farm Account)
	$userdomain = $env:userdomain
	$username = $env:username
	Write-Host "Logged in as $userdomain\$username"
	
	# Start IISAdmin if needed
	$iisadmin = Get-Service IISADMIN
	if ($iisadmin.Status -ne "Running") {
		#set Automatic and Start
		Set-Service -Name IISADMIN -StartupType Automatic -ErrorAction SilentlyContinue
		Start-Service IISADMIN -ErrorAction SilentlyContinue
	}
	
	# Attempt to detect password from IIS Pool (if current user is local admin and farm account)
	Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null
	$m = Get-Module WebAdministration
	if ($m) {
		#PowerShell ver 2.0+ IIS technique
		$appPools = Get-ChildItem "IIS:\AppPools\"
		foreach ($pool in $appPools) {	
			if ($pool.processModel.userName -like "*$username") {
				Write-Host "Found - "$pool.processModel.userName
				$pass = $pool.processModel.password
				if ($pass) {
					break
				}
			}
		}
	} else {
		#PowerShell ver 3.0+ WMI technique
		$appPools = Get-CimInstance -Namespace "root/MicrosoftIISv2" -ClassName "IIsApplicationPoolSetting" -Property Name, WAMUserName, WAMUserPass | select WAMUserName, WAMUserPass
		foreach ($pool in $appPools) {	
			if ($pool.WAMUserName -like "*$username") {
				Write-Host "Found - "$pool.WAMUserName
				$pass = $pool.WAMUserPass
				if ($pass) {
					break
				}
			}
		}
	}

	# Prompt for password
	if (!$pass) {
		$pass = Read-Host "Enter password for $userdomain\$username"
	} 
	$sec = $pass | ConvertTo-SecureString -AsPlainText -Force
	$global:pass = $pass
	$global:cred = New-Object System.Management.Automation.PSCredential -ArgumentList "$userdomain\$username", $sec
}

Function DetectVendor() {
	"<DetectVendor>"
	# SharePoint Servers in local farm
	$spservers = Get-SPServer |? {$_.Role -ne "Invalid"} | sort Address

	# Detect if Vendor software installed
	$coll = @()
	foreach ($s in $spservers) {
		$found = Get-ChildItem "\\$($s.Address)\C$\Program Files (x86)\Sharegate\Sharegate.exe" -ErrorAction SilentlyContinue
		if ($found) {
			if ($settings.settings.optionalLimitServers) {
				if ($settings.settings.optionalLimitServers.ToUpper().Contains($s.Address.ToUpper())) {
					$coll += $s.Address
				}
			} else {
				$coll += $s.Address
			}
		}
	}
	
	# Display and return
	$coll |% {Write-Host $_ -Fore Green}
	$global:servers = $coll
	
	# Safety
	if (!$coll) {
		Write-Host "No Servers Have ShareGate Installed.  Please Verify." -Fore Red
		Exit
	}
}

Function ReadCloudPW() {
	"<ReadCloudPW>"
	# Prompt for admin password
	if ($settings.settings.tenant.adminPass) {
		$global:cloudPW =$settings.settings.tenant.adminPass
	} else {
		$global:cloudPW = Read-Host "Enter O365 Cloud Password for $($settings.settings.tenant.adminUser)"
	}
}

Function CloseSession() {
	"<CloseSession>"
	# Close remote PS sessions
	Get-PSSession | Remove-PSSession
}

Function CreateWorkers() {
	"<CreateWorkers>"
	# Open worker sessions per server.  Runspace to create local SCHTASK on remote PC
    # Template command
    $cmdTemplate = @'
mkdir "d:\InsaneMove" -ErrorAction SilentlyContinue | Out-Null

Function VerifySchtask($name, $file) {
	$found = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
	if ($found) {
		$found | Unregister-ScheduledTask -Confirm:$false
	}

	$user = "[RUNASDOMAIN]\[RUNASUSER]"
	$pw = "[RUNASPASS]"
	
	$folder = Split-Path $file
	$a = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $file -WorkingDirectory $folder
	$p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType Password
	$task = New-ScheduledTask -Action $a -Principal $p
	return Register-ScheduledTask -TaskName $name -InputObject $task -Password $pw -User $user
}

VerifySchtask "worker1-[RUNASUSER]" "d:\InsaneMove\worker1-[RUNASUSER].ps1"
'@
$cmdTemplate = $cmdTemplate.replace("[RUNASDOMAIN]", $env:userdomain)

# Loop available servers
	$global:workers = @()
	$wid = 0
	
	foreach ($pc in $global:servers) {
		# Loop maximum worker
		$s = New-PSSession -ComputerName $pc -Credential $global:cred -Authentication CredSSP -ErrorAction SilentlyContinue
        $s
        1..$maxWorker |% {
			# Optional - run odd SCHTASK (1,3,5...) as different account 
			$runAsUser = $env:username
			if ($settings.settings.optionalSchtaskUser) {
				if ($wid % 2 -eq 1) {
					# Odd number worker # schtask
					$runAsUser = $settings.settings.optionalSchtaskUser
				}
			}
			
			# Assume both RUN AS account share the global password
			$runAsPass = $global:pass.replace("`$","``$")
		
            # create worker
			$runAsUser = $runAsUser.ToUpper()
			$current = $cmdTemplate.replace("[RUNASUSER]", $runAsUser)
			$current = $current.replace("[RUNASPASS]", $runAsPass)
            $current = $current.replace("worker1","worker$wid")
			$current
            Write-Host "CREATE Worker$wid-$runAsUser on $pc ..." -Fore Yellow
            $sb = [Scriptblock]::Create($current)
            $result = Invoke-Command -Session $s -ScriptBlock $sb
			"[RESULT]"
            $result | ft -a
			
			# purge old worker XML output
			$resultfile = "\\$pc\d$\insanemove\worker$wid-$runAsUser.xml"
            Remove-Item $resultfile -confirm:$false -ErrorAction SilentlyContinue
			
			# upload user
			$uploadUsers = $settings.settings.tenant.uploadUsers.Split(",")
			
            # track worker
			$worker = New-Object -TypeName PSObject -Prop (@{"Id"=$wid;"PC"=$pc;"RunAsUser"=$runAsUser;"UploadUser"=$uploadUsers[$wid]})
			$global:workers += $worker

			# Windows - Credential Manager
			New-StoredCredential -Target "InsaneMove-$runAsUser" -UserName $runAsUser -Password $runAsPass -Persist LocalMachine
			
			# Increment counters
			$wid++
		}
	}
	Write-Host "WORKERS" -Fore Green
	$global:workers | ft -a
}

Function PrepareCloudUrl($destUrl) {
	# find managed path
	$destEndChar = $destUrl.IndexOf("-sites/")
	$managedPath = $destUrl.Substring(0,$destEndChar).Split("/")[-1]

	# parse pre/post that go before/after managed path
	$pre = $destUrl.Substring(0, $destEndChar - $managedPath.length)
	$post = $destUrl.Substring($destEndChar+7, $destUrl.length-7-$destEndChar)
	$pre = $pre.Replace("http://sharepoint","https://tenant.sharepoint.com")

	# return final URL
	return $pre + "sites/" + $managedPath + "-" + $post
}

Function CreateTracker() {
	"<CreateTracker>"
	# CSV migration source/destination URL
	Write-Host "===== Populate Tracking table ===== $(Get-Date)" -Fore Yellow

	$global:track = @()
	$csv = Import-Csv $fileCSV
	$i = 0	
	foreach ($row in $csv) {		
		# Get SharePoint total storage
		$site = Get-SPSite $row.SourceURL
		if ($site) {
			$SPStorage = [Math]::Round($site.Usage.Storage / 1MB, 2)
		} else {
			Write-Host "SITE NOT FOUND $($row.SourceURL)"
		}
		
		# MySite URL Lookup
		if ($row.MySiteEmail) {
			$destUrl = FindCloudMySite $row.MySiteEmail
		} else {
			$destUrl = $row.DestinationURL;
			$destEndChar = $destUrl.IndexOf("-sites/")

			# adjust destination URL
			if ($destUrl.contains("-sites")) {
				$destUrl = PrepareCloudUrl $destUrl
			}
		}

		# Add row
		$obj = New-Object -TypeName PSObject -Prop (@{
			"SourceURL"=$row.SourceURL;
			"DestinationURL"=$destUrl;
			"MySiteEmail"=$row.MySiteEmail;
			"CsvID"=$i;
			"WorkerID"="";
			"PC"="";
			"RunAsUser"="";
			"Status"="New";
			"SGResult"="";
			"SGServer"="";
			"SGSessionId"="";
			"SGSiteObjectsCopied"="";
			"SGItemsCopied"="";
			"SGWarnings"="";
			"SGErrors"="";
			"Error"="";
			"ErrorCount"="";
			"TaskXML"="";
			"SPStorage"=$SPStorage;
			"TimeCopyStart"="";
			"TimeCopyEnd"=""
		})
		$global:track += $obj

		# Increment ID
		$i++
	}
	
	# Sort by SharePoint site storage (GB) ascending (small sites first)
    $global:track = $global:track | sort SPStorage
	
	# Display
	"[SESSION-CreateTracker]"
	Get-PSSession | ft -a
}

Function UpdateTracker () {
	"<UpdateTracker>"
	# Update tracker with latest SCHTASK status
	$active = $global:track |? {$_.Status -eq "InProgress"}
	foreach ($row in $active) {
		# Monitor remote SCHTASK
		$wid = $row.WorkerID
        $pc = $row.PC
		
		# Reconnect Broken remote PS
		$broken = Get-PSSession |? {$_.State -ne "Opened"}
		if ($broken) {
			# Make new session
			if ($broken -is [array]) {
				# Multiple
				foreach ($brokenCurrent in $broken) {
					New-PSSession -ComputerName $brokenCurrent.ComputerName -Credential $global:cred -Authentication CredSSP -ErrorAction SilentlyContinue
					$brokenCurrent | Remove-PSSession
				}
			} else {
				# Single
				New-PSSession -ComputerName $broken.ComputerName -Credential $global:cred -Authentication CredSSP -ErrorAction SilentlyContinue
				$broken | Remove-PSSession
			}
		}
		
		# Lookup worker user
		$worker = $global:workers |? {$_.Id -eq $row.WorkerID}
		$runAsUser = $worker.RunAsUser
		
		# Check SCHTASK State=Ready
		$s = Get-PSSession |? {$_.ComputerName -eq $pc}
		$cmd = "Get-Scheduledtask -TaskName 'worker$wid-$runAsUser'"
		$sb = [Scriptblock]::Create($cmd)
		$schtask = $null
		$schtask = Invoke-Command -Session $s -Command $sb
		if ($schtask) {
			"[SCHTASK]"
			$schtask | select {$pc},TaskName,State | ft -a
			"[SESSION-UpdateTracker]"
			Get-PSSession | ft -a
			if ($schtask.State -eq 3) {
				# Completed
				$row.Status = "Completed"
				$row.TimeCopyEnd = (Get-Date).ToString()
				
				# Do we have ShareGate XML?
				$resultfile = "\\$pc\d$\insanemove\worker$wid-$runAsUser.xml"
				if (Test-Path $resultfile) {
					# Read XML
					$x = $null
					[xml]$x = Get-Content $resultfile
					if ($x) {
						# Parse XML nodes
						$row.SGServer = $pc
						$row.SGResult = ($x.Objs.Obj.Props.S |? {$_.N -eq "Result"})."#text"
						$row.SGSessionId = ($x.Objs.Obj.Props.S |? {$_.N -eq "SessionId"})."#text"
						$row.SGSiteObjectsCopied = ($x.Objs.Obj.Props.I32 |? {$_.N -eq "SiteObjectsCopied"})."#text"
						$row.SGItemsCopied = ($x.Objs.Obj.Props.I32 |? {$_.N -eq "ItemsCopied"})."#text"
						$row.SGWarnings = ($x.Objs.Obj.Props.I32 |? {$_.N -eq "Warnings"})."#text"
						$row.SGErrors = ($x.Objs.Obj.Props.I32 |? {$_.N -eq "Errors"})."#text"
						
						# TaskXML
						$row.TaskXML = $x.OuterXml
						
						# Delete XML
						Remove-Item $resultfile -confirm:$false -ErrorAction SilentlyContinue
					}

					# Error
					$err = ""
					$errcount = 0
					$task.Error |% {
						$err += ($_|ConvertTo-Xml).OuterXml
						$errcount++
					}
					$row.ErrorCount = $errCount
				}
			}
		}
	}
}

Function ExecuteSiteCopy($row, $worker) {
	# Parse fields
	$name = $row.Name
	$srcUrl = $row.SourceURL
	
	# Destination
	if ($row.MySiteEmail) {
		# MySite /personal/
		$destUrl = $row.DestinationURL
	} else {
		# Team /sites/
		$destUrl = FormatCloudMP $row.DestinationURL
	}
	
	# Grant SCA
	$adminUser = $settings.settings.tenant.adminUser
	$adminRole = $settings.settings.tenant.adminRole
	$uploadUser = $worker.UploadUser
	Write-Host "Grant SCA $upn to $destUrl" -Fore Green
	
	# Unlock site collection
	$site = Get-MSPOSite $destUrl
	Set-SPSite -Identity $srcUrl -LockState Unlock
	
	# SPO - Site Collection Admin
	Set-MSPOUser -Site $site -LoginName $adminUser -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue
	Set-MSPOUser -Site $site -LoginName $adminRole -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue
	Set-MSPOUser -Site $site -LoginName $uploadUser -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue
	
	# Make NEW Session - remote PowerShell
    $wid = $worker.Id
    $pc = $worker.PC
	$runAsUser = $worker.RunAsUser
	$s = Get-PSSession |? {$_.ComputerName -eq $pc}
	
	# Generate local secure CloudPW
	$sb = [Scriptblock]::Create("""$global:cloudPW"" | ConvertTo-SecureString -Force -AsPlainText | ConvertFrom-SecureString")
	$localHash = Invoke-Command $sb -Session $s
	
	# Generate PS1 worker script
	$now = (Get-Date).tostring("yyyy-MM-dd_hh-mm-ss")
	if ($incremental) {
		# Team site INCREMENTAL
		$copyparam = "-CopySettings `$csIncr"
	}
	if ($row.MySiteEmail) {
		# MySite /personal/ = always RENAME
		$copyparam = "-CopySettings `$csMysite"

		# Set Version History before migration (ShareGate will copy setting to O365)
		Write-Host "Set Version History 99" -Fore Green
		$site = Get-SPSite $srcUrl
		$docLib = $site.RootWeb.Lists["Documents"]
		$docLib.EnableVersioning = $true
		$docLib.MajorVersionLimit = 99
		$docLib.Update()
	}
	$uploadUser = $worker.UploadUser
	$uploadPass = $settings.settings.tenant.uploadPass
	
	$ps = "Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null`n`$pw='$uploadPass';md ""d:\insanemove\log"" -ErrorAction SilentlyContinue;`nStart-Transcript ""d:\insanemove\log\worker$wid-$runAsUser-$now.log"";`n""uploadUser=$uploadUser"";`n""SOURCE=$srcUrl"";`n""DESTINATION=$destUrl"";`nImport-Module ShareGate;`n`$src=`$null;`n`$dest=`$null;
	`$secpw = ConvertTo-SecureString -String `$pw -AsPlainText -Force;
	`$cred = New-Object System.Management.Automation.PSCredential (""$uploadUser"", `$secpw);
	;`n`$src = Connect-Site ""$srcUrl"";`n`$dest = Connect-Site ""$destUrl"" -Cred `$cred;`nif (`$src.Url -eq `$dest.Url) {`n""SRC""`n`$src|fl`n""DEST""`n`$dest|fl`n`$csMysite = New-CopySettings -OnSiteObjectExists Merge -OnContentItemExists Rename;`n`$csIncr = New-CopySettings -OnSiteObjectExists Merge -OnContentItemExists IncrementalUpdate;`n`$result = Copy-Site -Site `$src -DestinationSite `$dest -Subsites -Merge -InsaneMode -VersionLimit 50;`n`$result | Export-Clixml ""d:\insanemove\worker$wid-$runAsUser.xml"" -Force;`n} else {`n""URLs don't match""`n}`nREMSet-SPSite -Identity ""$srcUrl"" -LockState ReadOnly`nWrite-Host ""Source locked read only""`nGet-SPSite ""$srcUrl"" | Select Url,ReadOnly,*Lock* | Ft -a`nStop-Transcript"
    #OLD ;`n`$src = Connect-Site ""$srcUrl"";`n`$dest = Connect-Site ""$destUrl"" -Cred `$cred;`nif (`$src.Url -eq `$dest.Url) {`n""SRC""`n`$src|fl`n""DEST""`n`$dest|fl`n`$csMysite = New-CopySettings -OnSiteObjectExists Merge -OnContentItemExists Rename;`n`$csIncr = New-CopySettings -OnSiteObjectExists Merge -OnContentItemExists IncrementalUpdate;`n# READ ONLY - Cred current user`n`$rouser = `$env:userdomain + ""\"" + `$env:username`n`$rocred = Get-StoredCredential |? {`$_.UserName -eq `$rouser}`n# READ ONLY - Open PS Session`n`$rosess = New-PSSession -ComputerName `$env:computername -Credential `$rocred -Authentication Credssp`n# READ ONLY - Invoke Delay`n`$rocmd = ""Add-PSSnapin Microsoft.SharePoint.PowerShell`nSleep (5*60)`nSet-SPSite '$srcUrl' -LockState ReadOnly""`n`$rosb = [Scriptblock]::Create(`$rocmd)`n`$rojob = Invoke-Command -ScriptBlock `$rosb -Session `$rosess -AsJob`n`$result = Copy-Site -Site `$src -DestinationSite `$dest -Subsites -Merge -InsaneMode -VersionLimit 50;`n`$result | Export-Clixml ""d:\insanemove\worker$wid-$runAsUser.xml"" -Force;`n} else {`n""URLs don't match""`n}`nREMSet-SPSite -Identity ""$srcUrl"" -LockState ReadOnly`nWrite-Host ""Source locked read only""`nGet-SPSite ""$srcUrl"" | Select Url,ReadOnly,*Lock* | Ft -a`nStop-Transcript"
    # READ ONLY
    # n# READ ONLY - Cred current user`n`$rouser = `$env:userdomain + ""\"" + `$env:username`n`$rocred = Get-StoredCredential |? {`$_.UserName -eq `$rouser}`n# READ ONLY - Open PS Session`n`$rosess = New-PSSession -ComputerName `$env:computername -Credential `$rocred -Authentication Credssp`n# READ ONLY - Invoke Delay`n`$rocmd = ""Add-PSSnapin Microsoft.SharePoint.PowerShell`nSleep (5*60)`nSet-SPSite '$srcUrl' -LockState ReadOnly""`n`$rosb = [Scriptblock]::Create(`$rocmd)`n`$rojob = Invoke-Command -ScriptBlock `$rosb -Session `$rosess -AsJob`



	# Dry run
	if ($dryRun) {
		$ps = $ps.Replace("Copy-Site","NoCopy-Site")
		$ps = $ps.Replace("Set-SPSite","NoSet-SPSite")
	}
    $ps | Out-File "\\$pc\d$\insanemove\worker$wid-$runAsUser.ps1" -Force
    Write-Host $ps -Fore Yellow
	
    # Invoke SCHTASK
    $cmd = "Get-ScheduledTask -TaskName ""worker$wid-$runAsUser"" | Start-ScheduledTask"
	
	# Display
    Write-Host "START worker $wid on $pc" -Fore Green
	Write-Host "$srcUrl,$destUrl" -Fore yellow
	
	# Execute
	$sb = [Scriptblock]::Create($cmd) 
	Invoke-Command $sb -Session $s
}

Function FindCloudMySite ($MySiteEmail) {
	# Lookup /personal/ site URL based on User Principal Name (UPN)
	$coll = @()
	$coll += $MySiteEmail
	$profile = Get-PnPUserProfileProperty -Account $coll
	if ($profile) {
		if ($profile.PersonalUrl) {
			$url = $profile.PersonalUrl.TrimEnd('/')
		}
	}
	Write-Host "SEARCH for $MySiteEmail found URL $url" -Fore Yellow
	return $url
}

Function WriteCSV() {
	"<WriteCSV>"
    # Write new CSV output with detailed results
    $file = $fileCSV.Replace(".csv", "-results.csv")
    $global:track | Select SourceURL,DestinationURL,MySiteEmail,CsvID,WorkerID,PC,RunAsUser,Status,SGResult,SGServer,SGSessionId,SGSiteObjectsCopied,SGItemsCopied,SGWarnings,SGErrors,Error,ErrorCount,TaskXML,SPStorage | Export-Csv $file -NoTypeInformation -Force -ErrorAction Continue
}

Function CopySites() {
	"<CopySites>"
	# Monitor and Run loop
	Write-Host "===== Start Site Copy to O365 ===== $(Get-Date)" -Fore Yellow
	CreateTracker
	
	# Safety
	if (!$global:workers) {
		Write-Host "No Workers Found" -Fore Red
		return
	}
	
	$csvCounter = 0
	$emailCounter = 0
	do {
		$csvCounter++
		$emailCounter++
		# Get latest Job status
		UpdateTracker
		Write-Host "." -NoNewline
		
		# Ensure all sessions are active
		foreach ($worker in $global:workers) {
			# Count active sessions per server
			$wid = $worker.Id
			$active = $global:track |? {$_.Status -eq "InProgress" -and $_.WorkerID -eq $wid}
            
			if (!$active) {
				# Available session.  Assign new work
				Write-Host " -- AVAIL" -Fore Green
				Write-Host $wid -Fore Yellow
				$row = $global:track |? {$_.Status -eq "New"}
			
                if ($row) {
					# First row only, no array
                    if ($row -is [Array]) {
                        $row = $row[0]
                    }
					
					# Update DB tracking
					$row.WorkerID = $wid
					$row.PC = $global:workers[$wid].PC
					$row.RunAsUser = $global:workers[$wid].RunAsUser
				    $row.Status = "InProgress"
					$row.TimeCopyStart = (Get-Date).ToString()
					
					# Display
					$row |ft -a

                    # Kick off copy
					Start-Sleep 5
					"sleep 5 sec..."
				    ExecuteSiteCopy $row $worker				    
                }
			} else {
				Write-Host " -- NO AVAIL" -Fore Green
			}
				
			# Progress bar %
			$complete = ($global:track |? {$_.Status -eq "Completed"}).Count
			$total = $global:track.Count
			$prct = [Math]::Round(($complete/$total)*100)
			
			# ETA
			if ($prct) {
				$elapsed = (Get-Date) - $start
				$remain = ($elapsed.TotalSeconds) / ($prct / 100.0)
				$eta = (Get-Date).AddSeconds($remain - $elapsed.TotalSeconds)

				# Display
				Write-Progress -Activity "Copy site - ETA $eta" -Status "$name ($prct %)" -PercentComplete $prct
			}

			# Progress table
			"[TRACK]"
			$wip = $global:track |? {$_.Status -eq "InProgress"} | select CsvID,WorkerID,PC,RunAsUser,SourceURL,DestinationURL 
			$wip | ft -a
			$wip = $wip | Out-String
			
			$grp = $global:track | group Status
			$grp | select Count,Name | sort Name | ft -a
		}
		
		# Write CSV with partial results.  Enables monitoring long runs.
		if ($csvCounter -gt 5) {
			WriteCSV
			$csvCounter = 0
		}
		
		# Progress table
		# 5 sec space, 12 per min, 15 minute spacing
		$summary = $grp | select Count,Name | sort Name | Out-String
		if ($emailCounter -gt (12 * 15)) {
			EmailSummary
			$emailCounter = 0
		}

		# Latest counter
		$remain = $global:track |? {$_.status -ne "Completed" -and $_.status -ne "Failed"}
		"Sleep 5 sec..."
		Start-Sleep 5
	} while ($remain)
	
	# Complete
	Write-Host "===== Finish Site Copy to O365 ===== $(Get-Date)" -Fore Yellow
	"[TRACK]"
	$global:track | group status | ft -a
	$global:track | select CsvID,JobID,SessionID,SGSessionId,PC,RunAsUser,SourceURL,DestinationURL | ft -a
}

Function EmailSummary ($style) {
	# Email config
	$smtpServer = $settings.settings.notify.smtpServer
	$from = $settings.settings.notify.from
	$to = $settings.settings.notify.to

	# Done
	if (!$prct) {$style = "done"}
	if ($style -eq "done") {
		$prct = "100"
		$eta = "done"
		$summary = "--DONE-- "
	}
	
	# Send message
	if ($smtpServer -and $to -and $from) {
		$summary = $grp | select Count,Name | sort Name | Out-String
		Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -Subject "Copy Site ($prct %) - ETA $eta - $name" -Body "$summary <br/> $wip" -BodyAsHtml
	}
}

Function VerifyCloudSites() {
	"<VerifyCloudSites>"
	# Read CSV and ensure cloud sites exists for each row
	Write-Host "===== Verify Site Collections exist in O365 ===== $(Get-Date)" -Fore Yellow
	$global:collMySiteEmail = @()

	
	# Loop CSV
	$csv = Import-Csv $fileCSV
	foreach ($row in $csv) {
		$row | ft -a
		if ($row.MySiteEmail) {
			# MySite
			$global:collMySiteEmail += $row.MySiteEmail
		} else {
			# Team Site
			EnsureCloudSite $row.SourceURL $row.DestinationURL $row.MySiteEmail
		}
	}
	
	# Execute creation of OneDrive /personal/ sites in batches (200 each) https://technet.microsoft.com/en-us/library/dn792367.aspx
	if ($global:collMySiteEmail) {
		Write-Host " - PROCESS MySite bulk creation"
	}
	$i = 0
	$batch = @()
	foreach ($MySiteEmail in $global:collMySiteEmail) {
		if ($i -lt 199) {
			# append batch
			$batch += $MySiteEmail
			Write-Host "." -NoNewline
		} else {
			$batch += $MySiteEmail
			BulkCreateMysite $batch
			$i = 0
			$batch = @()
		}
		$i++
	}
	if ($batch.count) {
		BulkCreateMysite $batch
	}
	Write-Host "OK"
}

Function BulkCreateMysite ($batch) {
	"<BulkCreateMysite>"
	# execute and clear batch
	Write-Host "`nBATCH New-PnPPersonalSite $($batch.count)" -Fore Green
	$batch
	$batch.length
	New-PnPPersonalSite -Email $batch
}

Function EnsureCloudSite($srcUrl, $destUrl, $MySiteEmail) {
	"<EnsureCloudSite>"
	# Create site in O365 if does not exist
	$destUrl = FormatCloudMP $destUrl
	Write-Host $destUrl -Fore Yellow
	$srcUrl
	if ($srcUrl) {
		$site = Get-SPSite $srcUrl
		$web = $site.RootWeb
		if ($web.RequestAccessEmail) {
			#REM $upn = $web.RequestAccessEmail.Split(",;")[0].Split("@")[0] + "@fanniemae.com"; #REM + $settings.settings.tenant.suffix;
			$upn = $settings.settings.tenant.adminUser
		}
		if (!$upn) {
			$upn = $settings.settings.tenant.adminUser
		}
	}
	
	# Verify Site
	try {
		if ($destUrl) {
			$cloud = Get-PnPTenantSite -Url $destUrl -ErrorAction SilentlyContinue
		}
	} catch {}
	if (!$cloud) {
		Write-Host "- CREATING $destUrl"
		
		if ($MySiteEmail) {
			# Provision MYSITE
			$global:collMySiteEmail += $MySiteEmail
		} else {
			# Provision TEAMSITE
			$quota = 1024*50
			$splits = $destUrl.split("/")
			$title = $splits[$splits.length-1]
			New-PnPTenantSite -Owner $upn -Url $destUrl -StorageQuota $quota -Title $title -TimeZone 10
		}
	} else {
		Write-Host "- FOUND $destUrl"
	}
}

Function FormatCloudMP($url) {
	# Replace Managed Path with O365 /sites/ only
	if (!$url) {return}
	$managedPath = "sites"
	$i = $url.Indexof("://")+3
	$split = $url.SubString($i, $url.length-$i).Split("/")
	$split[1] = $managedPath
	$final = ($url.SubString(0,$i) + ($split -join "/")).Replace("http:","https:")
	return $final
}

Function ConnectCloud {
	"<ConnectCloud>"
	# Prepare
	$pw = $global:cloudPW
	$pw
	$settings.settings.tenant.adminUser
	$secpw = ConvertTo-SecureString -String $pw -AsPlainText -Force
	$c = New-Object System.Management.Automation.PSCredential ($settings.settings.tenant.adminUser, $secpw)
	
	# Connect PNP
	#Connect-PnpOnline -URL $settings.settings.tenant.adminURL -Credential $c
	#Connect-MSPOService -URL $settings.settings.tenant.adminURL -Credential $c
	Connect-PnpOnline -URL https://fnma-admin.sharepoint.com -Credential $c
	Connect-MSPOService -URL https://fnma-admin.sharepoint.com -Credential $c
	}

Function MeasureSiteCSV {
	"<MeasureSiteCSV>"
	# Populate CSV with local farm SharePoint site collection size
	$csv = Import-Csv $fileCSV
	foreach ($row in $csv) {
		$s = Get-SPSite $row.SourceURL
		if ($s) {
			$storage = [Math]::Round($s.Usage.Storage / 1MB, 2)
            if (!($row.PSObject.Properties.name -contains "SPStorage")) {
				# add property SPStorage to collection, if missing
                $row | Add-Member –MemberType NoteProperty –Name SPStorage –Value ""
            }
			$row.SPStorage = $storage
		}
	}
	$csv | Export-Csv $fileCSV -Force -NoTypeInformation
}

Function LockSite($lock) {
	"<LockSite>"
	# Modfiy on-prem site collection lock
	Write-Host $lock -Fore Yellow
	$csv = Import-Csv $fileCSV
	foreach ($row in $csv) {
		$url = $row.SourceURL
		Set-SPSite $url -LockState $lock
		"[SPSITE]"
		Get-SPSite $url | Select URL,*Lock* | ft -a
	}
}

Function CompileAudiences() {
	# Find all local Audiences
	$AUDIENCEJOB_START       = '1'
	$AUDIENCEJOB_INCREMENTAL = '0'
	$site          = (Get-SPSite)[0]
	$context       = Get-SPServiceContext $site  
	$proxy         = $context.GetDefaultProxy([Microsoft.Office.Server.Audience.AudienceJob].Assembly.GetType('Microsoft.Office.Server.Administration.UserProfileApplicationProxy'))
	$applicationId = $proxy.GetType().GetProperty('UserProfileApplication', [System.Reflection.BindingFlags]'NonPublic, Instance').GetValue($proxy, $null).Id.Guid
	$auManager     = New-Object Microsoft.Office.Server.Audience.AudienceManager $context
	$auManager.Audiences | Sort-Object AudienceName |% {
		# Compile each Audience
		$an = $_.AudienceName
		$an
		[Microsoft.Office.Server.Audience.AudienceJob]::RunAudienceJob(@($applicationId, $AUDIENCEJOB_START, $AUDIENCEJOB_INCREMENTAL, $an))
	}
}

Function UserProfileSetHybridURL() {
	# UPS Manager
	$site = (Get-SPSite)[0]
	$context = Get-SPServiceContext $site
	$profileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($context)
	
	# MySite Host URL
	$myhost =  $settings.settings.tenant.adminURL.replace("-admin","-my")
	if (!$myhost.EndsWith("/")) {$myhost += "/"}
	
	# Loop CSV
	$csv = Import-Csv $fileCSV
	foreach ($row in $csv) {
		$login = $row.MySiteEmail.Split("@")[0]
		$p = $profileManager.GetUserProfile($login)
		if ($p) {
			# User Found
			$dest = FindCloudMySite $row.MySiteEmail
			if (!$dest.EndsWith("/")) {$dest += "/"}
			$dest = $dest.Replace($myhost,"/")
			
			# Update Properties - drives URL redirect Audience
			Write-Host "SET UPS for $login to $dest"
			$p["PersonalSpace"].Value = $dest
			$p.Commit()
		}
	}
}

Function Clean() {
	Write-Host "<Clean>"
	DetectVendor
	foreach ($server in $global:servers) {
		# File system
		Write-Host " - File system"
		$pc = $server
		Remove-Item "\\$pc\d$\insanemove\worker*.*" -Confirm:$false -Force

		# User accounts
		$runasuser = @()
		$runasuser += $env:username
		if ($settings.settings.optionalSchtaskUser) {
			$runasuser += $settings.settings.optionalSchtaskUser
		}

		# Scheduled Task
		Write-Host " - Scheduled Task"
		1..$settings.maxWorker |% {
			$i = $_
			foreach ($user in $runasuser) {
				$taskname = "worker1-[RUNASUSER]"
				$taskname = $taskname.Replace("1", $i).Replace("[RUNASUSER]", $user)
				$cmd = "schtasks.exe /delete /s $pc /tn $taskname /F"
				Invoke-Expression $cmd
			}
		}

		# Stop ShareGate EXE running
		Write-Host " - ShareGate EXE Running"
		$proc = Get-WmiObject Win32_Process -ComputerName $server |? {$_.ProcessName -match "Sharegate"}
		$proc |% {$_.Terminate()}

		# ShareGate Application Cache
		Write-Host " - ShareGate Application Cache"
		foreach ($user in $runasuser) {
			$folder = "C:\Users\[USER]\AppData\Local\Sharegate\ApplicationLogs".Replace("[USER]", $user)
			$folder
			Remove-Item $folder -Confirm:$false -Recurse -Force -ErrorAction SilentlyContinue
			$folder = "C:\Users\[USER]\AppData\Local\Sharegate\Sharegate.Migration.txt".Replace("[USER]", $user)
			$folder
			Remove-Item $folder -Confirm:$false -Force -ErrorAction SilentlyContinue
		}
	}
}
Function Main() {
	"<Main>"
	# Clean
	if ($clean) {
		Clean
		Exit
	}

	# Start LOG
	$start = Get-Date
	$when = $start.ToString("yyyy-MM-dd-hh-mm-ss")
	$logFile = "$root\log\InsaneMove-$when.txt"
	mkdir "$root\log" -ErrorAction SilentlyContinue | Out-Null
	if (!$psISE) {
		try {
			Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
		} catch {}
		Start-Transcript $logFile
	}
	Write-Host "fileCSV = $fileCSV"

	# Core logic
	if ($userProfileSetHybridURL) {
		# Update local user profiles.  Set Personal site URL for Hybrid OneDrive audience compilation and redirect
		ReadCloudPW
		ConnectCloud
		UserProfileSetHybridURL
		CompileAudiences
	} elseif ($measure) {
		# Populate CSV with size (GB)
		MeasureSiteCSV
	} elseif ($readOnly) {
		# Lock on-prem sites
		LockSite "ReadOnly"
	} elseif ($readWrite) {
		# Unlock on-prem sites
		LockSite "Unlock"
	} elseif ($noAccess) {
		# NoAccess on-prem sites
		LockSite "NoAccess"	
	} else {
		if ($verifyCloudSites) {
			# Create site collection
			ReadCloudPW
			ConnectCloud
			VerifyCloudSites
		} else {
			if (!$dryRun) {
				# Prompt to verify
				$continue = $false
				Write-Host "Do you want to continue? (Y/N)" -Fore Yellow
				$choice = Read-Host
				if ($choice -like "y*") {
					$continue = $true
				} else {
					Write-Host "HALT - User did not confirm data copy." -Fore Red
				}
			}
			if ($dryRun -or $continue) {
				# Copy site content
				VerifyPSRemoting
				ReadIISPW
				ReadCloudPW
				ConnectCloud
				DetectVendor
				CloseSession
				CreateWorkers
				CopySites
				EmailSummary "done"
				CloseSession
				WriteCSV
			}
		}
	}
	
	# Finish LOG
	Write-Host "===== DONE ===== $(Get-Date)" -Fore Yellow
	$th				= [Math]::Round(((Get-Date) - $start).TotalHours, 2)
	$attemptMb		= ($global:track | measure SPStorage -Sum).Sum
	$actualMb		= ($global:track |? {$_.SGSessionId -ne ""} | measure SPStorage -Sum).Sum
	$actualSites	= ($global:track |? {$_.SGSessionId -ne ""}).Count
	Write-Host ("Duration Hours              : {0:N2}" -f $th) -Fore Yellow
	Write-Host ("Total Sites Attempted       : {0}" -f $($global:track.count)) -Fore Green
	Write-Host ("Total Sites Copied          : {0}" -f $actualSites) -Fore Green
	Write-Host ("Total Storage Attempted (MB): {0:N0}" -f $attemptMb) -Fore Green
	Write-Host ("Total Storage Copied (MB)   : {0:N0}" -f $actualMb) -Fore Green
	Write-Host ("Total Objects               : {0:N0}" -f $(($global:track | measure SGItemsCopied -Sum).Sum)) -Fore Green
	Write-Host ("Total Servers               : {0}" -f $global:servers.Count) -Fore Green
	Write-Host ("Total Worker Threads        : {0}" -f $maxWorker) -Fore Green
	Write-Host "====="  -Fore Yellow
	Write-Host ("GB per Hour                 : {0:N2}" -f (($actualMb/1KB)/$th)) -Fore Green
	
	
	$zeroItems = $global:track |? {$_.SGItemsCopied -eq 0}
	if ($zeroItems) {
		Write-Host ("Sites with zero items       : {0}" -f $zeroItems.Length) -Fore Red
		$zeroItems | ft -a
	}

	Write-Host $fileCSV
	if (!$psISE) {Stop-Transcript}
}
Main