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
	[switch]$userProfile = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Dry run replaces core "Copy-Site" with "NoCopy-Site" to execute all queueing but not transfer any data.')]
	[Alias("d")]
	[switch]$dryRun = $false,

	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Clean servers to preprae for next migration batch.')]
	[Alias("c")]
	[switch]$clean = $false,

	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Delete source SharePoint sites on-premise.')]
	[Alias("ds")]
	[switch]$deleteSource = $false,

	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Delete destination SharePoint sites in cloud.')]
	[Alias("dd")]
	[switch]$deleteDest = $false,

	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Compare source and destination lists for QA check.')]
	[Alias("qa")]
	[switch]$qualityAssurance = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Copy sites to Office 365.  Main method.')]
	[Alias("mig")]
	[switch]$migrate = $false,
	
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='Pre-Migration Report.  Runs Copy-Site with -WhatIf')]
	[switch]$whatif = $false
)

# Plugin
Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module SharePointPnPPowerShellOnline -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
Import-Module CredentialManager -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

# Config
$root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[xml]$settings = Get-Content "$root\InsaneMove.xml"
$maxWorker = $settings.settings.maxWorker

# Quality Assurance
Function InspectSource($url) {
    # Connect source SPSite
	# SOM - Server Object Model OK
    write-host -Fore Yellow "Source:" $url
    $site = Get-SPSite $url
	
	# Empty collection
    $sourceLists = @()

	# Display
    Write-Host "`n*************  $($site.RootWeb.Title) *************`n" -BackgroundColor DarkGray

	# Loop all webs
    foreach ($web in $site.AllWebs) {
        foreach ($list in $web.Lists) {
			# Format URL
            $source = $list.RootFolder.ServerRelativeUrl
            $SourceEndChar = $source.IndexOf("-sites/")
            if ($SourceEndChar -gt 0) {
                $source = $source.remove(1,$SourceEndChar)
			}
			
			# Detail object
			$listObj = New-Object -TypeName PSObject -Prop (@{
				"List"=$source;
				"ItemCount"=$list.ItemCount;
				"ListTitle"=$list.Title;
				"Web"=$web.Url;
				"ServerRelativeUrl"=$web.ServerRelativeUrl;
				"BaseType"=$list.BaseType;
			})
			# Append to collection
            $sourceLists += $listObj
		}
		# Memory leak
        $web.Dispose() | Out-Null
    }
	return $sourceLists 
	# Memory leak
    $site.Dispose() | Out-Null
}

Function InspectDestination($url) {
    # Connect to destination
	write-host -Fore Yellow "Destination:" $url
    Connect-PnPOnline -Url $url -Credentials $global:cloudcred | Out-Null
	
	# Open web and lists
    $rootLists = Get-PnPList
    $rootWeb = Get-PnPWeb
    $subWebs = Get-PnPSubWebs -Recurse

	# Open subweb lists
    foreach ($web in $subWebs) {
        $subLists += Get-PnPList -Web $web
    } 
	
	# Empty collections
    $global:allWebs = @()
    $global:allWebs += $subWebs 
    $global:allWebs += $rootWeb

    $global:allLists = @()
    $global:allLists += $rootLists
    $global:allLists += $subLists
        
    $destinationLists = @()
	
	# Loop all lists
    foreach ($list in $allLists) {
		# Collect detail
		$listObj = New-Object -TypeName PSObject -Prop (@{
			"List"              = $list.RootFolder.ServerRelativeUrl.Replace('sites/490-','490-sites/').Replace('sites/ba_d-','ba_d-sites/').Replace('sites/cei-','cei-sites/').Replace('sites/comm-','comm-sites/').Replace('sites/corpadmin-','corpadmin-sites/').Replace('sites/dro-','dro-sites/').Replace('sites/dw-','dw-sites/').Replace('sites/ebusiness-','ebusiness-sites/').Replace('sites/entops-','entops-sites/').Replace('sites/eso-','eso-sites/').Replace('sites/exec-','exec-sites/').Replace('sites/finance-','finance-sites/').Replace('sites/hcd-','hcd-sites/').Replace('sites/hr-','hr-sites/').Replace('sites/lawpolicy-','lawpolicy-sites/').Replace('sites/portfolio-','portfolio-sites/').Replace('sites/ro-','ro-sites/').Replace('sites/sfmb-','sfmb-sites/').Replace('sites/sox-','sox-sites/');
			"ItemCount"=$list.ItemCount;
			"ListTitle"=$list.Title;
			"Web"=$list.Context.url;
			"ServerRelativeUrl" = $list.ParentWebUrl.Replace('sites/490-','490-sites/').Replace('sites/ba_d-','ba_d-sites/').Replace('sites/cei-','cei-sites/').Replace('sites/comm-','comm-sites/').Replace('sites/corpadmin-','corpadmin-sites/').Replace('sites/dro-','dro-sites/').Replace('sites/dw-','dw-sites/').Replace('sites/ebusiness-','ebusiness-sites/').Replace('sites/entops-','entops-sites/').Replace('sites/eso-','eso-sites/').Replace('sites/exec-','exec-sites/').Replace('sites/finance-','finance-sites/').Replace('sites/hcd-','hcd-sites/').Replace('sites/hr-','hr-sites/').Replace('sites/lawpolicy-','lawpolicy-sites/').Replace('sites/portfolio-','portfolio-sites/').Replace('sites/ro-','ro-sites/').Replace('sites/sfmb-','sfmb-sites/').Replace('sites/sox-','sox-sites/');
			"BaseType"=$list.BaseType;
		})
		Write-Host $listObj -Fore Yellow
        $destinationLists += $listObj
    }
	
    return $destinationLists   
}

Function CompareSites($sourceLists, $destinationLists) {
	
	# Unique
    $sourceLists = $sourceList | select -Unique
    $destinationLists = $destinationLists | select -Unique		

	# Empty collection
    $missingItems = @()
	$missingList = @()
	
	# Define exclusion
    $excludeLists = "Community Members", "Style Library", "Content and Structure Reports", "wfsvc","Converted Forms", "Workflow History", "Long Running Operation Status", "Access Requests", "Reporting Metadata", "Reporting Templates", "Workflows", "MicroFeed"
	
	# Inspect source lists
    foreach ($list in $sourceLists) {
		Write-Host $list -Fore White 
		# Define exclusion
        $skipList = $excludeLists -contains $list.ListTitle
		$skipPath = $list.List -like "*_catalogs*"
		
        # Exlcude filter
        if (!$skipList -and !$skipPath) {
            $match = $destinationLists |? { $_.List -eq $list.List }
            $sList = $list.List
            $sTitle = $list.ListTitle
            $sCount = $list.ItemCount
            $dList = $match.List
            $dCount = $match.ItemCount
            $diff = $sCount - $dCount
            
            if ($match) {
                if ($diff -gt 0) {
					# Missing list item
                    $baseType = $match.BaseType
                    Write-Host "$sList ($sCount) - $dList ($dCount) - Missing $diff items ($baseType)" -ForegroundColor Yellow
                    $itemText = "$sList ($sCount) - $dList ($dCount) - Missing $diff items ($baseType)"
                    $missingItems += $itemText 
                    $sWeb = Get-SPWeb $list.Web                   
                    $dWeb = $global:allWebs |? { $_.ServerRelativeUrl -eq $match.ServerRelativeUrl }

					# Document Library
                    if ($match.BaseType -eq "DocumentLibrary") {
                        # Fetch source and destination
                        $sDocs = $sWeb.Lists["$sTitle"].items
                        $dDocs = Get-PNPListItem -Web $dWeb.ServerRelativeUrl -List $match.ListTitle -Fields "FileRef" | % {New-Object PSObject -Property  @{FileRef = $_["FileRef"]}}

						# Compare each document
                        foreach ($doc in $sDocs) {
							# Format URL
                            $docUrl = $doc.File.ServerRelativeUrl
                            $docEndChar = $docUrl.IndexOf("-sites/")
                            if ($docEndChar -gt 0) {
                                $docUrl = $docUrl.remove(1,$docEndChar)
                            }

							# Do we have a match?
                            $found = $dDocs |? { $_.FileRef -eq $docUrl }
                            if (!$found) {
                                Write-Host "    - $docUrl"
                            }
                        }
                    }
                }
            } else {
				# No match found.  Missing list.
                Write-Host "$sList ($sCount) - No Match Found" -ForegroundColor Red
                $listText = $sList
                $missingList += $listText
            }
        }
    }

	# Result text
    if ($missingItems -or $missingList) {
        $result = "Fail"
    } else {
        $result = "Pass"
        Write-Host "Site migrated successfully" -BackgroundColor DarkGreen
    }

	# Return data object
	$reportData = New-Object -TypeName PSObject -Prop (@{
		"Result"=$result;
		"Lists w Missing Items"= $missingItems.Count;
		"Missing List"=$missingList.Count
	})

    return $reportData
}

Function SaveQACSV($report) {
    # Save CSV report
    $datestamp = (Get-Date).tostring("yyyy-MM-dd-hh-mm-ss")
	$file = $fileCSV.Replace(".csv", "-qa-report-$datestamp.csv")
    $report | Select "SourceURL","DestinationURL","Result","Lists w Missing Items","Missing Lists" | Export-Csv $file -NoTypeInformation
	Write-Host "Saved : $file" -Fore Green
	return $file
}

Function EmailQACSV($file) {
	# Email config
	$smtpServer = $settings.settings.notify.smtpServer
	$from = $settings.settings.notify.from
	$to = $settings.settings.notify.to

	# Stop and attach transcript
	$log = Stop-Transcript
	$log = $log.Replace("Transcript stopped, output file is ","")

	# Attachment collection
	$attach = @()
	$attach += $log
	$attach += $file

	# Attach CSV and send
	Send-MailMessage -SmtpServer $smtpServer -From $from -To $to -Subject "$file" -Body "Attached CSV" -Attachments $attach -BodyAsHtml
}

Function QualityAssurance() {
	$sites = Import-CSV $fileCSV
	$report = @()
	foreach ($s in $sites) {
        write-host -Fore Yellow $s 
        # Get on-premise site lists
        $sourceLists = InspectSource $s.SourceURL
		Write-Host $sourceLists.List -Fore Cyan
        # Get O365 site lists
        $destinationLists = InspectDestination $s.DestinationURL
		Write-Host $destinationLists.List -Fore Magenta
        # Compare source/destination
        $reportData = CompareSites $sourceLists $destinationLists
		Write-Host $reportData -Fore Cyan
		$obj = New-Object -TypeName PSObject -Prop (@{
			"SourceURL"=$s.SourceURL;
			"DestinationURL"=$s.DestinationURL;
			"Result"=$reportData.("Result");
			"Lists w Missing Items"=$reportData.("Lists w Missing Items");
			"Missing Lists"=$reportData.("Missing Lists")
		})
		$report += $obj
	}
	# Write QA report
	$file = SaveQACSV $report
	return $file
}

Function DeleteSourceSites() {
	$sites = Import-CSV $fileCSV
	foreach ($s in $sites) {
		Write-Host "Delete source site $($s.SourceURL)"
 		# SOM - Server Object Model OK
		Remove-SPSite $s.SourceURL -Confirm:$false -GradualDelete
	}
}

Function DeleteDestinationSites() {
	$sites = Import-CSV $fileCSV
	foreach ($s in $sites) {
		$u = FormatCloudMP $s.DestinationURL
		Write-Host "Delete destination site $u"
		Remove-PnPTenantSite $u -Confirm:$false -Force
	}
}

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
	$global:farmcred = New-Object System.Management.Automation.PSCredential -ArgumentList "$userdomain\$username", $sec
}

Function DetectVendor() {
	"<DetectVendor>"
	# SharePoint Servers in local farm
	$spservers = Get-SPServer |? {$_.Role -ne "Invalid"} | sort Address
	$settings.settings.optionalLimitServers.Split(",") | % {
        $obj = New-Object -TypeName PSObject -Prop (@{"Address" = $_})
        $servers += $obj
    }

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
		write-host -fore cyan "new session $pc"
		$s = New-PSSession -ComputerName $pc -Credential $global:farmcred -Authentication CredSSP # -ErrorAction SilentlyContinue
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
			$worker = New-Object -TypeName PSObject -Prop (@{
				"Id" = $wid;
				"PC" = $pc;
				"RunAsUser" = $runAsUser;
				"UploadUser" = $uploadUsers[$wid]
			})
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

Function CreateTracker() {
	"<CreateTracker>"
	# CSV migration source/destination URL
	Write-Host "===== Populate Tracking table ===== $(Get-Date)" -Fore Yellow

	$global:track = @()
	$csv = Import-Csv $fileCSV
	$i = 0	
	foreach ($row in $csv) {		
        # Get SharePoint total storage (remotely)
        $sb = {
            param($siteUrl)
            Add-PSSnapin Microsoft.SharePoint.PowerShell
            $site = Get-SPSite $siteUrl
            if ($site) {
                return [Math]::Round($site.Usage.Storage / 1MB, 2)
            }
		}
		if (!$global:remoteFarm) {
			OpenRemoteFarm
		}
        $SPStorage = Invoke-Command -ScriptBlock $sb -Session $global:remoteFarm -ArgumentList $row.SourceURL
		
		# MySite URL Lookup
		if ($row.MySiteEmail) {
			$destUrl = FindCloudMySite $row.MySiteEmail
		} else {
			$destUrl = FormatCloudMP $row.DestinationURL;
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
			"SPStorageMB"=$SPStorage;
			"TimeCopyStart"="";
			"TimeCopyEnd"="";
			"Durationhours"=""
		})
		$global:track += $obj

		# Increment ID
		$i++
	}
	
	# Sort by SharePoint site storage (GB) ascending (small sites first)
    $global:track = $global:track | sort SPStorageMB
	
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
					New-PSSession -ComputerName $brokenCurrent.ComputerName -Credential $global:farmcred -Authentication CredSSP -ErrorAction SilentlyContinue
					$brokenCurrent | Remove-PSSession
				}
			} else {
				# Single
				New-PSSession -ComputerName $broken.ComputerName -Credential $global:farmcred -Authentication CredSSP -ErrorAction SilentlyContinue
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
				$row.Durationhours = ([datetime]($row.TimeCopyEnd) - [datetime]($row.TimeCopyStart)).TotalHours
				
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
	$site = Get-SPOSite $destUrl
	Set-SPSite -Identity $srcUrl -LockState Unlock
	
	# SPO - Site Collection Admin
	Set-SPOUser -Site $site -LoginName $adminUser -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue
	Set-SPOUser -Site $site -LoginName $adminRole -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue
	Set-SPOUser -Site $site -LoginName $uploadUser -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue
	
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
	`n`$src = Connect-Site ""$srcUrl"";`n`$dest = Connect-Site ""$destUrl"" -Cred `$cred;
	`nif (`$src.Url -eq `$dest.Url) {`n""SRC""`n`$src|fl`n""DEST""`n`$dest|fl
	`n`$csMysite = New-CopySettings -OnSiteObjectExists Merge -OnContentItemExists Rename;
	`n`$csIncr = New-CopySettings -OnSiteObjectExists Merge -OnContentItemExists IncrementalUpdate;
	`n`$result = Copy-Site -Site `$src -DestinationSite `$dest -Subsites -Merge -InsaneMode -VersionLimit 50;
	`n`$result | Export-Clixml ""D:\insanemove\worker$wid-$runAsUser.xml"" -Force;`n} else {`n""URLs don't match""`n}
	`nStop-Transcript"

	# What If
	if ($whatif) {
		$ps = $ps.Replace("Copy-Site ","Copy-Site -WhatIf ")
	}

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

Function SaveMigrationCSV() {
	"<SaveMigrationCSV>"
    # Write new CSV output with detailed results
    $file = $fileCSV.Replace(".csv", "-migration-results.csv")
    $global:track | Select SourceURL,DestinationURL,MySiteEmail,CsvID,WorkerID,PC,RunAsUser,Status,SGResult,SGServer,SGSessionId,SGSiteObjectsCopied,SGItemsCopied,SGWarnings,SGErrors,Error,ErrorCount,TaskXML,SPStorageMB,TimeCopyStart,TimeCopyEnd,Durationhours | Export-Csv $file -NoTypeInformation -Force -ErrorAction Continue
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

					# Check in Files
					CheckInDocs $row.SourceURL					

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
			if (!$total) {$total = 1}
			$prct = [Math]::Round(($complete/$total)*100)
			
			# ETA
			if ($prct) {
				$elapsed = (Get-Date) - $global:start
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
			Write-Host (Get-Date)
		}
		
		# Write CSV with partial results.  Enables monitoring long runs.
		if ($csvCounter -gt 5) {
			SaveMigrationCSV
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
	$verifiedCsv = @()
	$csv = Import-Csv $fileCSV
	foreach ($row in $csv) {
		$row | ft -a
		if ($row.MySiteEmail) {
			# MySite
			$global:collMySiteEmail += $row.MySiteEmail
		} else {
            $siteFound = $null
			
            # Detect if Site Collection exists (remotely)
            $sb = {
                param($siteUrl)
                Add-PSSnapin Microsoft.SharePoint.PowerShell
                return Get-SPSite $siteUrl
			}
			if (!$global:remoteFarm) {
				OpenRemoteFarm
			}
            $siteFound = Invoke-Command -ScriptBlock $sb -Session $global:remoteFarm -ArgumentList $row.SourceURL

            if ($siteFound) {
                # Team Site
                EnsureCloudSite $row.SourceURL $row.DestinationURL $row.MySiteEmail
                $verifiedCsv += $row
            }
		}
	}

	# Write Verified CSV
	$verifiedFileCSV = $fileCSV.Replace(".csv", "-verified.csv")
	$verifiedCsv | Export-Csv $verifiedFileCSV -NoTypeInformation -Force
	
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
		$upn = $settings.settings.tenant.adminUser
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
			# Get Site Collection quota max (remotely)
            $sb = {
                param($siteUrl)
                Add-PSSnapin Microsoft.SharePoint.PowerShell
                $site = Get-SPSite $siteUrl
                if ($site) {
                    return $site.Quota.StorageMaximumLevel / 1MB
                }
			}
			if (!$global:remoteFarm) {
				OpenRemoteFarm
			}
            $quota = Invoke-Command -ScriptBlock $sb -Session $global:remoteFarm -ArgumentList $srcUrl

			# Provision TEAMSITE
			$sourceSite = Get-SPSite $srcUrl
			$quota = $sourceSite.Quota.StorageMaximumLevel / 1MB
			$quotaWarn = $quota * 0.90
			$splits = $destUrl.split("/")
			$title = $splits[$splits.length-1]
			Write-Host "New-PnPTenantSite"
			Write-Host "Source URL: $srcUrl"
			Write-Host "Quota (MB): $quota"
			Write-Host "Quota Warn (90%): $quotaWarn"

			# Create site
			Write-Host "Creating site collection $destUrl"
			New-PnPTenantSite -Owner $upn -Url $destUrl -StorageQuota $quota -StorageQuotaWarningLevel $quotaWarn -Title $title -TimeZone 10

			# Wait for site online
			$i = 0
			do {
				$site = Get-SPOSite $destUrl
				Start-Sleep 2
				Write-Host "." -NoNewline
				$i++
			} while (!$site -and ($i -lt 30))
		
			# Appy storage warning
			Set-SPOSite $destUrl -StorageQuotaWarningLevel $quotaWarn
			Get-SPOSite $destUrl | select Storage* | fl

			# Detect Wiki libraries
			$srcWikis = @()
			foreach ($web in $sourceSite.allwebs) {
				  foreach ($l in $web.lists) {
						if ($l.BaseTemplate -eq "WebPageLibrary") {
							  $row = New-Object PSObject -Property @{ 
							  "ParentWebUrl" = $l.ParentWebUrl;
							  "RootFolder" = $l.RootFolder;
							  "Title" = $l.Title
							  }
						$srcWikis += $row
						}
				  }
			}

	  		# Create Wiki libraries
			foreach ($w in $srcWikis) {
				Write-Host "Creating Wiki $($w.Title) on $($w.ParentWebUrl)"
				Connect-PNPOnline -Url $destUrl -Credentials $cred;
				New-PNPList -Title $w.Title -Url $w.RootFolder -Template "WebPageLibrary" -EnableVersioning -OnQuickLaunch -Web $l.ParentWebUrl
			}
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
	$settings.settings.tenant.adminUser
	$secpw = ConvertTo-SecureString -String $pw -AsPlainText -Force
	$global:cloudcred = New-Object System.Management.Automation.PSCredential ($settings.settings.tenant.adminUser, $secpw)
	
	# Connect PNP
	Connect-PnpOnline -URL $settings.settings.tenant.adminURL -Credential $global:cloudcred
	Connect-SPOService -URL $settings.settings.tenant.adminURL -Credential $global:cloudcred
}

Function MeasureSiteCSV {
	"<MeasureSiteCSV>"
	# Populate CSV with local farm SharePoint site collection size
	$csv = Import-Csv $fileCSV
	foreach ($row in $csv) {
		$s = Get-SPSite $row.SourceURL
		if ($s) {
			$storage = [Math]::Round($s.Usage.Storage / 1MB, 2)
            if (!($row.PSObject.Properties.name -contains "SPStorageMB")) {
				# add property SPStorageMB to collection, if missing
                $row | Add-Member –MemberType NoteProperty –Name SPStorageMB –Value ""
            }
			$row.SPStorageMB = $storage
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
			$folder = "C:\Users\[USER]\AppData\Local\Sharegate\userdata".Replace("[USER]", $user)
			$folder
			Remove-Item $folder -Confirm:$false -Recurse -Force -ErrorAction SilentlyContinue
			$folder = "C:\Users\[USER]\AppData\Local\Sharegate\Sharegate.Migration.txt".Replace("[USER]", $user)
			$folder
			Remove-Item $folder -Confirm:$false -Force -ErrorAction SilentlyContinue
		}
	}
}

Function CheckInDocs ($url) {
	# SOM - Server Object Model OK
    $site = Get-SPSite $url
    
    foreach($web in $site.AllWebs) {
        
        foreach($list in $web.GetListsOfType([Microsoft.SharePoint.SPBaseType]::DocumentLibrary)) {
            $list.CheckedOutFiles | % { if ($_) {$_.TakeOverCheckOut() }}

            $list.CheckedOutFiles | % {
				$item = $list.GetItemById($_.ListItemId)
				if ($item.File) {
					$item.File.CheckIn("File checked in by administrator")
					Write-Host $item.File.ServerRelativeUrl -NoNewline; 
					Write-Host " Checked in " -ForegroundColor Green
				}
            }
        }
        $web.dispose();
    }
    $site.dispose();
}

Function SummaryFooter() {
	# Summary LOG footer
	Write-Host "===== DONE ===== $(Get-Date)" -Fore Yellow
	$th				= [Math]::Round(((Get-Date) - $global:start).TotalHours, 2)
	$attemptMb		= ($global:track | measure SPStorageMB -Sum).Sum
	$actualMb		= ($global:track |? {$_.SGSessionId -ne ""} | measure SPStorageMB -Sum).Sum
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
}

Function NewLog() {
	# Start LOG
	$global:start = Get-Date
	$when = $global:start.ToString("yyyy-MM-dd-hh-mm-ss")
	$logFile = "$root\log\InsaneMove-$when.txt"
	mkdir "$root\log" -ErrorAction SilentlyContinue | Out-Null
	if (!$psISE) {
		try {
			Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
		} catch {}
		Start-Transcript $logFile
	}
	Write-Host "fileCSV = $fileCSV"
}

function OpenRemoteFarm() {
    $global:remoteFarm = New-PSSession -ComputerName $settings.settings.sourceFarmServer -Credential $global:farmcred -Authentication Credssp
}

Function Main() {	
	# Migrate with -WhatIf paramter
	if ($whatif) {
		$migrate = $true
	}

	# Delete source SPSites
	if ($deleteSource) {
		DeleteSourceSites
		Exit
	}

	# Delete destination SPSites
	if ($deleteDest) {
		DeleteDestinationSites
		Exit
	}
	
	# Clean
	if ($clean) {
		Clean
		Exit
	}

	# Core logic
	if ($userProfile) {
		# Update local user profiles.  Set Personal site URL for Hybrid OneDrive audience compilation and redirect
		NewLog
		ReadCloudPW
		ConnectCloud
		UserProfileSetHybridURL
		CompileAudiences
	} elseif ($measure) {
		# Populate CSV with size (GB)
		NewLog
		MeasureSiteCSV
	} elseif ($readOnly) {
		# Lock on-prem sites
		NewLog
		LockSite "ReadOnly"
	} elseif ($readWrite) {
		# Unlock on-prem sites
		NewLog
		LockSite "Unlock"
	} elseif ($noAccess) {
		# NoAccess on-prem sites
		NewLog
		LockSite "NoAccess"	
	} elseif ($verifyCloudSites) {
		# Create site collection
		NewLog
		ReadCloudPW
		ConnectCloud
		VerifyCloudSites
	} elseif ($qualityAssurance) {
		# Run QA check
		NewLog
		ReadCloudPW
		ConnectCloud
		QualityAssurance | Out-Null
	} else {
		if ($migrate) {
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
				NewLog
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
				SaveMigrationCSV
				SummaryFooter
				
				# Delay and QA
				Write-Host "Wait 60 seconds ..." -Fore Green
				Sleep 60
				$file = QualityAssurance
				EmailQACSV $file
			}
		} else {
			# HelpMessage
			Write-Host "InsaneMove - USAGE" -Fore Yellow
			Write-Host "==================" -Fore Yellow
			Write-Host @"
-verifyCloudSites (-v)		Verify all Office 365 site collections.  Prep step before real migration.
-migrate (-mig)			Copy sites to Office 365.  Main method.
-fileCSV (-csv)			CSV input list of source and destination SharePoint URLs to copy to Office 365.	

-deleteSource (-ds)		Delete source SharePoint sites on-premise.
-deleteDest (-dd)		Delete destination SharePoint sites in cloud.

-incremental (-i)		Copy incremental changes only.  ShareGate Copy-Settings.
-measure (-m)			Measure size of site collections in GB.
-readOnly (-ro)			Lock sites read-only.
-readWrite (-rw)		Unlock sites read-write.
-noAccess (-na)			Lock sites no access.	
-userProfile (-ups)		Update local User Profile Service with cloud personal URL.
-dryRun (-d)			Replaces core "Copy-Site" with "NoCopy-Site" to run without data copy.
-clean (-c)				Clean servers to prepare for next migration batch.
-whatif (-whatif)		Pre-Migration Report for Copy-Site command with results of migration issues.

-qualityAssurance (-qa)		Compare source and destination lists for QA check.	
`n`n
"@
		}
	}
	# Stop transcript if running (suppress error)
	if (!$psISE) {
		try {
			Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
		} catch {}
	}
}
Main
