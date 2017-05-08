<#
.SYNOPSIS
	Insane Move - QA Report to verify files copied successfully.

.DESCRIPTION
	Count files in the "Documents" library on both source and destination sites.  Leverages PNP and SP Server cmdlets.
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory=$false, ValueFromPipeline=$false, HelpMessage='CSV list of source and destination SharePoint site URLs to copy to Office 365.')]
	[string]$fileCSV
)

$datestamp = (Get-Date).tostring("yyyy-MM-dd-hh-mm-ss")
Start-Transcript "migratedsites-$datestamp.csv"

# Plugins
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
Import-Module SharePointPnPPowerShellOnline -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null

Function ReadCSV($fileCSV) {
    # Read CSV data file
    $csv = Import-Csv $fileCSV

    # DataTable
    $global:dt = New-Object System.Data.DataTable
    $cols = @("Source","Destination","PersonalSpace","SourceCount","DestinationCount","MissingDocs")
    foreach ($col in $cols) {
        # Add column
        $dt.Columns.Add($col) | Out-Null
    }

    # Create UPM Object
    $mySiteWebApp = Get-SPWebApplication | ? {$_.Name -like "*MySite*"}
    $mySite = Get-SPSite $mySiteWebApp.Url
    $serviceContext = Get-SPServiceContext $mySite
    $upm = new-object Microsoft.Office.Server.UserProfiles.UserProfileManager($serviceContext)

    # Append rows
    foreach ($row in $csv) {
        # Lookup MySite
        $destUrl = FindCloudMySite $row.MySiteEmail

        # Lookup PersonalSpace attribute
        $personalSpace = FindPersonalSpace $row.MySiteEmail
				
        # Add Row
        $new = $global:dt.NewRow()
        $new["Source"] = $row.SourceURL
        $new["Destination"] = $destUrl
        $new["PersonalSpace"] = $personalSpace
        $global:dt.Rows.Add($new)
    }
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

Function FindPersonalSpace ($MySiteEmail) {
    # Lookup PersonalSpace attribute from MySite profile
    $uProfile = $upm.GetUserProfile("$MySiteEmail")
    $personalSpace = $uProfile["personalSpace"].value
    return $personalSpace
}

Function InspectSource($url) {
    # Connect MySite
    $site = Get-SPSite $url
    $web = $site.Rootweb
    $list = $web.Lists["Documents"]
    $c = $list.Items.Count + $list.Folders.Count

	Write-Host "`n--------------------------------------------------------------------------------------"
	Write-Host "--------------------------------------------------------------------------------------"
	Write-Host "--------------------------------------------------------------------------------------`n"
	
	Write-Host "`n MySite: $url `n"
	
    # Collect file URL detail to Data Table
    $global:mysite = New-Object System.Data.DataTable
    $global:mysite.Columns.Add("File") | Out-Null
    foreach ($item in $list.items) {
        $new = $global:mysite.NewRow()
	    $new.File = $item.File
	    $global:mysite.Rows.Add($new) | Out-Null
        Write-Host $new.File -ForegroundColor Green
    }

    # Return count
    $site.Dispose() | Out-Null
    return $c
}

Function InspectDestination($url) {
    # Connect OneDrive
    Connect-PnPOnline -Url $url | Out-Null
    $items = Get-PNPListItem -List "Documents" -Fields "FileRef" | %{New-Object PSObject -Property  @{Name = $_["FileRef"]}}
    $c = 0

	Write-Host "`n OneDrive: $url `n"
		
    # Parse URL for username
    $splits = $url.Split('/', [StringSplitOptions]::RemoveEmptyEntries)
    $user = $splits[$splits.length -1 ].Split('_')[0]

    # Collect file URL detail to Data Table
    $global:onedrive = New-Object System.Data.DataTable
    $global:onedrive.Columns.Add("File") | Out-Null
	
    foreach ($item in $items) {
        $c++
        $new = $global:onedrive.NewRow()
        $new.File = $item.Name.replace('_company_com','').Replace("/personal/$user/","")
        $global:onedrive.Rows.Add($new) | Out-Null
        Write-Host $new.File -ForegroundColor Magenta
    }

    # Return count
    return $c
}

Function CountMissing() {
    # Compare MySite and OneDrive file URLs
    $dv = New-Object System.Data.DataView($global:onedrive)
    $c = 0    
	Write-Host "********** MISSING DOCS **********"
	
    # Filter
     foreach ($row in $global:mysite) {
		# Escape apostrophe within file names
        $file = $row.File.Replace("'","''")
        $dv.RowFilter = "File = '$file'"
		
        if ($dv.Count -eq 0) {
            Write-Host $file -ForegroundColor Yellow
            $c++
        }
    }

    # Return count
    return $c
}

Function SaveReport() {
    # Save CSV report
    #$datestamp = (Get-Date).tostring("yyyy-MM-dd-hh-mm-ss")
	$file = "QAReport-migratedsites-$datestamp.csv"
    $global:dt | Export-Csv $file -NoTypeInformation
	Write-Host "Saved : $file"
}

# Main
Function Main() {
    # Read list of migrated sites CSV
	Connect-PnPOnline "https://tenant-admin.sharepoint.com"
    $dt = ReadCSV $fileCSV
	$dt |ft
	
    foreach ($row in $global:dt.Rows) {
	
        # Count source files (MySite On-Premise)
        $c = InspectSource $row.Source
        $row["SourceCount"] = $c

        # Count destination files (SharePoint Online)
        $c = InspectDestination $row.Destination
        $row["DestinationCount"] = $c

        # Count missing (compare MySite/OneDrive)
        $c = CountMissing
        $row["MissingDocs"] = $c
    }

    # Migration report to CSV
    SaveReport
}
Main

Stop-Transcript