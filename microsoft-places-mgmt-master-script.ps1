#Requires -Version 7.0
<#
.SYNOPSIS
Manage Microsoft Places resources interactively or through reusable functions.
.PARAMETER DefaultDomain
Domain appended to bare aliases in room onboarding, room unlinking, and desk pool
attribute updates. Omit this parameter when those CSVs use full email addresses.
.PARAMETER NoMenu
Load functions without starting the menu or a session transcript. Dot-source the
script with this switch to use its functions in your own automation.
#>
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [ValidatePattern('^[^\s@/\\]+(?:\.[^\s@/\\]+)+$')]
    [string]$DefaultDomain,

    [switch]$NoMenu
)

function Resolve-PlacesResourceAddress {
    param(
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Address
    )

    if ([string]::IsNullOrWhiteSpace($Address)) {
        throw "Resource address is missing. Supply a full email address or an alias with -DefaultDomain."
    }

    $Address = $Address.Trim()
    if ($Address.Contains('@')) {
        return $Address
    }

    if ([string]::IsNullOrWhiteSpace($script:DefaultDomain)) {
        throw "Alias '$Address' requires -DefaultDomain. Alternatively, supply a full email address."
    }

    return "$Address@$($script:DefaultDomain)"
}

if (-not $NoMenu) {
Write-Host "
█▀▄▀█ █ █▀▀ █▀█ █▀█ █▀ █▀█ █▀▀ ▀█▀   █▀█ █░░ ▄▀█ █▀▀ █▀▀ █▀   █▀▄▀█ ▄▀█ █▀ ▀█▀ █▀▀ █▀█   █▀ █▀▀ █▀█ █ █▀█ ▀█▀
█░▀░█ █ █▄▄ █▀▄ █▄█ ▄█ █▄█ █▀░ ░█░   █▀▀ █▄▄ █▀█ █▄▄ ██▄ ▄█   █░▀░█ █▀█ ▄█ ░█░ ██▄ █▀▄   ▄█ █▄▄ █▀▄ █ █▀▀ ░█░
" -ForegroundColor Gray
Write-host "Updated on: 2025-Aug-21" -ForegroundColor Gray `n
}

#Menu function
function Show-Menu {
    Write-Host "↓ Select the operation you want to perform: ↓ `n " -ForegroundColor Cyan
    Write-Host "0 - Connect to Microsoft Places [Prerequisite]" -ForegroundColor Green 
    Write-Host "1 - Search for Existing Building, Floor or Section by Name (wildcards supported)"
    Write-Host "2 - Onboard Rooms from CSV file"
    Write-Host "3 - Change Desk Pool Attribute"
    Write-Host "4 - Create New Building or Floor"
    Write-Host "5 - Setup auto room release for a room"
    Write-Host "6 - BSSID Update for a Building"
    Write-Host "7 - Break Fixes for Places [Updated]"
    write-host "8 - Create Desk (Bookable or non-bookable) OR DeskPool (Workspace)"
    Write-Host "9 - Upload Building MAP"
    Write-Host "10 - Remove Building, Floor or Rooms"
    write-host "11 - Remove Complete Building with Building ID [Unlink all Rooms, Floors, Sections]"
    Write-Host "12 - Set Desk Mode (Bookable, Assigned, Walkup) for Desks"
    Write-Host "Help - For Help"
    Write-host "X - to Exit" -ForegroundColor Yellow `n
    Write-Host "(Press Enter for default option [0])" -ForegroundColor DarkGray  
}

# func to connect-microsoftplaces, includes "IdleTimeout change to 18000000 (5 hours)"
function connectPlaces {
    Write-Host "Connecting to Microsoft Places..." -ForegroundColor Green
    try {
        # Connect to Microsoft Places
        New-PSSessionOption -IdleTimeout 18000000 | Out-Null
        Connect-MicrosoftPlaces -ErrorAction Stop
        Write-Output "Successfully connected to Microsoft Places"
    } catch {
        Write-Error "Failed to connect to Microsoft Places. Check the module installation and permissions. Error: $($_.Exception.Message)"
        Write-Host "Hint: Do 'Get-InstalledModule | Where-Object {$_.Name -like '*Places*'}', if no module found, run 'Install-Module -Name MicrosoftPlaces -Force' with PowerShell 7" 
        exit 1
    }
}

# Func to onboard Rooms to Places
function OnboardRoom {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )
    
    if ($FileName -eq "sample") {
        GenerateSampleCSV
        return
    }

    # Import the CSV file
    if (-not (Test-Path $FileName)) {
        Write-Error "File not found: $FileName"
        return
    } elseif ($FileName -notlike "*.csv") {
        Write-Error "The specified file is not a CSV file: $FileName"
        return
    }

    $rooms = @(Import-Csv $FileName)
    foreach ($room in $rooms) {
        $room.PrimarySmtpAddress = Resolve-PlacesResourceAddress -Address $room.PrimarySmtpAddress
    }

    # Format the output file names
    $cron_stamp = Get-Date -Format "MMdyyHHmmss"
    $file_name = "update_placesv3_output_$cron_stamp.csv"
    $UserObject_file_name = "Failed_Onboarding_$cron_stamp.csv"
    $transcript_file_name = "update_places_transcript_$cron_stamp.txt"
    
    # Array to track successful updates
    $successful_rooms = @()

    try {
        # Start transcript for logging
        Start-Transcript -Path $transcript_file_name

        # Variables
        $total_users_to_process = $rooms.Count
        $total_users_processed = 0
        Write-Host "Total rooms to process: $total_users_to_process" -ForegroundColor Cyan

        # Iterate through each room in the list
        foreach ($room in $rooms) {
            $Error.Clear()
            $total_users_processed++

            Write-Host "Processing room: $($room.PrimarySmtpAddress) ($total_users_processed of $total_users_to_process)" -ForegroundColor Cyan

            $primaryaddress = $room.PrimarySmtpAddress

            # Create a PSObject to track room processing
            $UserObject = [PSCustomObject]@{
                REQ_Type           = $room.REQ_Type
                PrimarySmtpAddress = $primaryaddress
                Building           = $room.Building
                Floor              = $room.Floor
                FloorLabel         = $room.FloorLabel
                Error              = $null
                Comments           = $null
                Status             = "Processing"
            }

            Write-Host "Processing $($UserObject.PrimarySmtpAddress) ($total_users_processed of $total_users_to_process)"

            # Find the building
            $building = Get-PlaceV3 -Type Building | Where-Object { $_.DisplayName -eq $room.Building }
            if ($building.PlaceId) {
                $buildingId = $building.PlaceId
            } else {
                $UserObject.Comments = "No Building Found. Verify the Building is correct for this room."
                $UserObject.Status = "Failed"
                $UserObject | Export-Csv -Path $UserObject_file_name -Append -NoTypeInformation
                continue
            }

            # Find or create the floor
            if (-not $room.FloorLabel) {
                $UserObject.Comments = "FloorLabel is missing for room $($room.PrimarySmtpAddress)."
                $UserObject.Status = "Failed"
                $UserObject | Export-Csv -Path $UserObject_file_name -Append -NoTypeInformation
                Write-Warning "FloorLabel is missing for room $($room.PrimarySmtpAddress). Skipping..."
                continue
            }

            $floor = (Get-PlaceV3 -Type Floor -AncestorId $buildingId) | Where-Object { $_.DisplayName -eq $room.FloorLabel }
            if ($floor.PlaceId) {
                $floorId = $floor.PlaceId
            } else {
                try {
                    $floor = New-Place -Type Floor -ParentID $buildingId -Name $room.FloorLabel
                    $floorId = $floor.Identity
                    Write-Host "Created new floor: $($room.FloorLabel) with ID: $floorId" -ForegroundColor Green
                } catch {
                    $UserObject.Comments = "Failed to create floor: $($_.Exception.Message)"
                    $UserObject.Error = $_.Exception.Message
                    $UserObject.Status = "Failed"
                    $UserObject | Export-Csv -Path $UserObject_file_name -Append -NoTypeInformation
                    continue
                }
            }
            Write-Host "Floor ID: $floorId"

            # Set floor properties - ensure Floor value is treated as integer
            $floorNumber = ($room.Floor -as [int])
            if ($null -eq $floorNumber) { $floorNumber = 1 }
            
            try {
                Set-PlaceV3 -Identity $floorId -SortOrder $floorNumber -DisplayName $room.FloorLabel
            } catch {
                Write-Warning "Failed to set floor properties: $($_.Exception.Message)"
            }

            # Handle MTR settings
            if ([string]::IsNullOrEmpty($room.MTREnabled)) {
                $Current_Setting = Get-PlaceV3 -Identity $UserObject.PrimarySmtpAddress -ErrorAction SilentlyContinue
                $mtrEnabled = if ($Current_Setting) { [System.Convert]::ToBoolean($Current_Setting.MTREnabled) } else { $false }
            } else {
                $mtrEnabled = [System.Convert]::ToBoolean($room.MTREnabled)
            }

            # Handle capacity - ensure it's treated as integer
            $capacity = ($room.Capacity -as [int])
            if ($null -eq $capacity) { $capacity = 0 }

            # Set room properties
            try {
                $setPlaceParams = @{
                    Identity = $primaryaddress
                    Capacity = $capacity
                    Label = $room.Label
                    ParentID = $floorId
                    Floor = $floorNumber
                    Building = $room.Building
                    FloorLabel = $room.FloorLabel
                    MTREnabled = $mtrEnabled
                }
                
                # Only add parameters that have values
                if ($room.AudioDeviceName) { $setPlaceParams.AudioDeviceName = $room.AudioDeviceName }
                if ($room.DisplayDeviceName) { $setPlaceParams.DisplayDeviceName = $room.DisplayDeviceName }
                if ($room.VideoDeviceName) { $setPlaceParams.VideoDeviceName = $room.VideoDeviceName }
                
                Set-PlaceV3 @setPlaceParams

                Write-Host "Successfully updated room: $primaryaddress" -ForegroundColor Green
                
                # Add to successful rooms list
                $UserObject.Status = "Success"
                $successful_rooms += $UserObject
            } catch {
                $UserObject.Error = $_.Exception.Message
                $UserObject.Comments = "Failed to update room properties"
                $UserObject.Status = "Failed"
                $UserObject | Export-Csv -Path $UserObject_file_name -Append -NoTypeInformation
            }
        }

        # Export successful rooms to CSV
        if ($successful_rooms.Count -gt 0) {
            $successful_rooms | Export-Csv -Path $file_name -NoTypeInformation
            Write-Host "Successfully processed $($successful_rooms.Count) rooms. Output saved in: $file_name" -ForegroundColor Green
        }
        
        Write-Host "Finished processing all $total_users_processed rooms." -ForegroundColor Green
    } finally {
        Stop-Transcript
    }
}

# func to generate a samplecsv for Onboarding Rooms
function GenerateSampleCSV {
    param (
        [string]$OutputPath = ".\Sample(template)_OnboardRoom.csv"
    )

    # Define sample data as PSCustomObject
    $sampleData = @(
        [PSCustomObject]@{
            REQ_Type            = "Focus Room"
            PrimarySmtpAddress  = "room101@example.com"
            Alias               = "room101"
            Label               = "Focus Room 101 (4)"
            Building            = "Building 1"
            FloorLabel          = "1"
            Capacity            = 4
            AudioDeviceName     = ""
            DisplayDeviceName   = ""
            VideoDeviceName     = ""
            MTREnabled          = $false
            Floor               = 1
            StreetAddress       = ""
            City                = ""
            StateOrProvince     = ""
            PostalCode          = ""
            CountryOrRegion     = ""
            RoomList            = "Building 1 Rooms"
            WheelChairAccessible = ""
        }
    )

    try {
        # Export sample data to CSV
        $sampleData | Export-Csv -Path $OutputPath -NoTypeInformation -Force
        Write-Host "Sample/Template CSV file created at: $OutputPath" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to create sample CSV file. Error: $($_.Exception.Message)"
    }
}

# Func to find a building or floor
function FindBuildingOrFloor {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SearchString,
        
        [Parameter(Mandatory = $false)]
        [switch]$ExportResults
    )

    Write-Host "Searching for Places objects matching: '$SearchString'" -ForegroundColor Cyan
    Write-Host "(Wildcards like * are supported for broader searches)" -ForegroundColor DarkGray

    # Initialize counters for summary
    $buildingCount = 0
    $floorCount = 0
    $sectionCount = 0
    $roomCount = 0
    $deskCount = 0
    
    # Create timestamp for exports
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportPath = "PlacesSearch_$timestamp.csv"
    $resultsToExport = @()  # Renamed from $exportResults to avoid parameter conflict

    # Check if the string matches a Building
    Write-Host "`nSearching for buildings..." -ForegroundColor Yellow
    $buildings = Get-PlaceV3 -Type Building -ResultSize unlimited | Where-Object { $_.DisplayName -like $SearchString }
    
    if ($buildings) {
        $buildingCount = if ($buildings -is [array]) { $buildings.Count } else { 1 }
        Write-Host "Found $buildingCount building(s)" -ForegroundColor Green
        
        # Process each building
        $buildingIndex = 0
        foreach ($building in @($buildings)) {
            $buildingIndex++
            
            # Visual separator for each building
            Write-Host ("`n" + "=" * 80) -ForegroundColor Blue
            Write-Host "BUILDING $($buildingIndex) of $($buildingCount): $($building.DisplayName)" -ForegroundColor White -BackgroundColor DarkBlue
            Write-Host ("=" * 80) -ForegroundColor Blue
            
            # Display building details in a structured format
            Write-Host "Building Details:" -ForegroundColor Cyan
            $buildingProps = [ordered]@{
                DisplayName = $building.DisplayName
                Type = $building.Type
                PlaceId = $building.PlaceId
                City = $building.City
                CountryOrRegion = $building.CountryOrRegion
                PostalCode = $building.PostalCode
                State = $building.State
                Street = $building.Street
            }
            [PSCustomObject]$buildingProps | Format-List
            
            # Add to export collection
            if ($ExportResults) {
                $resultsToExport += [PSCustomObject]@{
                    ObjectType = "Building"
                    Type = $building.Type
                    DisplayName = $building.DisplayName
                    PlaceId = $building.PlaceId
                    ParentId = $null
                    Level = 1
                    City = $building.City
                    CountryOrRegion = $building.CountryOrRegion
                    State = $building.State
                    Street = $building.Street
                }
            }
            
            # Get linked floors for this building
            Write-Host "Retrieving floors..." -ForegroundColor DarkGray
            $linkedFloors = Get-PlaceV3 -Type Floor -AncestorId $building.PlaceId -ResultSize unlimited | 
                            Sort-Object DisplayName
            
            $buildingFloorCount = if ($linkedFloors) { if ($linkedFloors -is [array]) { $linkedFloors.Count } else { 1 } } else { 0 }
            $floorCount += $buildingFloorCount
            
            if ($linkedFloors) {
                Write-Host "`nFLOORS ($buildingFloorCount total):" -ForegroundColor Green
                Write-Host ("-" * 40) -ForegroundColor DarkGreen
                # ADDED TYPE TO FLOOR TABLE OUTPUT
                $floorTable = $linkedFloors | Select-Object Type, DisplayName, PlaceId, Floor, Label | 
                               Sort-Object -Property Floor
                $floorTable | Format-Table -AutoSize
                
                # Add floors to export
                if ($ExportResults) {
                    foreach ($floor in $linkedFloors) {
                        $resultsToExport += [PSCustomObject]@{
                            ObjectType = "Floor"
                            Type = $floor.Type
                            DisplayName = $floor.DisplayName
                            PlaceId = $floor.PlaceId
                            ParentId = $building.PlaceId
                            Level = 2
                            Floor = $floor.Floor
                            Label = $floor.Label
                        }
                    }
                }
            } else {
                Write-Host "`nNo floors found in this building." -ForegroundColor Yellow
            }
            
            # Get sections for this building
            Write-Host "Retrieving sections..." -ForegroundColor DarkGray
            $buildingSections = @()
            
            # Get sections from each floor in the building
            foreach ($floor in @($linkedFloors)) {
                $floorSections = Get-PlaceV3 -Type Section -AncestorId $floor.PlaceId -ResultSize unlimited
                if ($floorSections) {
                    foreach ($section in @($floorSections)) {
                        # Only add sections with Type = Section
                        if ($section.Type -eq "Section") {
                            $buildingSections += [PSCustomObject]@{
                                DisplayName = $section.DisplayName
                                Type = $section.Type
                                PlaceId = $section.PlaceId
                                ParentId = $section.ParentPlaceId
                                FloorName = $floor.DisplayName
                                FloorId = $floor.PlaceId
                            }
                        }
                    }
                }
            }
            
            $buildingSectionCount = $buildingSections.Count
            $sectionCount += $buildingSectionCount
            
            if ($buildingSectionCount -gt 0) {
                Write-Host "`nSECTIONS ($buildingSectionCount total):" -ForegroundColor Magenta
                Write-Host ("-" * 40) -ForegroundColor DarkMagenta
                $buildingSections | Format-Table Type, DisplayName, PlaceId, FloorName -AutoSize
                
                # Add sections to export
                if ($ExportResults) {
                    foreach ($section in $buildingSections) {
                        $resultsToExport += [PSCustomObject]@{
                            ObjectType = "Section"
                            Type = $section.Type
                            DisplayName = $section.DisplayName
                            PlaceId = $section.PlaceId
                            ParentId = $section.ParentId
                            FloorName = $section.FloorName
                            Level = 3
                        }
                    }
                }
            } else {
                Write-Host "`nNo sections found in this building." -ForegroundColor Yellow
            }
            
            # Get rooms for this building
            Write-Host "Retrieving rooms..." -ForegroundColor DarkGray
            $buildingRooms = @()
            
            # Get rooms from each section in the building
            foreach ($section in $buildingSections) {
                $sectionRooms = Get-PlaceV3 -Type Room -AncestorId $section.PlaceId -ResultSize unlimited
                if ($sectionRooms) {
                    foreach ($room in @($sectionRooms)) {
                        $buildingRooms += [PSCustomObject]@{
                            DisplayName = $room.DisplayName
                            Type = $room.Type
                            PlaceId = $room.PlaceId
                            ParentId = $room.ParentPlaceId
                            SectionName = $section.DisplayName
                            FloorName = $section.FloorName
                            Capacity = $room.Capacity
                        }
                    }
                }
            }
            
            # Get rooms directly under floors (no section)
            foreach ($floor in @($linkedFloors)) {
                $floorRooms = Get-PlaceV3 -Type Room -AncestorId $floor.PlaceId -ResultSize unlimited
                if ($floorRooms) {
                    foreach ($room in @($floorRooms)) {
                        $buildingRooms += [PSCustomObject]@{
                            DisplayName = $room.DisplayName
                            Type = $room.Type
                            PlaceId = $room.PlaceId
                            ParentId = $room.ParentPlaceId
                            SectionName = "<No Section>"
                            FloorName = $floor.DisplayName
                            Capacity = $room.Capacity
                        }
                    }
                }
            }
            
            # Get desks for this building
            $buildingDesks = @()
            foreach ($section in $buildingSections) {
                $sectionDesks = Get-PlaceV3 -Type Desk -AncestorId $section.PlaceId -ResultSize unlimited
                if ($sectionDesks) {
                    foreach ($desk in @($sectionDesks)) {
                        $buildingDesks += [PSCustomObject]@{
                            DisplayName = $desk.DisplayName
                            Type = "Desk"
                            PlaceId = $desk.PlaceId
                            ParentId = $desk.ParentPlaceId
                            SectionName = $section.DisplayName
                            FloorName = $section.FloorName
                            Capacity = $desk.Capacity
                        }
                    }
                }
            }
            
            # Get desks directly under floors (no section)
            foreach ($floor in @($linkedFloors)) {
                $floorDesks = Get-PlaceV3 -Type Desk -AncestorId $floor.PlaceId -ResultSize unlimited
                if ($floorDesks) {
                    foreach ($desk in @($floorDesks)) {
                        $buildingDesks += [PSCustomObject]@{
                            DisplayName = $desk.DisplayName
                            Type = "Desk"
                            PlaceId = $desk.PlaceId
                            ParentId = $desk.ParentPlaceId
                            SectionName = "<No Section>"
                            FloorName = $floor.DisplayName
                            Capacity = $desk.Capacity
                        }
                    }
                }
            }
            
            $buildingRoomCount = $buildingRooms.Count
            $buildingDeskCount = $buildingDesks.Count
            $roomCount += $buildingRoomCount
            $deskCount += $buildingDeskCount
            
            # Combine for display
            $combinedItems = $buildingRooms + $buildingDesks
            $combinedCount = $combinedItems.Count
            
            if ($combinedCount -gt 0) {
                Write-Host "`nROOMS & DESKS ($combinedCount total - $buildingRoomCount rooms, $buildingDeskCount desks):" -ForegroundColor Yellow
                Write-Host ("-" * 40) -ForegroundColor DarkYellow
                
                # Only show first 20 rooms to avoid flooding the console
                if ($combinedCount > 20) {
                    $combinedItems | Select-Object Type, DisplayName, PlaceId, SectionName, FloorName, Capacity | 
                    Sort-Object -Property Type, FloorName, SectionName | 
                    Select-Object -First 20 | Format-Table -AutoSize
                    
                    Write-Host "... and $($combinedCount - 20) more rooms/desks (use export option to see all)" -ForegroundColor DarkYellow
                } else {
                    $combinedItems | Select-Object Type, DisplayName, PlaceId, SectionName, FloorName, Capacity |
                    Sort-Object -Property Type, FloorName, SectionName |
                    Format-Table -AutoSize
                }
                
                # Add rooms to export
                if ($ExportResults) {
                    foreach ($item in $combinedItems) {
                        $resultsToExport += [PSCustomObject]@{
                            ObjectType = $item.Type
                            Type = $item.Type
                            DisplayName = $item.DisplayName
                            PlaceId = $item.PlaceId
                            ParentId = $item.ParentId
                            SectionName = $item.SectionName
                            FloorName = $item.FloorName
                            Capacity = $item.Capacity
                            Level = 4
                        }
                    }
                }
            } else {
                Write-Host "`nNo rooms or desks found in this building." -ForegroundColor Yellow
            }
            
            # Building summary
            Write-Host "`nBUILDING SUMMARY:" -ForegroundColor Cyan
            Write-Host "Floors:   $buildingFloorCount" -ForegroundColor Green
            Write-Host "Sections: $buildingSectionCount" -ForegroundColor Magenta
            Write-Host "Rooms:    $buildingRoomCount" -ForegroundColor Yellow
            Write-Host "Desks:    $buildingDeskCount" -ForegroundColor Yellow
        }
    } else {
        Write-Warning "No buildings found matching '$SearchString'."
    }
    
    # Check for direct floor matches
    Write-Host "`nSearching for floors directly..." -ForegroundColor Yellow
    $floors = Get-PlaceV3 -Type Floor -ResultSize unlimited | Where-Object { $_.DisplayName -like $SearchString }
    
    $directFloorCount = 0
    if ($floors) {
        $directFloorCount = if ($floors -is [array]) { $floors.Count } else { 1 }
        Write-Host "Found $directFloorCount floor(s) directly matching the search" -ForegroundColor Green
        Write-Host ("-" * 80) -ForegroundColor DarkCyan
        $floors | Format-Table Type, DisplayName, PlaceId, ParentId, Floor, Label -AutoSize
    } else {
        Write-Warning "No floors found directly matching '$SearchString'."
    }
    
    # Check for direct section matches
    Write-Host "`nSearching for sections directly..." -ForegroundColor Yellow
    $sections = Get-PlaceV3 -Type Section -ResultSize unlimited | Where-Object { 
        $_.DisplayName -like $SearchString -and $_.Type -eq "Section" 
    }
    
    $directSectionCount = 0
    if ($sections) {
        $directSectionCount = if ($sections -is [array]) { $sections.Count } else { 1 }
        Write-Host "Found $directSectionCount section(s) directly matching the search" -ForegroundColor Green
        Write-Host ("-" * 80) -ForegroundColor DarkCyan
        $sections | Format-Table Type, DisplayName, PlaceId, ParentId -AutoSize
    } else {
        Write-Warning "No sections found directly matching '$SearchString'."
    }
    
    # Calculate safely
    $floorCountInBuildings = $floorCount - $directFloorCount
    $sectionCountInBuildings = $sectionCount - $directSectionCount
    
    # Global search summary
    Write-Host "`nOVERALL SEARCH SUMMARY:" -ForegroundColor Cyan
    Write-Host "Buildings: $buildingCount" -ForegroundColor Blue
    Write-Host "Floors:    $floorCount (including $floorCountInBuildings in buildings)" -ForegroundColor Green
    Write-Host "Sections:  $sectionCount (including $sectionCountInBuildings in buildings)" -ForegroundColor Magenta
    Write-Host "Rooms:     $roomCount" -ForegroundColor Yellow
    Write-Host "Desks:     $deskCount" -ForegroundColor Yellow
    
    # Export results if requested
    if ($ExportResults -and $resultsToExport.Count -gt 0) {
        $resultsToExport | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "`nExported search results to: $exportPath" -ForegroundColor Green
    }
    
    # Return results as a hashtable for further use
    return @{
        Buildings = $buildings
        Floors = $floors
        Sections = $sections
        Statistics = @{
            BuildingCount = $buildingCount
            FloorCount = $floorCount
            SectionCount = $sectionCount
            RoomCount = $roomCount
            DeskCount = $deskCount
        }
    }
}

#func for Breakfix issues in Places
function Breakfix {
    Write-Host "Review the intended calendar-processing change before applying a fix. See Readme.md for guidance." -ForegroundColor Yellow
    do {
        Write-host "0 - Connect to Exchange Online"
        Write-host "1 - Capacity not showing for Wokspaces/Desk/Room use [Connect-ExchangeOnline Required]"
        Write-host "2 - Capacity not showing for Wokspaces/Desk/Room, [BULK Update] update using CSV [Connect-ExchangeOnline Required]" -ForegroundColor Cyan
        Write-host "3 - Wokspaces/Desk/Room 'Unavailable' [Connect-ExchangeOnline Required]"
        Write-host "4 - Wokspaces/Desk/Room 'Unavailable', [BULK Update] update using CSV [Connect-ExchangeOnline Required]" -ForegroundColor Cyan
        Write-host "? - Check the CSV formatting for both the option 2 & 4" -ForegroundColor Green
        Write-host "5 - Rooms/Workspace/Desk capacity update"
        Write-host "9 - Exit" -ForegroundColor Yellow `n 
    
        $select = Read-Host "Select the issue you are facing:" 

        switch ($select) {
            "0" { Write-Host "Connecting to Exchange Online..." -ForegroundColor Green
            try {
                # Attempt to connect to Exchange Online
                Connect-ExchangeOnline -ErrorAction Stop -ShowBanner:$false | Out-Null
                Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
            } catch {
                # Handle connection failure
                Write-Error "Cannot connect to Exchange Online. Check the module installation or your permissions. Error: $($_.Exception.Message)"
            }
        }
        "1" {
            # Option 1: Process a single room
            Write-Host "Make sure you are running this against resource mailbox (and not user mailbox)." -ForegroundColor Yellow
            $room = Read-Host "Enter the alias/UPN of the Room/Workspace"
            try {
                Write-Host "Processing room: $room" -ForegroundColor Green
                Set-CalendarProcessing -Identity $room -EnforceCapacity $True
                Write-Host "Successfully updated room: $room" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to process room: $room. Error: $($_.Exception.Message)"
            }
        }
    
        "2" {
            # Option 2: Bulk update rooms using a CSV file
            Write-Host "Make sure you are running this against resource mailbox (and not user mailbox)." -ForegroundColor Yellow
            $csvPath = Read-Host "Enter the path to the CSV file containing the list of rooms"
            if (Test-Path $csvPath) {
                $rooms = Import-Csv $csvPath
                foreach ($room in $rooms) {
                    try {
                        Write-Host "Processing room: $($room.Alias)" -ForegroundColor Green
                        Set-CalendarProcessing -Identity $room.Alias -EnforceCapacity $True
                        Write-Host "Successfully updated room: $($room.Alias)" -ForegroundColor Green
                    } catch {
                        Write-Warning "Failed to process room: $($room.Alias). Error: $($_.Exception.Message)"
                    }
                }
            } else {
                Write-Warning "The specified CSV file does not exist. Please check the path and try again."
            }
        }
            "3" {
            # Option 1: Process a single room
            Write-Host "Make sure you are running this against resource mailbox (and not user mailbox)." -ForegroundColor Yellow
            $room = Read-Host "Enter the alias/UPN of the Room/Workspace"
            try {
                Write-Host "Processing room: $room" -ForegroundColor Green
                Set-CalendarProcessing -Identity $room -AllBookInPolicy $True
                Write-Host "Successfully updated room: $room" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to process room: $room. Error: $($_.Exception.Message)"
            }
        }
    
        "4" {
            # Option 2: Bulk update rooms using a CSV file
            Write-Host "Make sure you are running this against resource mailbox (and not user mailbox)." -ForegroundColor Yellow
            $csvPath = Read-Host "Enter the path to the CSV file containing the list of rooms"
            if (Test-Path $csvPath) {
                $rooms = Import-Csv $csvPath
                foreach ($room in $rooms) {
                    try {
                        Write-Host "Processing room: $($room.Alias)" -ForegroundColor Green
                        Set-CalendarProcessing -Identity $room.Alias -AllBookInPolicy $True
                        Write-Host "Successfully updated room: $($room.Alias)" -ForegroundColor Green
                    } catch {
                        Write-Warning "Failed to process room: $($room.Alias). Error: $($_.Exception.Message)"
                    }
                }
            } else {
                Write-Warning "The specified CSV file does not exist. Please check the path and try again."
            }
        }
            "5" {$alias = Read-Host "Provide alias/upn or the room you want to update the capacity for:"
                 [int] $capacity = Read-Host "Provide the capacity you want to set for the room:"
                 Set-PlaceV3 -Identity $alias -Capacity $capacity}

            "?" {Write-host "The CSV file should just have a column named Alias containing the room aliases/UPNs. For example:
            Alias
            room1@example.com
            room2@example.com
            room3@example.com
            " -ForegroundColor Green}

            "9" {Write-Host "Exiting script." -ForegroundColor Red; (Mainblock)}
            
            Default { Write-Warning "Invalid choice." }
        }
    
    } while ($true)

} 

# fun to Remove Room
function Remove-Room {
    param(
        [Parameter(Mandatory=$true)]
        [string]$csvFilePath
    )

    # Get the current folder 
    $currentFolder = (Get-Location).Path
    $fileName = "Remove_Room_Output.csv"
    $filePath = Join-Path $currentFolder $fileName
    $results = @()

    try {
        # Import the data from the CSV file
        $csvData = Import-Csv -Path $csvFilePath
        
        # Validate CSV has the required column
        if (-not ($csvData | Get-Member -Name PrimarySmtpAddress)) {
            Write-Error "CSV file must contain a 'PrimarySmtpAddress' column"
            return
        }

        foreach ($room in $csvData) {
            $room.PrimarySmtpAddress = Resolve-PlacesResourceAddress -Address $room.PrimarySmtpAddress
        }
        
        # Iterate through each room
        $count = 1
        $total = $csvData.Count
        
        foreach ($room in $csvData) {
            Write-Host "Processing $count/$total" -ForegroundColor Cyan
            
            # Initialize the result object
            $result = [PSCustomObject]@{
                Identity = $room.PrimarySmtpAddress
                Error = $null
                Comment = "Processing"
            }
            
            $roomAddress = $room.PrimarySmtpAddress
            
            # Attempt to remove the room
            try {
                Set-PlaceV3 -Identity $roomAddress -ParentId $null
                $result.Comment = "SUCCEED"
            } catch {
                $result.Error = $_.Exception.Message
                $result.Comment = "FAILED"
            }
            
            # Add to results array
            $results += $result
            $count++
        }
        
        # Export all results at once
        $results | Export-Csv -Path $filePath -NoTypeInformation
        Write-Host "Operation complete. Results saved to: $filePath" -ForegroundColor Green
        
    } catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
    }
}

# fun to remove floors
function Remove-Floors {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$CsvFilePath
    )

    # Create timestamp for logs and output files
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFileName = "Floor_Removal_Log_$timestamp.csv"
    $currentFolder = (Get-Location).Path
    $logFilePath = Join-Path $currentFolder $logFileName
    
    # Create array for logging results
    $results = @()
    
    try {
        # Import CSV and validate required columns
        $csvData = Import-Csv -Path $CsvFilePath
        
        if (-not ($csvData | Get-Member -MemberType NoteProperty -Name "ParentName") -or
            -not ($csvData | Get-Member -MemberType NoteProperty -Name "FloorLabel")) {
            Write-Error "CSV file must contain 'ParentName' and 'FloorLabel' columns"
            return
        }
        
        # Display overview and confirmation prompt
        Write-Host "Found $($csvData.Count) floors to remove." -ForegroundColor Cyan
        $csvData | Format-Table -AutoSize
        
        if (-not $WhatIfPreference) {
            Write-Host "WARNING: This will permanently remove floors from buildings." -ForegroundColor Red
            $confirmation = Read-Host "Type 'Yes' to continue or anything else to abort"
            if ($confirmation -ne "Yes") {
                Write-Host "Operation cancelled." -ForegroundColor Yellow
                return
            }
        }
        
        # Process each row
        $count = 1
        $total = $csvData.Count
        
        foreach ($row in $csvData) {
            Write-Host "Processing $count/$total - Building: $($row.ParentName), Floor: $($row.FloorLabel)" -ForegroundColor Cyan
            
            # Initialize result object for this row
            $result = [PSCustomObject]@{
                ParentName = $row.ParentName
                FloorLabel = $row.FloorLabel
                BuildingPlaceId = $null
                FloorPlaceId = $null
                Error = $null
                Status = "Processing"
                Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
            
            # Find the building
            try {
                $building = Get-PlaceV3 -Type Building -ResultSize unlimited | 
                            Where-Object { $_.DisplayName -like $row.ParentName }
                
                if (-not $building) {
                    $result.Error = "Building not found"
                    $result.Status = "Failed"
                    $results += $result
                    Write-Warning "Building not found: $($row.ParentName)"
                    $count++
                    continue
                }
                
                if ($building.Count -gt 1) {
                    $result.Error = "Multiple buildings found with this name"
                    $result.Status = "Failed"
                    $results += $result
                    Write-Warning "Multiple buildings found with name: $($row.ParentName)"
                    $count++
                    continue
                }
                
                $result.BuildingPlaceId = $building.PlaceId
                
                # Find the floor within the building
                $floor = Get-PlaceV3 -Type Floor -AncestorId $building.PlaceId | 
                         Where-Object { $_.DisplayName -like $row.FloorLabel }
                
                if (-not $floor) {
                    $result.Error = "Floor not found in building"
                    $result.Status = "Failed"
                    $results += $result
                    Write-Warning "Floor not found: $($row.FloorLabel) in building: $($row.ParentName)"
                    $count++
                    continue
                }
                
                if ($floor.Count -gt 1) {
                    $result.Error = "Multiple floors found with this label"
                    $result.Status = "Failed"
                    $results += $result
                    Write-Warning "Multiple floors found with label: $($row.FloorLabel) in building: $($row.ParentName)"
                    $count++
                    continue
                }
                
                $result.FloorPlaceId = $floor.PlaceId
                
                if ($PSCmdlet.ShouldProcess("Floor '$($floor.DisplayName)' in building '$($building.DisplayName)' ($($floor.PlaceId))", "Remove floor")) {
                    Remove-Place -Identity $floor.PlaceId -ErrorAction Stop
                    $result.Status = "Success"
                    Write-Host "Successfully removed floor: $($row.FloorLabel) from building: $($row.ParentName)" -ForegroundColor Green
                } else {
                    $result.Status = if ($WhatIfPreference) { "WhatIf" } else { "Skipped" }
                }
                
            } catch {
                $result.Error = $_.Exception.Message
                $result.Status = "Failed"
                Write-Warning "Error processing $($row.ParentName) - $($row.FloorLabel): $($_.Exception.Message)"
            }
            
            # Add to results collection
            $results += $result
            if (-not $WhatIfPreference) {
                $result | Export-Csv -Path $logFilePath -NoTypeInformation -Append -Confirm:$false
            }
            $count++
        }
        
        if ($WhatIfPreference) {
            Write-Host "Preview complete. No floors were removed and no log file was written." -ForegroundColor Cyan
        } else {
            $results | Export-Csv -Path $logFilePath -NoTypeInformation -Confirm:$false
            Write-Host "Operation complete. Log saved to: $logFilePath" -ForegroundColor Green
        }
        
        # Display summary
        $successful = ($results | Where-Object { $_.Status -eq "Success" }).Count
        $failed = ($results | Where-Object { $_.Status -eq "Failed" }).Count
        $previewed = ($results | Where-Object { $_.Status -eq "WhatIf" }).Count
        $skipped = ($results | Where-Object { $_.Status -eq "Skipped" }).Count
        
        Write-Host "Summary: Successfully removed $successful floors, Failed: $failed, Previewed: $previewed, Skipped: $skipped" -ForegroundColor Cyan
        
    } catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
    }
}

# func to remove buildings
function Remove-Building {
    # Define the function parameters
    param(
        [string]$csvFilePath
    )

    $Error.Clear()

    # Import the data from the CSV file
    $csvData = Import-Csv -Path $csvFilePath

    
    # Get the current folder 
    $currentFolder = (Get-Location).Path

    $fileName = "Remove_bldg_Output.csv"
    $filePath = Join-Path $currentFolder $fileName

    #start the count
    $count = 1

    # Iterate over each row of data from the CSV file
    foreach ($bldg in $csvData) {
        
        Write-Host('{0}/{1}' -f $count, $csvData.count)
        
        # Initialize a hashtable of parameters to pass to the Set-PlaceV3 command
        $params = New-Object -TypeName PSObject -Property ([ordered] @{
                BuildingId = $bldg.BuildingId
                BuildingName = $bldg.BuildingName
                Error=$null
                Comment='SUCCEED'
            })
        
        $buildingId = (Get-PlaceV3 -type Building -Name ("{0}" -f $params.BuildingName))[0].Identity
        Remove-Place $buildingId
        
        $error_results = switch($Error.count){
            0{"No Errors"}
            default{"Errors"}
        }

        if($error_results -eq "Errors"){
            $params.Error = $Error.Exception
            $params.Comment = 'FAILED'
        }else{
            $params.Comment = 'SUCCEED'
        }
        # Export the error log to the specified file on the user's desktop
        $params | Export-Csv -Path $filePath -NoTypeInformation -Append

        $count++
        $Error.Clear()
    }
}

#Func to remove entire building and all its child objects
function RemoveEntireBuilding {
    param(
        [Parameter(Mandatory=$true)]
        [string]$BuildingPlaceId
    )

    # Define the correct removal order explicitly
    $removalOrder = @('Room', 'Desk', 'Section', 'Floor', 'Building')

    # Get all child objects recursively
    $childPlaces = Get-PlaceV3 -AncestorId $BuildingPlaceId -ResultSize unlimited
     if (-not $childPlaces) { 
        Write-Warning "No child places or Building found for Building PlaceId: $BuildingPlaceId"
        return 0 
    }

    # Get the building itself to include in the export
    $building = Get-PlaceV3 -Identity $BuildingPlaceId
    
    # Create timestamp for filename
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportPath = "Building_Hierarchy_Export_$timestamp.csv"
    
    # Combine the building and all child places for the export
    $allPlaces = @($building) + $childPlaces
    
    # Export all places data to CSV before removal
    $allPlaces | Select-Object DisplayName, Alias, PlaceId, Type, Label, Identity, ParentId, Capacity, CountryOrRegion |
    Export-Csv -Path $exportPath -NoTypeInformation
        
    Write-Host "Exported building hierarchy to: $exportPath" -ForegroundColor Green
    Write-Host "Total exported objects: $($allPlaces.Count)" -ForegroundColor Cyan
    
    # Continue with the rest of removal function...
    # Sort child objects explicitly based on defined removal order
    $sortedChildPlaces = $childPlaces | Sort-Object @{Expression = { $removalOrder.IndexOf($_.Type) }}

    # Display the removal order (for verification)
    Write-Host "Objects will be processed & REMOVED in this order:" -ForegroundColor Yellow
    $sortedChildPlaces | Select-Object Type, DisplayName | Format-Table

    # Add confirmation prompt
    Write-Host "WARNING: This will remove or unlink all objects shown above." -ForegroundColor Red
    Write-Host "Type 'Yes' to confirm removal of these objects." -ForegroundColor Red
    $confirmation = Read-Host "Proceed with removal?"

    # Check if confirmation is explicitly "Yes"
    if ($confirmation -ne "Yes") {
        Write-Host "Operation cancelled. No changes were made." -ForegroundColor Cyan
        return
    }

    # Use the SORTED list in the foreach loop with different handling based on type
    foreach ($place in $sortedChildPlaces) {
        try {
            # Different handling based on type
            if ($place.Type -eq "Room" -or $place.Type -eq "Desk") {
                # For Rooms, Desks, and Sections - just unlink them by setting ParentId to null
                Write-Host "Unlinking $($place.Type): $($place.DisplayName)" -ForegroundColor Cyan
                Set-PlaceV3 -Identity $place.PlaceId -ParentId $null
                Write-Host "Unlinked $($place.Type): $($place.DisplayName)" -ForegroundColor Green
            } else {
                # For Floors, Section, Buildings - use Remove-Place
                Write-Host "Removing $($place.Type): $($place.DisplayName)" -ForegroundColor Cyan
                Remove-Place -Identity $place.PlaceId
                Write-Host "Removed $($place.Type): $($place.DisplayName)" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to process $($place.Type) '$($place.DisplayName)': $($_.Exception.Message)"
        } Write-host "Building $($building.DisplayName) removed successfully" -ForegroundColor Green
    }

    # Finally, remove the building itself explicitly
    try {
        Write-Host "Removing Building (trying again if above failed): $BuildingPlaceId" -ForegroundColor Cyan
        Remove-Place -Identity $BuildingPlaceId
        Write-Host "Successfully removed Building with PlaceId: $BuildingPlaceId" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to remove Building: $($_.Exception.Message)"
    }
}

# func to give the user options to remove building, floors and rooms
function RemoveBuildingOptionMenu {

    Write-Host "
    IMPORTANT: Removal of Places hierarchy must follow this order (you can skip the ones you don't want to remove):
       1. Remove Rooms/Workspaces
       2. Remove Floors/Sections
       3. Remove Building
          
       Removal Hierarchy:
            --Room/Workspace
              --Floor/Section
                --Building" -ForegroundColor Yellow `n

    do {
        Write-host "1- Remove Room using csv"
        Write-host "2- Deskpool or Workspace removal [Unavailable]" -ForegroundColor DarkGray
        Write-host "3- Remove Floor using csv"
        Write-host "4- Remove Building using csv"
        Write-host "9- to Exit out of script" -ForegroundColor Yellow `n

        $select = Read-Host "Select the option you want to perform:"
        switch ($select) {
            "1" {Remove-Room -csvFilePath (Read-Host "Enter the CSV file path")}
            "2" {Write-Warning "Deskpool or Workspace removal is not implemented. No changes were made."}
            "3" {Remove-Floors -csvFilePath (Read-Host "Enter the CSV file path")}
            "4" {Remove-Building -csvFilePath (Read-Host "Enter the CSV file path")}
            "9" {Write-Host "Exiting script." -ForegroundColor Red; (Mainblock)}
            Default { Write-Warning "Invalid choice." }
        }

    } while ($true)
    
}

# Func to update BSSID (updated on 18 Aug 25)
function BSSID {
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputFile
    )

    # Helper function: Determines whether an update is required comparing key fields.
    function NeedsUpdate($existingDevice, $newDevice) {
        if ($existingDevice.DeviceId -ne $newDevice.BSSID) { return $true }
        if ($existingDevice.MACAddress -ne $newDevice.BSSID) { return $true }

        return $false
    }

    function AddDevice($bssid, $placeId, [ref]$reportAdded, [ref]$reportUpdated, [ref]$reportFailed) {    
        # Write-Host "Adding new device with BSSID $bssid"
        $result = New-PlaceDevice `
            -DeviceId $bssid `
            -Manufacturer "MSPlaces" `
            -PlaceId $placeId `
            -MACAddress $bssid `
            -Sensors (New-Object MicrosoftPlaces.PlacesDevices.Sensor -Property @{SensorType="Wifi"}) `
            -ErrorVariable addError

        if ($result) {
            # Write-Host "    - Add succeeded"
            $reportAdded.Value += [PSCustomObject]@{
                BSSID   = $bssid                    
                PlaceId = $placeId
                Action  = "Added"
            }
        } else {
            if ($addError -Match "Duplicate record exists") {
                Write-Host "Add failed, will update"
                $existingDevice = Get-PlaceDevice -Filter "MacAddress eq '$bssid'" -Top 1
                UpdateDevice $existingDevice[0].Id $newDevice.BSSID $newDevice.PlaceId $reportUpdated $reportFailed
            } else {
                # Write-Host "    - Add failed"
                $reportFailed.Value += [PSCustomObject]@{
                    PlaceId = $placeId
                    BSSID   = $bssid
                    Action  = "Add Failed"
                    Reason  = "" + $addError
                }
            }
        }
    }

    function UpdateDevice($id, $bssid, $placeId, [ref]$reportUpdated, [ref]$reportFailed) {
        # Write-Host "Updating BSSID $bssid"
        $result = Set-PlaceDevice `
            -Id $id `
            -Manufacturer "MSPlaces" `
            -DeviceId $bssid `
            -PlaceId $placeId `
            -MACAddress $bssid `
            -Sensors (New-Object MicrosoftPlaces.PlacesDevices.Sensor -Property @{SensorType="Wifi"})

        if ($result) {
            # Write-Host "    - Update succeeded"
            $reportUpdated.Value += [PSCustomObject]@{
                BSSID   = $bssid
                PlaceId = $placeId                        
                Action  = "Updated"
            }
        } else {
            # Write-Host "    - Update failed"
            $reportFailed.Value += [PSCustomObject]@{
                BSSID   = $bssid                        
                PlaceId = $placeId
                Action  = "Update Failed"
                Reason  = $_.Exception.Message
            }
        }
    }

    # Import the CSV data from the provided input file.
    $inputDevices = Import-Csv -Path $InputFile

    # Load already processed buildings.
    if (Test-Path "processed_buildings.csv") {
        $processedBuildings = Import-Csv -Path "processed_buildings.csv"
        Write-Output "Previously processed buildings count is $($processedBuildings.Count)"
    } else {
        $processedBuildings = @()
        Write-Output "There's no previously proccessed buildings."
    }

    # Group the devices by PlaceId so that we query inventory only once per PlaceId.
    $groupedDevices = $inputDevices | Group-Object PlaceId
    $groupedDevices = $groupedDevices | Sort-Object -Property Count
    Write-Host "There are $($groupedDevices.Count) buildings in the file"
    $i = 0
    $consecutiveFailures = 0

    foreach ($group in $groupedDevices) {
        $i++
        $placeId = $group.Name    

        $wasProcessed = $processedBuildings | Where-Object { $_.PlaceId -eq $placeId }
        if ($wasProcessed) {
            Write-Host "${i}: Skipping already processed PlaceId: $placeId"
            continue
        }

        Write-Host "${i}: Processing PlaceId: $placeId"

        # Initialize report arrays for this PlaceId.
        $reportAdded   = @()
        $reportUpdated = @()
        $reportIgnored = @()
        $reportFailed  = @()

        # Query existing devices for this PlaceId.
        try {
            $existingDevices = Get-PlaceDevice -Filter "PlaceId eq '$placeId'" -Top 5000
        }
        catch {
            Write-Host "Error querying devices for PlaceId $($placeId): $_"
            continue
        }

        # Process each input device record.
        foreach ($newDevice in $group.Group) {
            $bssid = $newDevice.BSSID
            # Write-Host "Processing BSSID $bssid"
            $existing = $existingDevices | Where-Object { $_.MACAddress -eq $bssid }
            if ($existing) {
                # Device exists; check if an update is needed.
                if (NeedsUpdate $existing $newDevice) {
                    $beforeFailures = $reportFailed.Count
                    UpdateDevice $existing.Id $newDevice.BSSID $newDevice.PlaceId ([ref]$reportUpdated) ([ref]$reportFailed)
                    if ($reportFailed.Count -gt $beforeFailures) {
                        $consecutiveFailures++
                    } else {
                        $consecutiveFailures = 0
                    }
                } else {
                    # Devices that require no change are simply skipped.
                    # Write-Host "    - Update ignored"
                    $reportIgnored += [PSCustomObject]@{
                        BSSID   = $bssid                    
                        PlaceId = $placeId
                        Action  = "Ignored"
                    }
                }
            } else {
                $beforeFailures = $reportFailed.Count
                # Device does not exist; attempt to add it.
                AddDevice $newDevice.BSSID $newDevice.PlaceId ([ref]$reportAdded) ([ref]$reportUpdated) ([ref]$reportFailed)
                if ($reportFailed.Count -gt $beforeFailures) {
                    $consecutiveFailures++
                } else {
                    $consecutiveFailures = 0
                }
            }

            # Stop if failure count is too high
            if ($consecutiveFailures -ge 10) {
                Write-Host "Too many consecutive failures ($consecutiveFailures). Stopping script."
                break
            }
        }

        $outputDir = "output"
        if (-not (Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory | Out-Null
            Write-Host "Created 'output' directory for reports."
        }

        $addedFile   = Join-Path $outputDir "$placeId-Added.csv"
        $updatedFile = Join-Path $outputDir "$placeId-Updated.csv"
        $ignoredFile = Join-Path $outputDir "$placeId-Ignored.csv"
        $failedFile  = Join-Path $outputDir "$placeId-Failed.csv"

        # Export each report array to its respective CSV file if it contains any entries.
        Write-Host "Added $($reportAdded.Count) entries."
        # if ($reportAdded.Count -gt 0) {
            $reportAdded | Export-Csv -Path $addedFile -NoTypeInformation
            # Write-Host "Added report generated: $addedFile"
        # }

        Write-Host "Updated $($reportUpdated.Count) entries."
        # if ($reportUpdated.Count -gt 0) {
            $reportUpdated | Export-Csv -Path $updatedFile -NoTypeInformation
            # Write-Host "Updated report generated: $updatedFile"
        # }

        Write-Host "Ignored $($reportIgnored.Count) entries."
        # if ($reportIgnored.Count -gt 0) {
            $reportIgnored | Export-Csv -Path $ignoredFile -NoTypeInformation
            # Write-Host "Updated report generated: $ignoredFile"
        # }

        Write-Host "Failed $($reportFailed.Count) entries."
        # if ($reportFailed.Count -gt 0) {
            $reportFailed | Export-Csv -Path $failedFile -NoTypeInformation
            # Write-Host "Failed report generated: $failedFile"
        # }

        if ($consecutiveFailures -lt 3) {
            [PSCustomObject]@{ PlaceId = $group.Name} | Export-Csv -Path "processed_buildings.csv" -Append -NoTypeInformation
        }

        if ($consecutiveFailures -ge 5) {
            return
        }    
    }
}

# Updated DeskPoolAttri function with default Space filtering
function DeskPoolAttri {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputFile, # Import CSV
        
        [Parameter(Mandatory = $false)]
        [switch]$ProcessAllTypes # Set to true to process ALL mailbox types (default is Space only)
    )
    
    Write-Host "By default this function will only process mailboxes/Resrouces with 'type' -Space, you'll get output file for processed resources" -ForegroundColor Yellow
    Write-Host "Use -ProcessAllTypes to process all mailbox types, this usally don't work, please check out this article: https://learn.microsoft.com/en-us/microsoft-365/places/configure-desk-booking#configure-desk-pools" -ForegroundColor Gray

    # Validate the input file
    if (-not (Test-Path $InputFile)) {
        Write-Error "The specified input file does not exist: $InputFile"
        return
    } elseif ($InputFile -notlike "*.csv") {
        Write-Error "The specified file is not a CSV file: $InputFile"
        return
    }

    # Create timestamp for output filenames
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFolder = (Get-Location).Path
    $resultsFile = Join-Path $outputFolder "DeskPool_Results_$timestamp.csv"

    # Import the CSV data
    $Data = Import-Csv $InputFile
    foreach ($room in $Data) {
        $room.Alias = Resolve-PlacesResourceAddress -Address $room.Alias
    }
    
    # Track rooms by type for filtering
    $validRooms = @{}
    $processResults = @() # Track all processing results in one collection
    $script:SectionCache = @{} # Cache for sections

    Write-Host "Checking resource mailbox types..." -ForegroundColor Cyan
    
    # Check and validate all rooms first
    foreach ($room in $Data) {
        try {
            $alias = $room.Alias
            
            Write-Host "Checking $alias type..." -ForegroundColor Gray
            $detail = Get-PlaceV3 -Identity $alias -ErrorAction Stop | 
                      Select-Object DisplayName, Type, PlaceID, ParentID
            
            if ($detail.Type -eq "Space") {
                Write-Host "✓ $alias is of type 'Space'" -ForegroundColor Green
                $validRooms[$alias] = $true
            } else {
                Write-Host "✗ $alias is type '$($detail.Type)' (not Space)" -ForegroundColor Yellow
                $processResults += [PSCustomObject]@{
                    Alias = $alias
                    Type = $detail.Type
                    BuildingName = $room.'Building Name'
                    FloorName = $room.'Floor Name'
                    DeskPoolName = $room.DeskPoolName
                    Capacity = $room.Capacity
                    Status = "Skipped"
                    Reason = "Not of type 'Space'"
                    ProcessedOn = Get-Date
                }
                $validRooms[$alias] = $false
            }
        } catch {
            Write-Warning "Error checking room '$alias': $($_.Exception.Message)"
            $processResults += [PSCustomObject]@{
                Alias = $alias
                Type = "Unknown"
                BuildingName = $room.'Building Name'
                FloorName = $room.'Floor Name'
                DeskPoolName = $room.DeskPoolName
                Capacity = $room.Capacity
                Status = "Skipped"
                Reason = $_.Exception.Message
                ProcessedOn = Get-Date
            }
            $validRooms[$alias] = $false
        }
    }
    
    # Summary of room types
    $spaceCount = ($validRooms.GetEnumerator() | Where-Object { $_.Value -eq $true }).Count
    $totalRooms = $Data.Count
    
    Write-Host "`nFound $spaceCount Space-type rooms out of $totalRooms total rooms." -ForegroundColor Cyan
    
    if ($ProcessAllTypes) {
        Write-Host "Will process ALL rooms regardless of type (including non-Space rooms)." -ForegroundColor Yellow
    } else {
        Write-Host "Will only process rooms of type 'Space' (default behavior)." -ForegroundColor Yellow
        Write-Host "Use -ProcessAllTypes to process all room types." -ForegroundColor Gray
    }
    
    Write-Host "Processing rooms..." -ForegroundColor Cyan
    
    # Process each room based on type filter
    foreach ($room in $Data) {
        $alias = $room.Alias
        
        # Skip if room is not Space type (default behavior)
        if (-not $ProcessAllTypes -and -not $validRooms[$alias]) {
            # Already logged in skipped rooms above
            Write-Host "Skipping $alias (not Space type)" -ForegroundColor Yellow
            continue
        }
        
        try {
            $buildingName = $room.'Building Name'
            $floorName = $room.'Floor Name'
            
            Write-Host "Processing $alias..." -ForegroundColor Cyan

            # Get the building
            $building = Get-PlaceV3 -Type Building | Where-Object { $_.DisplayName -eq $buildingName }
            if (-not $building) {
                $processResults += [PSCustomObject]@{
                    Alias = $alias
                    Type = $validRooms[$alias] ? "Space" : "Other" 
                    BuildingName = $buildingName
                    FloorName = $floorName
                    DeskPoolName = $room.DeskPoolName
                    Capacity = $room.Capacity
                    Status = "Failed"
                    Reason = "Building not found: '$buildingName'"
                    ProcessedOn = Get-Date
                }
                Write-Warning "Building '$buildingName' not found. Skipping room '$alias'."
                continue
            }
            
            if ($building.Count -gt 1) {
                Write-Warning "Multiple buildings found with name '$buildingName'. Using the first one."
                $building = $building[0]
            }

            # Get the floor
            $floor = Get-PlaceV3 -Type Floor -AncestorId $building.PlaceId | 
                     Where-Object { $_.DisplayName -eq $floorName }
                     
            if (-not $floor) {
                $processResults += [PSCustomObject]@{
                    Alias = $alias
                    Type = $validRooms[$alias] ? "Space" : "Other"
                    BuildingName = $buildingName
                    FloorName = $floorName
                    DeskPoolName = $room.DeskPoolName
                    Capacity = $room.Capacity
                    Status = "Failed"
                    Reason = "Floor not found: '$floorName' in building: '$buildingName'"
                    ProcessedOn = Get-Date
                }
                Write-Warning "Floor '$floorName' not found in building '$buildingName'. Skipping room '$alias'."
                continue
            }
            
            if ($floor.Count -gt 1) {
                Write-Warning "Multiple floors found with name '$floorName'. Using the first one."
                $floor = $floor[0]
            }

            # Cache for sections to avoid repeated prompts
            $sectionCacheKey = "$($floor.PlaceId)_$floorName"
            
            # Check section cache first
            $section = $null
            if ($script:SectionCache.ContainsKey($sectionCacheKey)) {
                $section = $script:SectionCache[$sectionCacheKey]
                Write-Host "Using cached section information." -ForegroundColor Gray
            } else {
                # Check if the section exists
                $section = Get-PlaceV3 -Type Section -AncestorId $floor.PlaceId | 
                           Where-Object { $_.DisplayName -eq "section $floorName" }
                           
                if ($section) {
                    Write-Host "Section already exists. Updating desk pool attributes..." -ForegroundColor Green
                    $script:SectionCache[$sectionCacheKey] = $section
                } else {
                    # Prompt to create the section
                    $selection = Read-Host "Section does not exist. Do you want to create the section? (Y/N)"
                    if ($selection -eq "Y") {
                        $sectionName = "section $floorName"
                        $section = New-Place -Type Section -Name $sectionName -ParentId $floor.PlaceId
                        Write-Host "Section '$sectionName' created successfully." -ForegroundColor Green
                        $script:SectionCache[$sectionCacheKey] = $section
                    } else {
                        $processResults += [PSCustomObject]@{
                            Alias = $alias
                            Type = $validRooms[$alias] ? "Space" : "Other"
                            BuildingName = $buildingName
                            FloorName = $floorName
                            DeskPoolName = $room.DeskPoolName
                            Capacity = $room.Capacity
                            Status = "Skipped"
                            Reason = "Section creation declined by user"
                            ProcessedOn = Get-Date
                        }
                        Write-Host "Skipping room '$alias' as section creation was declined." -ForegroundColor Yellow
                        continue
                    }
                }
            }

            # Handle capacity as integer
            $capacity = [int]($room.Capacity -as [int])
            
            # Update desk pool attributes
            Set-PlaceV3 -Identity $alias -Label $room.DeskPoolName -Capacity $capacity -ParentId $section.PlaceId
            
            # Add to results collection
            $processResults += [PSCustomObject]@{
                Alias = $alias
                Type = $validRooms[$alias] ? "Space" : "Other"
                BuildingName = $buildingName
                FloorName = $floorName
                DeskPoolName = $room.DeskPoolName
                Capacity = $capacity
                Status = "Success"
                Reason = "Desk pool attributes updated"
                ProcessedOn = Get-Date
            }
            
            Write-Host "✓ Desk pool attributes updated for '$alias'." -ForegroundColor Green
        } catch {
            $processResults += [PSCustomObject]@{
                Alias = $alias
                Type = $validRooms[$alias] ? "Space" : "Other"
                BuildingName = $room.'Building Name'
                FloorName = $room.'Floor Name'
                DeskPoolName = $room.DeskPoolName
                Capacity = $room.Capacity
                Status = "Failed"
                Reason = $_.Exception.Message
                ProcessedOn = Get-Date
            }
            Write-Warning "Error processing room '$alias': $($_.Exception.Message)"
        }
    }
    
    # Output all results to a single CSV file
    $processResults | Export-Csv -Path $resultsFile -NoTypeInformation
    
    # Summary statistics
    $successCount = ($processResults | Where-Object { $_.Status -eq "Success" }).Count
    $failedCount = ($processResults | Where-Object { $_.Status -eq "Failed" }).Count
    $skippedCount = ($processResults | Where-Object { $_.Status -eq "Skipped" }).Count
    
    Write-Host "`nProcessing summary:" -ForegroundColor Cyan
    Write-Host "  Success: $successCount" -ForegroundColor Green
    Write-Host "  Failed:  $failedCount" -ForegroundColor Red
    Write-Host "  Skipped: $skippedCount" -ForegroundColor Yellow
    Write-Host "`nFull results saved to: $resultsFile" -ForegroundColor Cyan
    
    # Display results table
    $processResults | Format-Table Alias, Status, Reason -AutoSize
}

# fun to setup auto release for a room, this sets the time "PostReservationMaxClaimTimeInMinutes" to 15 min.
function AutoRelease {
    # Define Releasetime function OUTSIDE the loop to fix scope issues
    function Get-ReleaseTime {
        $timeout = 15 # Default value
        
        do {
            try {
                $input = Read-Host "Enter the time in minutes for AutoRelease (default is '15' minutes)"
                
                # Use empty input for default
                if ([string]::IsNullOrWhiteSpace($input)) {
                    Write-Host "Using default value of 15 minutes." -ForegroundColor Yellow
                    return 15
                }
                
                # Try to convert to integer safely
                $timeout = [int]::Parse($input)
                
                # Validate range
                if ($timeout -lt 1 -or $timeout -gt 20) {
                    Write-Host "Value must be between 1 and 20 minutes." -ForegroundColor Red
                } else {
                    Write-Host "AutoRelease timeout set to $timeout minutes." -ForegroundColor Green
                    return $timeout
                }
            }
            catch [System.FormatException] {
                Write-Host "Please enter a valid number." -ForegroundColor Red
            }
        } while ($true)
    }
    
    do {
        Write-host "0 - Connect to Exchange Online [Pre-requisite]" -ForegroundColor Yellow
        Write-host "1 - Set AutoRelease for a single room" -ForegroundColor Green
        Write-host "2 - Set AutoRelease for multiple rooms using CSV" -ForegroundColor Green
        Write-host "? - Check the CSV formatting/Template for option 2" -ForegroundColor Green
        Write-host "9 - Exit" -ForegroundColor Yellow `n

        $select = Read-Host "Select the option you want to perform:"
         
        switch ($select) {
            "0" {
                Write-Host "Connecting to Exchange Online..." -ForegroundColor Green
                try {
                    Connect-ExchangeOnline -ErrorAction Stop -ShowBanner:$false | Out-Null
                    Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
                } catch {
                    Write-Error "Failed to connect to Exchange Online. Check your permissions or module installation."
                }
            }
            "1" {
                # Set AutoRelease for a single room
                $auto = Read-Host "Enter the Alias or UPN of the room you want to set AutoRelease for"
                
                # Get the release timeout value
                $releaseTimeout = Get-ReleaseTime
                Write-Host "DEBUG: Release timeout value is: $releaseTimeout" -ForegroundColor Cyan
                
                try {
                    Set-CalendarProcessing -Identity $auto -EnableAutoRelease $true -PostReservationMaxClaimTimeInMinutes $releaseTimeout
                    Write-Host "AutoRelease set for $auto with PostReservationMaxClaimTimeInMinutes to $releaseTimeout minutes" -ForegroundColor Green
                } catch {
                    Write-Warning "Failed to set AutoRelease for $auto. Error: $($_.Exception.Message)"
                }
            }
            "2" {
                # Set AutoRelease for multiple rooms using CSV
                $csvPath = Read-Host "Enter the path to the CSV file containing the list of rooms"
                
                # Get the release timeout value
                $releaseTimeout = Get-ReleaseTime
                Write-Host "DEBUG: Release timeout value is: $releaseTimeout" -ForegroundColor Cyan
                
                if (Test-Path $csvPath) {
                    $rooms = Import-Csv $csvPath
                    foreach ($room in $rooms) {
                        try {
                            # Verify we have the Alias property
                            if (-not $room.Alias) {
                                Write-Warning "Room entry missing Alias property. Skipping."
                                continue
                            }
                            
                            Set-CalendarProcessing -Identity $room.Alias -EnableAutoRelease $true -PostReservationMaxClaimTimeInMinutes $releaseTimeout
                            Write-Host "AutoRelease set for $($room.Alias) with PostReservationMaxClaimTimeInMinutes to $releaseTimeout minutes" -ForegroundColor Green
                        } catch {
                            Write-Warning "Failed to set AutoRelease for $($room.Alias). Error: $($_.Exception.Message)"
                        }
                    }
                } else {
                    Write-Warning "The specified CSV file does not exist. Please check the path and try again."
                }
            }
            "?" {
                # Display CSV formatting instructions
                Write-Host "The CSV file should have a column named 'Alias' containing the room aliases/UPNs. For example:" -ForegroundColor Green
                Write-Host "
                    Alias
                    room1@example.com
                    room2@example.com
                    room3@example.com
                "
            }    
            "9" {
                Write-Host "Exiting Script." -ForegroundColor Yellow
                return  # Use return instead of calling Mainblock directly
            }
            Default {
                Write-Warning "Invalid choice. Please select a valid option."
            }
        }
    } while ($true)
}

# Function to create new building or floor from CSV
function NewbldgFloor {
    do {
        Write-host "##################################################### Create New Building or Floor ##############################################################################" -ForegroundColor Green
        Write-host "Places follow heirarcy, make sure you create the building first (if it doesn't exsist) and then create the floor" `n
        Write-host "1- Create New Building from CSV"
        Write-host "2- Create new Floor from CSV" 
        Write-host "9- Exit back to main menu" `n

        $option = Read-host "Select the option you want to perform:"
        
        # Check for exit option first before asking for CSV
        if ($option -eq "9") { 
            Write-Host "Returning to main menu..." -ForegroundColor Cyan
            return  # Exit the function cleanly
        }
        
        # Only ask for CSV if we're creating buildings or floors
        if ($option -eq "1" -or $option -eq "2") {
            $csv = Read-host "Enter the CSV file path"

            # Test if the CSV file exists and is valid
            do {
                if (-not (Test-Path $csv)) {
                    Write-Error "File not found: $csv"
                    $csv = Read-host "Enter the CSV file path"
                } elseif ($csv -notlike "*.csv") {
                    Write-Error "The specified file is not a CSV file: $csv"
                    $csv = Read-host "Enter the CSV file path"
                } else {
                    Write-host "Will be processing the file $csv" -ForegroundColor Green
                    break
                }
            } while ($true)
            
            switch ($option) {
                "1" {
                    # Create New Building from CSV
                    $buildingData = Import-Csv -Path $csv
                    foreach ($building in $buildingData) {
                        try {
                            # Capture the new building object returned from New-Place
                            $newBuilding = New-Place -Type Building -Name $building.BuildingName -CountryOrRegion $building.CountryOrRegion -City $building.City -State $building.State -Street $building.Street -PostalCode $building.PostalCode -ErrorAction Stop
                        
                            # Display success message
                            Write-Host "Building '$($building.BuildingName)' created successfully." -ForegroundColor Green `n
                            
                            Write-host "Exporting building details to CSV..." -ForegroundColor Green
                            # Create a custom object with the properties we want
                            $exportObject = [PSCustomObject]@{
                                DisplayName = $newBuilding.DisplayName
                                PlaceId = $newBuilding.PlaceId
                                City = $building.City
                                CountryOrRegion = $building.CountryOrRegion
                                BuildingName = $building.BuildingName
                            }
                            
                            # Export to current directory with building name in filename
                            $exportObject | Export-Csv -Path "NewBuilding_$($building.BuildingName).csv" -NoTypeInformation -Append
                        } catch {
                            Write-Warning "Failed to create building '$($building.BuildingName)': $_"
                        }
                    }
                }
                "2" {
                    # Create New Floor from CSV
                    $floorData = Import-Csv -Path $csv
                    foreach ($floor in $floorData) {
                        try {
                            # Get all buildings matching the display name
                            $matchingBuildings = Get-PlaceV3 -Type Building -ResultSize unlimited | Where-Object { $_.DisplayName -eq $floor.ParentName }
                            
                            # Handle different scenarios based on number of matches
                            if (-not $matchingBuildings) {
                                Write-Warning "No building found with name: $($floor.ParentName). Please check the CSV file."
                                continue
                            }
                            elseif ($matchingBuildings.Count -eq 1) {
                                $Parentid = $matchingBuildings.PlaceId
                                Write-Host "Building found: $($matchingBuildings.DisplayName) with ParentId: $Parentid" -ForegroundColor Green
                            }
                            else {
                                Write-Host "Multiple buildings found with name: $($floor.ParentName)" -ForegroundColor Yellow
                                
                                for ($i = 0; $i -lt $matchingBuildings.Count; $i++) {
                                    Write-Host "[$i] DisplayName: $($matchingBuildings[$i].DisplayName), PlaceId: $($matchingBuildings[$i].PlaceId), Location: $($matchingBuildings[$i].City), $($matchingBuildings[$i].CountryOrRegion)" -ForegroundColor Cyan
                                }
                                
                                $selection = Read-Host "Enter the index number of the correct building [0-$($matchingBuildings.Count - 1)]"
                                
                                if ($selection -match '^\d+$' -and [int]$selection -ge 0 -and [int]$selection -lt $matchingBuildings.Count) {
                                    $Parentid = $matchingBuildings[[int]$selection].PlaceId
                                    Write-Host "Selected building: $($matchingBuildings[[int]$selection].DisplayName) with ParentId: $Parentid" -ForegroundColor Green
                                }
                                else {
                                    Write-Warning "Invalid selection. Skipping floor creation for: $($floor.Name)"
                                    continue
                                }
                            }
                            
                            New-Place -Type Floor -Name $floor.FloorLabel -ParentId $Parentid
                            Write-Host "Floor '$($floor.FloorLabel)' created successfully under ParentId $Parentid." -ForegroundColor Green
                        
                        } catch {
                            Write-Warning "Failed to create floor '$($floor.FloorLabel)': $_"
                        }
                    }
                }
            }
        } else {
            Write-Warning "Invalid choice. Please select options 1, 2, or 9."
        }
    } while ($true)
}

# Function to create a Desk (bookable or non-bookable) and DeskPool (aka Workspace)
function BookableDesks {
    ## This script creates the Desk objects in Places, and make them Bookable desk (where mailbox is required)
    ## article: https://learn.microsoft.com/en-us/microsoft-365/places/configure-desk-booking
    Write-Host "######################################################## MAKE Sure you are Connect to Exchange Online ########################################################" -ForegroundColor Yellow
    $FileName = Read-Host "Enter the path to the CSV file containing bookable desk information"

    # Validate file exists
    if (-not (Test-Path $FileName)) {
        Write-Error "File not found: $FileName"
        return
    }

    $file = Import-Csv $FileName

    Write-Host "## This script creates the Desk objects in Places, and makes them Bookable desks (where mailbox is required)
    " -ForegroundColor Cyan

    # Summary counters
    $desksCreated = 0
    $sectionsCreated = 0
    $errors = 0
    $skipped = 0
    $mailboxesConverted = 0

    # Create array to store results for export
    $results = @()

    # STEP 1: PRE-PROCESS ALL MAILBOXES FIRST
    Write-Host "`n== PHASE 1: CONVERTING MAILBOXES ==" -ForegroundColor Cyan

    $mailboxesToProcess = $file | Where-Object { ![string]::IsNullOrEmpty($_.PrimarySmtpAddress) } | Select-Object -ExpandProperty PrimarySmtpAddress -Unique
    $totalMailboxes = $mailboxesToProcess.Count

    Write-Host "Found $totalMailboxes unique mailboxes to convert to Desk type" -ForegroundColor Yellow

    foreach ($mailbox in $mailboxesToProcess) {
        try {
            # Get current mailbox
            $mbx = Get-Mailbox $mailbox -ErrorAction Stop
            
            if ($mbx.Type -ne "Desk") {
                Write-Host "Converting mailbox '$mailbox' from type '$($mbx.Type)' to 'Desk'" -ForegroundColor Gray
                Set-Mailbox $mailbox -Type Desk -HiddenFromAddressListsEnabled $true
                $mailboxesConverted++
            } else {
                Write-Host "Mailbox '$mailbox' is already type 'Desk'" -ForegroundColor Green
                
                # Ensure it's hidden from address lists
                if (-not $mbx.HiddenFromAddressListsEnabled) {
                    Write-Host "Setting mailbox '$mailbox' to be hidden from address lists" -ForegroundColor Gray
                    Set-Mailbox $mailbox -HiddenFromAddressListsEnabled $true
                }
            }
        } catch {
            Write-Host "Error processing mailbox '$mailbox': $($_.Exception.Message)" -ForegroundColor Red
            $errors++
        }
    }

    # Give Exchange time to process the mailbox type changes
    if ($mailboxesConverted -gt 0) {
        $waitTime = 5
        Write-Host "`nConverted $mailboxesConverted mailboxes. Waiting $waitTime seconds for Exchange to process changes..." -ForegroundColor Yellow
        Start-Sleep -Seconds $waitTime
    }

    # STEP 2: PROCESS BUILDINGS, SECTIONS, AND DESKS

    Write-Host "`n== PHASE 2: CREATING DESK STRUCTURE ==" -ForegroundColor Cyan

    # New: caches to avoid repeated Get-PlaceV3 calls
    $buildingCache = @{}
    $floorCache    = @{}   # key: "<BuildingPlaceId>|<FloorLabel>"
    $sectionCache  = @{}   # key: "<FloorPlaceId>|<SectionName>"

    foreach ($row in $file) {
        try {
            Write-Host "Processing desk $($row.SpaceID)..." -ForegroundColor Yellow

            # ---------------- Building ----------------
            $buildingName = $row.Building
            if ($buildingCache.ContainsKey($buildingName)) {
                $building = $buildingCache[$buildingName]
                Write-Host "  Using cached Building '$buildingName'" -ForegroundColor DarkGray
            } else {
                Write-Host "  Checking if Building '$buildingName' exists" -ForegroundColor Gray
                $building = Get-PlaceV3 -Type Building | Where-Object DisplayName -eq $buildingName
                if (-not $building) {
                    Write-Host "  Building '$buildingName' does not exist, please check" -ForegroundColor Red
                    $errors++
                    continue
                }
                if ($building -is [array]) {
                    Write-Host "  Multiple buildings named '$buildingName', using first" -ForegroundColor Yellow
                    $building = $building[0]
                }
                $buildingCache[$buildingName] = $building
                Write-Host "  Cached Building '$buildingName'" -ForegroundColor Green
            }

            # ---------------- Floor ----------------
            $floorKey = "$($building.PlaceId)|$($row.FloorLabel)"
            if ($floorCache.ContainsKey($floorKey)) {
                $floor = $floorCache[$floorKey]
                Write-Host "  Using cached Floor '$($row.FloorLabel)'" -ForegroundColor DarkGray
            } else {
                Write-Host "  Checking if Floor '$($row.FloorLabel)' exists" -ForegroundColor Gray
                $floor = Get-PlaceV3 -Type Floor -AncestorId $building.PlaceId |
                        Where-Object { $_.Type -eq "Floor" -and $_.DisplayName -eq $row.FloorLabel }
                if (-not $floor) {
                    Write-Host "  Floor '$($row.FloorLabel)' does not exist, please check" -ForegroundColor Red
                    $errors++
                    continue
                }
                if ($floor -is [array]) {
                    Write-Host "  Multiple floors named '$($row.FloorLabel)', using first" -ForegroundColor Yellow
                    $floor = $floor[0]
                }
                $floorCache[$floorKey] = $floor
                Write-Host "  Cached Floor '$($row.FloorLabel)'" -ForegroundColor Green
            }

            # ---------------- Section ----------------
            $sectionName = $row.Section
            $sectionKey  = "$($floor.PlaceId)|$sectionName"
            if ($sectionCache.ContainsKey($sectionKey)) {
                $section = $sectionCache[$sectionKey]
                Write-Host "  Using cached Section '$sectionName'" -ForegroundColor DarkGray
            } else {
                Write-Host "  Checking if Section '$sectionName' exists" -ForegroundColor Gray
                $section = Get-PlaceV3 -Type Section -AncestorId $floor.PlaceId |
                        Where-Object DisplayName -eq $sectionName
                if (-not $section) {
                    $createSection = Read-Host "  Section '$sectionName' does not exist. Create it? (Y/N)"
                    if ($createSection -eq "Y") {
                        Write-Host "  Creating section '$sectionName'" -ForegroundColor Yellow
                        $section = New-Place -Type Section -Name $sectionName -ParentId $floor.PlaceId
                        if (-not $section) {
                            Write-Host "  Failed to create section '$sectionName'" -ForegroundColor Red
                            $errors++
                            continue
                        }
                        Write-Host "  Section created successfully" -ForegroundColor Green
                        $sectionsCreated++
                    } else {
                        Write-Host "  Skipping desk, section not created" -ForegroundColor Yellow
                        $skipped++
                        continue
                    }
                } else {
                    if ($section -is [array]) {
                        Write-Host "  Multiple sections named '$sectionName', using first" -ForegroundColor Yellow
                        $section = $section[0]
                    }
                    Write-Host "  Section exists, proceeding" -ForegroundColor Green
                }
                $sectionCache[$sectionKey] = $section
                Write-Host "  Cached Section '$sectionName'" -ForegroundColor Green
            }

            # ---------------- Desk Creation (unchanged except for using cached objects) ----------------
            $deskName = "Desk $($row.SpaceID)"
            Write-Host "  Creating desk '$deskName'..." -ForegroundColor Gray

            $existingDesk = Get-PlaceV3 -Type Desk -AncestorId $section.PlaceId |
                            Where-Object DisplayName -eq $deskName

            if ($existingDesk) {
                Write-Host "  Desk '$deskName' already exists, skipping creation" -ForegroundColor Yellow
                $skipped++
                $desk = $existingDesk
            } else {
                $desk = New-Place -Type Desk -Name $deskName -ParentId $section.PlaceId
                if (-not $desk) {
                    Write-Host "  Failed to create desk '$deskName'" -ForegroundColor Red
                    $errors++
                    continue
                } else {
                    Write-Host "  Desk '$deskName' created successfully" -ForegroundColor Green
                    $desksCreated++
                }
            }

            $results += [PSCustomObject]@{
                DisplayName  = $desk.DisplayName
                PlaceId      = $desk.PlaceId
                ParentId     = $desk.ParentPlaceId
                Type         = "Desk"
                Building     = $row.Building
                Floor        = $row.FloorLabel
                Section      = $sectionName
                SpaceID      = $row.SpaceID
                CreatedTime  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Mailbox      = $row.PrimarySmtpAddress
            }

            if ($row.PrimarySmtpAddress) {
                $linked = $false
                $maxRetries = 3
                $retryCount = 0
                $waitBetweenRetries = 5
                while (-not $linked -and $retryCount -le $maxRetries) {
                    try {
                        if ($retryCount -gt 0) {
                            Write-Host "  Retry attempt $retryCount of $maxRetries..." -ForegroundColor Yellow
                        }
                        Write-Host "  Linking desk '$deskName' to mailbox '$($row.PrimarySmtpAddress)'" -ForegroundColor Gray
                        Set-PlaceV3 $desk.PlaceId -Mailbox $row.PrimarySmtpAddress -ErrorAction Stop
                        Write-Host "  Mailbox linked successfully" -ForegroundColor Green
                        $linked = $true
                    } catch {
                        $retryCount++
                        if ($retryCount -le $maxRetries) {
                            Write-Host "  Error linking mailbox, retry in $waitBetweenRetries s: $($_.Exception.Message)" -ForegroundColor Yellow
                            Start-Sleep -Seconds $waitBetweenRetries
                        } else {
                            Write-Host "  Error linking mailbox after $maxRetries retries: $($_.Exception.Message)" -ForegroundColor Red
                            $errors++
                        }
                    }
                }
            } else {
                Write-Host "  No mailbox provided for desk '$deskName', skipping linking" -ForegroundColor Yellow
            }

            # Desk mode setting
            $modemetadata = New-Object 'System.Collections.Generic.Dictionary[String,object]'
            
            # Check if Mode property exists in the CSV
            if ($row.PSObject.Properties.Name -contains "Mode") {
                if ($row.Mode -eq "Assigned" -and $row.PSObject.Properties.Name -contains "AssignedTo") {
                    Write-Host "  Setting desk '$deskName' mode to 'Assigned' for user '$($row.AssignedTo)'" -ForegroundColor Gray
                    $modemetadata.Add("AssignedPersonEmailAddress",$($row.AssignedTo))
                    Set-PlaceV3 $desk.PlaceId -Mode @{Name="Assigned"; Metadata=$modemetadata}
                } 
                elseif ($row.Mode -eq "Bookable" -or $row.Mode -eq "Walkup") {
                    Write-Host "  Setting desk '$deskName' mode to '$($row.Mode)'" -ForegroundColor Gray
                    Set-PlaceV3 $desk.PlaceId -Mode @{Name="$($row.Mode)"; Metadata=$modemetadata}
                }
                else {
                    Write-Host "  Unrecognized mode '$($row.Mode)' for desk '$deskName', defaulting to Offline" -ForegroundColor Yellow
                    Set-PlaceV3 $desk.PlaceId -Mode @{Name="Offline"; Metadata=$modemetadata}
                }
            } 
            else {
                Write-Host "  No mode specified for desk '$deskName', defaulting to Offline" -ForegroundColor Green
                Set-PlaceV3 $desk.PlaceId -Mode @{Name="Offline"; Metadata=$modemetadata}
            }

        } catch {
            Write-Host "Error processing row: $($_.Exception.Message)" -ForegroundColor Red
            $errors++
        }
    }

    # Summary
    Write-Host "`nProcessing complete:" -ForegroundColor Cyan
    Write-Host "  Mailboxes converted: $mailboxesConverted" -ForegroundColor Blue
    Write-Host "  Desks created: $desksCreated" -ForegroundColor Green
    Write-Host "  Sections created: $sectionsCreated" -ForegroundColor Yellow
    Write-Host "  Skipped: $skipped" -ForegroundColor Yellow
    Write-Host "  Errors encountered: $errors" -ForegroundColor $(if ($errors -gt 0) {"Red"} else {"Green"})

    # Export results to CSV
    if ($results.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportPath = "DesksCreated_$timestamp.csv"
        $results | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "`nResults exported to: $exportPath" -ForegroundColor Green
    } else {
        Write-Host "`nNo desks were created, no results to export." -ForegroundColor Yellow
    }
}

# Function to create Non-bookable desks
function NonbookableDesks {
    ## article: https://learn.microsoft.com/en-us/microsoft-365/places/configure-desk-booking
    # Prompt for CSV file path
    $csvPath = Read-Host "Enter the path to your CSV file"

    # Validate file exists
    if (-not (Test-Path $csvPath)) {
        Write-Error "File not found: $csvPath"
        exit
    }

    $file = Import-Csv $csvPath

    # Validate required columns
    $requiredColumns = @('Building', 'FloorLabel', 'Section', 'SpaceID')
    $missingColumns = $requiredColumns | Where-Object { -not ($file[0].PSObject.Properties.Name -contains $_) }
    if ($missingColumns) {
        Write-Error "CSV is missing required columns: $($missingColumns -join ', ')"
        return
    }

    # Summary counters
    $desksCreated = 0
    $sectionsCreated = 0
    $errors = 0
    $skipped = 0

    # Create array to store results for export
    $results = @()

    # Caching to avoid redundant API calls
    $buildingCache = @{}
    $floorCache = @{}    # key: "<BuildingPlaceId>|<FloorLabel>"
    $sectionCache = @{}  # key: "<FloorPlaceId>|<SectionName>"
    
    Write-Host "Processing $($file.Count) desks from CSV..." -ForegroundColor Cyan

    foreach ($row in $file) {
        try {
            Write-Host "Processing desk $($row.SpaceID)..." -ForegroundColor Yellow
            
            # ------------- Building lookup with cache -------------
            $buildingName = $row.Building
            if ($buildingCache.ContainsKey($buildingName)) {
                $building = $buildingCache[$buildingName]
                Write-Host "  Using cached Building '$buildingName'" -ForegroundColor DarkGray
            } else {
                Write-Host "  Checking if Building '$buildingName' exists" -ForegroundColor Gray
                $building = Get-PlaceV3 -Type Building | Where-Object -Property DisplayName -eq $buildingName
                
                if ($null -eq $building) {
                    Write-Host "  Building '$buildingName' does not exist, please check" -ForegroundColor Red
                    $errors++
                    continue
                }
                
                if ($building -is [array]) {
                    Write-Host "  Multiple buildings found with name '$buildingName', using the first one" -ForegroundColor Yellow
                    $building = $building[0]
                }
                
                $buildingCache[$buildingName] = $building
                Write-Host "  Building cached for future use" -ForegroundColor Green
            }
            
            # ------------- Floor lookup with cache -------------
            $floorKey = "$($building.PlaceId)|$($row.FloorLabel)"
            if ($floorCache.ContainsKey($floorKey)) {
                $floor = $floorCache[$floorKey]
                Write-Host "  Using cached Floor '$($row.FloorLabel)'" -ForegroundColor DarkGray
            } else {
                Write-Host "  Checking if Floor '$($row.FloorLabel)' exists" -ForegroundColor Gray
                $floor = Get-PlaceV3 -Type Floor -AncestorId $building.PlaceId | 
                        Where-Object -Property DisplayName -eq $row.FloorLabel
                
                if ($null -eq $floor) {
                    Write-Host "  Floor '$($row.FloorLabel)' does not exist, please check" -ForegroundColor Red
                    $errors++
                    continue
                }
                
                if ($floor -is [array]) {
                    Write-Host "  Multiple floors found with name '$($row.FloorLabel)', using the first one" -ForegroundColor Yellow
                    $floor = $floor[0]
                }
                
                $floorCache[$floorKey] = $floor
                Write-Host "  Floor cached for future use" -ForegroundColor Green
            }
            
            # ------------- Section lookup with cache -------------
            $sectionName = $row.Section
            $sectionKey = "$($floor.PlaceId)|$sectionName"
            
            if ($sectionCache.ContainsKey($sectionKey)) {
                $section = $sectionCache[$sectionKey]
                Write-Host "  Using cached Section '$sectionName'" -ForegroundColor DarkGray
            } else {
                Write-Host "  Checking if Section '$sectionName' exists" -ForegroundColor Gray
                $section = Get-PlaceV3 -Type Section -AncestorId $floor.PlaceId | 
                        Where-Object -Property DisplayName -eq $sectionName
                
                if ($null -eq $section) {
                    $createSection = Read-Host "  Section '$sectionName' does not exist. Create it? (Y/N)"
                    
                    if ($createSection -eq "Y") {
                        Write-Host "  Creating section '$sectionName'" -ForegroundColor Yellow
                        $section = New-Place -Type Section -Name $sectionName -ParentId $floor.PlaceId
                        
                        if ($null -eq $section) {
                            Write-Host "  Failed to create section '$sectionName'" -ForegroundColor Red
                            $errors++
                            continue
                        } else {
                            Write-Host "  Section created successfully" -ForegroundColor Green
                            $sectionsCreated++
                            $sectionCache[$sectionKey] = $section
                        }
                    } else {
                        Write-Host "  Skipping desk, section not created" -ForegroundColor Yellow
                        $skipped++
                        continue
                    }
                } else {
                    if ($section -is [array]) {
                        Write-Host "  Multiple sections found with name '$sectionName', using the first one" -ForegroundColor Yellow
                        $section = $section[0]
                    }
                    $sectionCache[$sectionKey] = $section
                    Write-Host "  Section cached for future use" -ForegroundColor Green
                }
            }
            
            # ------------- Create desk with check for existing -------------
            $deskName = "Desk $($row.SpaceID)"
            Write-Host "  Creating desk '$deskName'..." -ForegroundColor Gray
            
            # Check if desk already exists
            $existingDesk = Get-PlaceV3 -Type Desk -AncestorId $section.PlaceId | 
                            Where-Object -Property DisplayName -eq $deskName
            
            if ($existingDesk) {
                Write-Host "  Desk '$deskName' already exists, skipping creation" -ForegroundColor Yellow
                $skipped++
                
                # Still add to results for tracking
                $results += [PSCustomObject]@{
                    DisplayName = $existingDesk.DisplayName
                    PlaceId = $existingDesk.PlaceId
                    ParentId = $existingDesk.ParentPlaceId
                    Type = "Desk"
                    Building = $row.Building
                    Floor = $row.FloorLabel
                    Section = $sectionName
                    SpaceID = $row.SpaceID
                    CreatedTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Status = "Skipped - Already Exists"
                }
            } else {
                # Set capacity to 1 by default (typical for a desk)
                $capacity = if ($row.PSObject.Properties.Name -contains "Capacity" -and $row.Capacity) { 
                    [int]$row.Capacity 
                } else { 
                    1 
                }
                
                # Create desk with retry logic
                $maxRetries = 2
                $retryCount = 0
                $success = $false
                
                while (-not $success -and $retryCount -le $maxRetries) {
                    try {
                        $desk = New-Place -Type Desk -Name $deskName -ParentId $section.PlaceId -Capacity $capacity
                        
                        if ($desk) {
                            Write-Host "  Desk '$deskName' created successfully with capacity $capacity" -ForegroundColor Green
                            $desksCreated++
                            $success = $true
                            
                            # Add desk info to results array
                            $results += [PSCustomObject]@{
                                DisplayName = $desk.DisplayName
                                PlaceId = $desk.PlaceId
                                ParentId = $desk.ParentPlaceId
                                Type = "Desk"
                                Building = $row.Building
                                Floor = $row.FloorLabel
                                Section = $sectionName
                                SpaceID = $row.SpaceID
                                Capacity = $capacity
                                CreatedTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                                Status = "Created"
                            }
                        } else {
                            throw "New-Place returned null"
                        }
                    } catch {
                        $retryCount++
                        if ($retryCount -le $maxRetries) {
                            Write-Host "  Attempt $retryCount failed, retrying in 3 seconds..." -ForegroundColor Yellow
                            Start-Sleep -Seconds 3
                        } else {
                            Write-Host "  Failed to create desk '$deskName' after $maxRetries retries" -ForegroundColor Red
                            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
                            $errors++
                        }
                    }
                }
            }
            
        } catch {
            Write-Host "Error processing row: $($_.Exception.Message)" -ForegroundColor Red
            $errors++
        }
        
        # Add progress indicator for large files
        if ($file.Count -gt 10 -and $desksCreated % 10 -eq 0) {
            Write-Host "Processed $($desksCreated + $skipped + $errors) of $($file.Count) entries..." -ForegroundColor Cyan
        }
    }

    # Summary
    Write-Host "`nProcessing complete:" -ForegroundColor Cyan
    Write-Host "  Desks created: $desksCreated" -ForegroundColor Green
    Write-Host "  Sections created: $sectionsCreated" -ForegroundColor Yellow
    Write-Host "  Skipped: $skipped" -ForegroundColor Yellow
    Write-Host "  Errors encountered: $errors" -ForegroundColor $(if ($errors -gt 0) {"Red"} else {"Green"})

    # Export results to CSV
    if ($results.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportPath = "NonBookableDesksCreated_$timestamp.csv"
        $results | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "`nResults exported to: $exportPath" -ForegroundColor Green
    } else {
        Write-Host "`nNo desks were created, no results to export." -ForegroundColor Yellow
    }
}

# Function to create DeskPool (Workspace) in Places
function WorkspaceORDeskpool {
        ## This script creates Workspace objects in Places, and makes them bookable (where mailbox is required)
        ## article: https://learn.microsoft.com/en-us/microsoft-365/places/configure-desk-booking

        Write-Host '######################################################## MAKE Sure you are Connect to Exchange Online & Places ########################################################' -ForegroundColor Yellow

        # Prompt for CSV file path
        $csvPath = Read-Host "Enter the path to your CSV file"

        # Validate file exists
        if (-not (Test-Path $csvPath)) {
            Write-Error "File not found: $csvPath"
            exit
        }

        $file = Import-Csv $csvPath

        # Validate required columns
        $requiredColumns = @('Building', 'FloorLabel', 'Section', 'Display Name', 'PrimarySmtpAddress')
        $missingColumns = $requiredColumns | Where-Object { -not ($file[0].PSObject.Properties.Name -contains $_) }
        if ($missingColumns) {
            Write-Error "CSV is missing required columns: $($missingColumns -join ', ')"
            exit
        }

        Write-Host "## This script creates Workspace objects in Places, and makes them bookable DeskPool (workspaces)" -ForegroundColor Cyan

        # Summary counters
        $workspacesCreated = 0
        $sectionsCreated = 0
        $errors = 0
        $skipped = 0
        $mailboxesConverted = 0

        # Create array to store results for export
        $results = @()

        # STEP 1: PRE-PROCESS ALL MAILBOXES FIRST
        Write-Host "`n== PHASE 1: CONVERTING MAILBOXES ==" -ForegroundColor Cyan

        $mailboxesToProcess = $file | Where-Object { ![string]::IsNullOrEmpty($_.PrimarySmtpAddress) } | Select-Object -ExpandProperty PrimarySmtpAddress -Unique
        $totalMailboxes = $mailboxesToProcess.Count

        Write-Host "Found $totalMailboxes unique mailboxes to convert to Workspace type" -ForegroundColor Yellow

        foreach ($mailbox in $mailboxesToProcess) {
            try {
                # Get current mailbox
                $mbx = Get-Mailbox $mailbox -ErrorAction Stop
                
                if ($mbx.Type -ne "Workspace") {
                    Write-Host "Converting mailbox '$mailbox' from type '$($mbx.Type)' to 'Workspace'" -ForegroundColor Gray
                    Set-Mailbox $mailbox -Type Workspace
                    $mailboxesConverted++
                    
                    # Set calendar processing for workspace
                    Write-Host "Setting calendar processing for '$mailbox'" -ForegroundColor Gray
                    Set-CalendarProcessing -Identity $mailbox -EnforceCapacity $True
                } else {
                    Write-Host "Mailbox '$mailbox' is already type 'Workspace'" -ForegroundColor Green
                    # Set calendar processing for workspace
                    Write-Host "Updating calendar processing for '$mailbox'" -ForegroundColor Gray
                    Set-CalendarProcessing -Identity $mailbox -EnforceCapacity $True
                }
            } catch {
                Write-Host "Error processing mailbox '$mailbox': $($_.Exception.Message)" -ForegroundColor Red
                $errors++
            }
        }

        # Give Exchange time to process the mailbox type changes
        if ($mailboxesConverted -gt 0) {
            $waitTime = 10  # Increased time for workspace mailbox processing
            Write-Host "`nConverted $mailboxesConverted mailboxes. Waiting $waitTime seconds for Exchange to process changes..." -ForegroundColor Yellow
            Start-Sleep -Seconds $waitTime
        }

        # STEP 2: PROCESS BUILDINGS, SECTIONS, AND WORKSPACES

        Write-Host "`n== PHASE 2: CREATING WORKSPACE STRUCTURE ==" -ForegroundColor Cyan

        # Caches to avoid repeated Get-PlaceV3 calls
        $buildingCache = @{}
        $floorCache    = @{}   # key: "<BuildingPlaceId>|<FloorLabel>"
        $sectionCache  = @{}   # key: "<FloorPlaceId>|<SectionName>"

        foreach ($row in $file) {
            try {
                Write-Host "Processing workspace $($row."Display Name")..." -ForegroundColor Yellow

                # ---------------- Building ----------------
                $buildingName = $row.Building
                if ($buildingCache.ContainsKey($buildingName)) {
                    $building = $buildingCache[$buildingName]
                    Write-Host "  Using cached Building '$buildingName'" -ForegroundColor DarkGray
                } else {
                    Write-Host "  Checking if Building '$buildingName' exists" -ForegroundColor Gray
                    $building = Get-PlaceV3 -Type Building | Where-Object DisplayName -eq $buildingName
                    if (-not $building) {
                        Write-Host "  Building '$buildingName' does not exist, please check" -ForegroundColor Red
                        $errors++
                        continue
                    }
                    if ($building -is [array]) {
                        Write-Host "  Multiple buildings named '$buildingName', using first" -ForegroundColor Yellow
                        $building = $building[0]
                    }
                    $buildingCache[$buildingName] = $building
                    Write-Host "  Cached Building '$buildingName'" -ForegroundColor Green
                }

                # ---------------- Floor ----------------
                $floorKey = "$($building.PlaceId)|$($row.FloorLabel)"
                if ($floorCache.ContainsKey($floorKey)) {
                    $floor = $floorCache[$floorKey]
                    Write-Host "  Using cached Floor '$($row.FloorLabel)'" -ForegroundColor DarkGray
                } else {
                    Write-Host "  Checking if Floor '$($row.FloorLabel)' exists" -ForegroundColor Gray
                    $floor = Get-PlaceV3 -Type Floor -AncestorId $building.PlaceId |
                            Where-Object { $_.Type -eq "Floor" -and $_.DisplayName -eq $row.FloorLabel }
                    if (-not $floor) {
                        Write-Host "  Floor '$($row.FloorLabel)' does not exist, please check" -ForegroundColor Red
                        $errors++
                        continue
                    }
                    if ($floor -is [array]) {
                        Write-Host "  Multiple floors named '$($row.FloorLabel)', using first" -ForegroundColor Yellow
                        $floor = $floor[0]
                    }
                    $floorCache[$floorKey] = $floor
                    Write-Host "  Cached Floor '$($row.FloorLabel)'" -ForegroundColor Green
                }

                # ---------------- Section ----------------
                $sectionName = $row.Section
                $sectionKey  = "$($floor.PlaceId)|$sectionName"
                if ($sectionCache.ContainsKey($sectionKey)) {
                    $section = $sectionCache[$sectionKey]
                    Write-Host "  Using cached Section '$sectionName'" -ForegroundColor DarkGray
                } else {
                    Write-Host "  Checking if Section '$sectionName' exists" -ForegroundColor Gray
                    $section = Get-PlaceV3 -Type Section -AncestorId $floor.PlaceId |
                            Where-Object DisplayName -eq $sectionName
                    if (-not $section) {
                        $createSection = Read-Host "  Section '$sectionName' does not exist. Create it? (Y/N)"
                        if ($createSection -eq "Y") {
                            Write-Host "  Creating section '$sectionName'" -ForegroundColor Yellow
                            $section = New-Place -Type Section -Name $sectionName -ParentId $floor.PlaceId
                            if (-not $section) {
                                Write-Host "  Failed to create section '$sectionName'" -ForegroundColor Red
                                $errors++
                                continue
                            }
                            Write-Host "  Section created successfully" -ForegroundColor Green
                            $sectionsCreated++
                        } else {
                            Write-Host "  Skipping workspace, section not created" -ForegroundColor Yellow
                            $skipped++
                            continue
                        }
                    } else {
                        if ($section -is [array]) {
                            Write-Host "  Multiple sections named '$sectionName', using first" -ForegroundColor Yellow
                            $section = $section[0]
                        }
                        Write-Host "  Section exists, proceeding" -ForegroundColor Green
                    }
                    $sectionCache[$sectionKey] = $section
                    Write-Host "  Cached Section '$sectionName'" -ForegroundColor Green
                }

                # ---------------- Workspace Creation ----------------
                Write-Host "  Linking mailbox '$($row.PrimarySmtpAddress)' to section..." -ForegroundColor Gray

                # Set capacity from CSV if available, otherwise default to 1
                $capacity = if ($row.PSObject.Properties.Name -contains "Capacity" -and ![string]::IsNullOrEmpty($row.Capacity)) { 
                    [int]$row.Capacity 
                } else { 
                    1
                }

                try {
                    Write-Host "  Linking mailbox '$($row.PrimarySmtpAddress)' to section and setting capacity $capacity" -ForegroundColor Gray
                    $workspace = Set-PlaceV3 -Identity $row.PrimarySmtpAddress -ParentId $section.PlaceId -Capacity $capacity -ErrorAction Stop
                    
                    if ($workspace) {
                        Write-Host "  Workspace linked successfully with capacity $capacity" -ForegroundColor Green
                        $workspacesCreated++
                    } else {
                        Write-Host "  Failed to link workspace for '$($row.PrimarySmtpAddress)'" -ForegroundColor Red
                        $errors++
                    }
                } catch {
                    Write-Host "  Error creating workspace: $($_.Exception.Message)" -ForegroundColor Red
                    $errors++
                    continue
                }

                $results += [PSCustomObject]@{
                    DisplayName  = $workspace.DisplayName
                    PlaceId      = $workspace.PlaceId
                    ParentId     = $workspace.ParentPlaceId
                    Type         = "Workspace"
                    Building     = $row.Building
                    Floor        = $row.FloorLabel
                    Section      = $sectionName
                    DisplayNameRow  = $row."Display Name"
                    Capacity     = $workspace.Capacity
                    CreatedTime  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Mailbox      = $row.PrimarySmtpAddress
                }

            } catch {
                Write-Host "Error processing row: $($_.Exception.Message)" -ForegroundColor Red
                $errors++
            }
        }

        # Summary
        Write-Host "`nProcessing complete:" -ForegroundColor Cyan
        Write-Host "  Mailboxes converted: $mailboxesConverted" -ForegroundColor Blue
        Write-Host "  Workspaces created: $workspacesCreated" -ForegroundColor Green
        Write-Host "  Sections created: $sectionsCreated" -ForegroundColor Yellow
        Write-Host "  Skipped: $skipped" -ForegroundColor Yellow
        Write-Host "  Errors encountered: $errors" -ForegroundColor $(if ($errors -gt 0) {"Red"} else {"Green"})

}

# function to select the type of object to create (Desk or Workspace)
function CreateDeskorWorkspace {
    Write-Host "Select the type of object you want to create:" -ForegroundColor Cyan
    Write-Host "Read more about Configure Desk bookings in Places: https://learn.microsoft.com/en-us/microsoft-365/places/configure-desk-booking"
    Write-Host "1 - Create Bookable Desk | These require Exchange mailbox to link them and make the desk bookable (connect-Exo* required)" -ForegroundColor Green
    Write-Host "2 - Create Non-Bookable Desk | Does not require Exchange mailbox but cannot be booked" -ForegroundColor Green
    Write-Host "3 - Create DeskPool (Called Workspace in Exchange) | This requires Exchange mailbox (Connect-Exo* required)" -ForegroundColor Green
    $choice = Read-Host "Enter your choice (1, 2, or 3):" -OutVariable choice

    switch ($choice) {
        "1" {
            Write-Host "Creating Bookable Desks..." -ForegroundColor Green
            BookableDesks
        }
        "2" {
            Write-Host "Creating Non-Bookable Desks..." -ForegroundColor Green
            NonbookableDesks
        }
        "3" {
            Write-Host "Creating DeskPool (Workspace)..." -ForegroundColor Green
            WorkspaceORDeskpool
        }
        Default {
            Write-Warning "Invalid choice. Please select 1, 2, or 3."
        }
    }
}

function SetDeskMode {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$NoExport
    )
    
    Write-Host "########################################################## SET DESK MODES ##########################################################" -ForegroundColor Yellow
    Write-Host "This function updates desk modes from a CSV file." -ForegroundColor Cyan
    
    # Validate file exists
    if (-not (Test-Path $CsvPath)) {
        Write-Error "File not found: $CsvPath"
        return
    }
    
    # Import CSV
    try {
        $desks = Import-Csv -Path $CsvPath
        Write-Host "Found $($desks.Count) desks in CSV file" -ForegroundColor Green
    } catch {
        Write-Error "Error importing CSV file: $($_.Exception.Message)"
        return
    }
    
    # Validate CSV structure - Display Name and Mode are now required
    $requiredColumns = @('Display Name', 'Mode')
    $missingColumns = $requiredColumns | Where-Object { -not ($desks[0].PSObject.Properties.Name -contains $_) }
    
    if ($missingColumns) {
        Write-Error "CSV is missing required columns: $($missingColumns -join ', ')"
        Write-Host "CSV must contain at least 'Display Name' and 'Mode' columns." -ForegroundColor Yellow
        return
    }
    
    # Check for AssignedTo column if there are Assigned mode entries
    $hasAssignedMode = $desks | Where-Object { $_.Mode -eq "Assigned" }
    if ($hasAssignedMode -and -not ($desks[0].PSObject.Properties.Name -contains 'AssignedTo')) {
        Write-Warning "Some desks have 'Assigned' mode but CSV is missing 'AssignedTo' column"
        Write-Host "Desks with Assigned mode require an AssignedTo email address" -ForegroundColor Yellow
    }
    
    # Create results array for tracking
    $results = @()
    $processedCount = 0
    $errorCount = 0
    $successCount = 0
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputPath = "DeskModeUpdate_$timestamp.csv"
    $validModes = @('Assigned', 'Bookable', 'Walkup', 'Offline')
    
    # Process each desk
    $totalDesks = $desks.Count
    foreach ($desk in $desks) {
        $processedCount++
        $progressParams = @{
            Activity = "Setting desk modes"
            Status = "Processing $processedCount of $totalDesks"
            PercentComplete = ($processedCount / $totalDesks * 100)
        }
        Write-Progress @progressParams
        
        Write-Host "[$processedCount/$totalDesks] Processing desk: $($desk.'Display Name')" -ForegroundColor Cyan
        
        try {
            # Lookup desk directly by Display Name
            Write-Host "  Looking up desk '$($desk.'Display Name')'..." -ForegroundColor Gray
            $deskObj = Get-PlaceV3 -Identity $desk.'Display Name' -ErrorAction SilentlyContinue
            
            if ($deskObj) {
                Write-Host "  Found desk: $($deskObj.DisplayName) (Type: $($deskObj.Type))" -ForegroundColor Green
                
                # Create metadata dictionary
                $modemetadata = New-Object 'System.Collections.Generic.Dictionary[String,object]'
                
                # Set mode based on provided parameters
                if (-not [string]::IsNullOrEmpty($desk.Mode)) {
                    if ($desk.Mode -eq "Assigned" -and -not [string]::IsNullOrEmpty($desk.AssignedTo)) {
                        Write-Host "  Setting desk '$($deskObj.DisplayName)' mode to 'Assigned' for user '$($desk.AssignedTo)'" -ForegroundColor Gray
                        $modemetadata.Add("AssignedPersonEmailAddress","$($desk.AssignedTo)")
                        
                        if ($PSCmdlet.ShouldProcess("Desk '$($deskObj.DisplayName)'", "Assign to user '$($desk.AssignedTo)'")) {
                            Set-PlaceV3 $deskObj.PlaceId -Mode @{Name="Assigned"; Metadata=$modemetadata} -ErrorAction Stop
                        }
                        
                        $modeResult = "Assigned to $($desk.AssignedTo)"
                        $successCount++
                        
                    } elseif ($validModes -contains $desk.Mode) {
                        Write-Host "  Setting desk '$($deskObj.DisplayName)' mode to '$($desk.Mode)'" -ForegroundColor Gray
                        
                        if ($PSCmdlet.ShouldProcess("Desk '$($deskObj.DisplayName)'", "Set mode to '$($desk.Mode)'")) {
                            Set-PlaceV3 $deskObj.PlaceId -Mode @{Name="$($desk.Mode)"; Metadata=$modemetadata} -ErrorAction Stop
                        }
                        
                        $modeResult = $desk.Mode
                        $successCount++
                        
                    } else {
                        Write-Host "  Unrecognized mode '$($desk.Mode)' for desk '$($deskObj.DisplayName)', defaulting to Offline" -ForegroundColor Yellow
                        
                        if ($PSCmdlet.ShouldProcess("Desk '$($deskObj.DisplayName)'", "Set mode to 'Offline' (default)")) {
                            Set-PlaceV3 $deskObj.PlaceId -Mode @{Name="Offline"; Metadata=$modemetadata} -ErrorAction Stop
                        }
                        
                        $modeResult = "Offline (default)"
                        $successCount++
                    }
                } else {
                    Write-Host "  No mode specified for desk '$($deskObj.DisplayName)', defaulting to Offline" -ForegroundColor Yellow
                    
                    if ($PSCmdlet.ShouldProcess("Desk '$($deskObj.DisplayName)'", "Set mode to 'Offline' (no mode specified)")) {
                        Set-PlaceV3 $deskObj.PlaceId -Mode @{Name="Offline"; Metadata=$modemetadata} -ErrorAction Stop
                    }
                    
                    $modeResult = "Offline (default)"
                    $successCount++
                }
                
                $status = "Success"
                $errorDetails = $null
                
            } else {
                Write-Warning "  Desk not found: '$($desk.'Display Name')'"
                $status = "Not Found"
                $errorDetails = "Desk not found in Places"
                $errorCount++
            }
            
            # Add result to tracking collection
            $results += [PSCustomObject]@{
                DisplayName = $desk.'Display Name'
                PlaceId = if ($null -ne $deskObj) { $deskObj.PlaceId } else { $null }
                RequestedMode = $desk.Mode
                AppliedMode = $modeResult
                AssignedTo = $desk.AssignedTo
                Status = $status
                Error = $errorDetails
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            
            Write-Host "  Status: $status" -ForegroundColor $(if ($status -eq "Success") {"Green"} else {"Yellow"})
            
        } catch {
            $errorCount++
            Write-Warning "  Error processing desk '$($desk.'Display Name')': $($_.Exception.Message)"
            
            # Add error result to tracking
            $results += [PSCustomObject]@{
                DisplayName = $desk.'Display Name'
                PlaceId = if ($null -ne $deskObj) { $deskObj.PlaceId } else { $null }
                RequestedMode = $desk.Mode
                AppliedMode = $null
                AssignedTo = $desk.AssignedTo
                Status = "Failed"
                Error = $_.Exception.Message
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }
    
    Write-Progress -Activity "Setting desk modes" -Completed
    
    # Display results summary
    Write-Host "`nResults summary:" -ForegroundColor Cyan
    Write-Host "  Total desks processed: $processedCount" -ForegroundColor Blue
    Write-Host "  Successful: $successCount" -ForegroundColor Green
    Write-Host "  Failed/Not Found: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) {"Red"} else {"Green"})
    
    # Export results unless -NoExport was specified
    if (-not $NoExport -and $results.Count -gt 0) {
        $results | Export-Csv -Path $outputPath -NoTypeInformation
        Write-Host "Results exported to: $outputPath" -ForegroundColor Green
    }
    
    # Display results table
    $results | Format-Table DisplayName, RequestedMode, AppliedMode, Status -AutoSize
}

function UploadBuildingMap {
    Write-Host "########################################################## UPLOAD BUILDING MAP ##########################################################" -ForegroundColor Yellow
    Write-Host "This script uploads a building map to Microsoft Places. Ensure that you are connected to Microsoft Places." -ForegroundColor Cyan

    # Initialize building ID
    $buildingPlaceId = $null

    # Prompt for Building PlaceId or search
    $initialInput = Read-Host "Enter the Building PlaceId where you want to upload the map, or type 'find' to search for a building with DisplayName"

    # Process based on input
    if ($initialInput -eq "find") {
        $bld = Read-Host "Enter the Building Name to search for PlaceId"
        $buildings = Get-PlaceV3 -Type Building | Where-Object { $_.DisplayName -like "$bld" }
        
        if (-not $buildings) {
            Write-Error "No buildings found matching '$bld'"
            return
        }
        
        $buildings | Format-Table DisplayName, PlaceId, City, CountryOrRegion -AutoSize
        Write-Host "Please copy the PlaceId of the building you want to use." -ForegroundColor Yellow
        $buildingPlaceId = Read-Host "Enter the Building PlaceId"
    }
    else {
        $buildingPlaceId = $initialInput
    }

    # Validate Building PlaceId is not empty
    if ([string]::IsNullOrWhiteSpace($buildingPlaceId)) {
        Write-Error "Building PlaceId cannot be empty."
        return
    }
    
    # Verify the building exists
    try {
        $building = Get-PlaceV3 -Identity $buildingPlaceId -ErrorAction Stop
        if (-not $building) {
            Write-Error "No building found with PlaceId: $buildingPlaceId"
            return
        }
        Write-Host "Found building: $($building.DisplayName), Type: $($building.type), City: $($building.City), Country: $($building.CountryOrRegion)" -ForegroundColor Green
    }
    catch {
        Write-Error "Invalid Building PlaceId: $buildingPlaceId. Error: $($_.Exception.Message)"
        return
    }

    # Prompt for file path
    $filePath = Read-Host "Enter the path to your building map file (paste full path without QUOTES, e.g. C:\path\to\your\file.zip)"

    # Validate file exists
    if (-not (Test-Path $filePath)) {
        Write-Error "File not found: $filePath"
        return
    }

    # Validate file extension (optional)
    $extension = [System.IO.Path]::GetExtension($filePath)
    if ($extension -notmatch '\.(zip|png|jpg|jpeg|svg)$') {
        Write-Warning "File extension '$extension' may not be supported. Recommended formats: .zip, .png, .jpg, .svg"
        $confirm = Read-Host "Do you want to continue anyway? (Y/N)"
        if ($confirm -ne "Y") {
            Write-Host "Upload cancelled." -ForegroundColor Yellow
            return
        }
    }

    try {
        # Upload the building map
        Write-Host "Uploading building map to building: $($building.DisplayName) (PlaceId: $buildingPlaceId)..." -ForegroundColor Gray
        New-Map -BuildingId $buildingPlaceId -FilePath $filePath -ErrorAction Stop
        Write-Host "Building map uploaded successfully!" -ForegroundColor Green
    } 
    catch {
        Write-Error "Failed to upload building map: $($_.Exception.Message)"
        Write-Host "You might want to check if a map for that building already exists or not" -ForegroundColor Yellow
    }
}

# Help Function
function help {
    Write-Host "##################################################### Help ##############################################################################" -ForegroundColor Green
    Write-Host "Required permissions depend on the operation. Use the appropriate Microsoft Places and Exchange Online administrative roles; activate PIM access if your organization requires it." `n
    Write-Host "See Readme.md for setup, CSV examples, automation, and safety guidance."
    Write-Host "Microsoft Places Learn articles: https://learn.microsoft.com/en-us/microsoft-365/places/places-overview" `n
    Write-Host "Desk booking documentation: https://learn.microsoft.com/en-us/microsoft-365/places/configure-desk-booking" `n
    Write-Host "##################################################### Help ##############################################################################" -ForegroundColor Green
}

# Main choice block function
function Mainblock {
 do {
    Show-Menu `n
    $choice = Read-Host "Enter your choice"
    if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "0" }

    switch ($choice) {
        "0" { connectPlaces }
        "1" { FindBuildingOrFloor }
        "2" { OnboardRoom -FileName (Read-Host "Enter the CSV file path or type sample to get a samplecsv") }
        "3" { DeskPoolAttri -InputFile (Read-Host "Enter the CSV file path, !Check your file should be in CSV!") }
        "4" { NewbldgFloor }
        "5" { AutoRelease }
        "6" { BSSID -InputFile (Read-Host "Enter the CSV file path") }
        "7" { Breakfix }
        "8" { CreateDeskorWorkspace }
        "9" { UploadBuildingMap }
        "10" { RemoveBuildingOptionMenu }
        "11" { RemoveEntireBuilding -BuildingPlaceId (Read-Host "Enter the Building PlaceId") }
        "12" { SetDeskMode }
        "help" { help }
        "X" {
            # Stop transcript and prompt to delete it
            Stop-Transcript
            Write-Host "Transcript stopped. Log saved to: $transcriptFile" -ForegroundColor Yellow
            exit
        }
        Default { Write-Warning "Invalid choice." }
    }
 } while ($true)

}

if (-not $NoMenu) {
    $transcriptFile = "Script_Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    Start-Transcript -Path $transcriptFile
    Mainblock
}