function Set-CMYearlyMaintenanceWindow {
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory, Position = 0)]
        [int]$Year,

        [Parameter(Mandatory, Position = 1)]
        [string]$CMServer
    )

    begin {
        #region Begin block

        #*=========================================
        #* Parameter Checks
        #*=========================================

        # Make sure the year parameter is not in the past
        $currentYear = (Get-Date).Year
        if ($Year -lt $currentYear) {
            Write-Warning -Message 'Year must be current or in the future.'
            continue
        }

        # If the year is the current year, start from the current month
        # Otherwise start from January
        if ($Year -eq $currentYear) {
            $startMonth = (Get-Date).Month
        } else {
            $startMonth = 1
        }


        #*=========================================
        #* Functions
        #*=========================================

        # Function to get the second Tuesday of the specific month and year
        function Get-SecondTuesday {
            param(
                [Parameter(Position = 0)]
                [int]$Month,

                [Parameter(Position = 1)]
                [int]$Year
            )

            # Use the current month and year if the user didn't specify anything
            if (-not $Month) { $Month = (Get-Date).Month }
            if (-not $Year) { $Year = (Get-Date).Year }

            # Cycle through the days of the month until we reach the first Tuesday
            $dayOfMonth = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
            while ($dayOfMonth.DayOfWeek -ne [System.DayOfWeek]::Tuesday) { $dayOfMonth = $dayOfMonth.AddDays(1) }

            # Add 7 days for the second Tuesday
            $secondTuesday = $dayOfMonth.AddDays(7)
            return $secondTuesday
        }

        #*=========================================
        #* Variables
        #*=========================================

        # Load FMG variables file
        $fmgVariablesFile = '\\fmg.local\itfiles\Operations\Scripts\PowerShellAdminScripts\Prod_Variables.ps1'
        try {
            . $fmgVariablesFile
        } catch {
            Write-Host -ForegroundColor Red 'Failed to load FMG Prod Variables - Unable to continue'
            Write-Host -ForegroundColor Red $_.Exception.Message
            continue
        }

        # Patch windows start times and offset days from patch tuesday
        $patchWindows = @{
            'DevTest'    = @{
                OffsetDays = 7
                StartTime  = '00:00'
            }
            'GIS'        = @{
                OffsetDays = 9
                StartTime  = '18:30'
            }
            'Production' = @{
                OffsetDays = 9
                StartTime  = '22:00'
            }
            'Manual'     = @{
                OffsetDays = 14
                StartTime  = '18:00'
            }
        }

        #*=========================================
        #* Connect to SCCM
        #*=========================================
        try {
            Write-Host -NoNewline -ForegroundColor Cyan 'Connecting to SCCM Server '
            Write-Host -NoNewline -ForegroundColor Yellow $CMServer
            Write-Host -NoNewline -ForegroundColor Cyan ' ... '

            if ($null -eq (Get-Module ConfigurationManager)) { Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" }
            if ($null -eq (Get-PSDrive -Name $FMG.SCCM.SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name $FMG.SCCM.SiteCode -PSProvider CMSite -Root $FMG.SCCM.Server -EA Stop
            }
            Set-Location "$($FMG.SCCM.SiteCode):" -EA Stop
            Write-Host -ForegroundColor Green 'Success.'

        } catch {
            Write-Host -ForegroundColor Red 'Failed.'
            Write-Host -ForegroundColor Red $_.Exception.Message
        }

        # Get the maintenance windows device collections from SCCM
        $allMaintenanceCollections = @()
        foreach ($window in $patchWindows.Keys) {
            $maintenanceCollectionWildcardName = "Servers | Maintenance | $window | *"
            $allMaintenanceCollections += Get-CMCollection -Name $maintenanceCollectionWildcardName
        }

        Write-Host ''

        #endregion
    }

    process {
        #region Process block

        try {


            #*=========================================
            #* Set the Maintenance Windows
            #*=========================================
            # Iterate through the months of the year selected
            for ($i = $startMonth; $i -le 12; $i++) {

                Write-Host -ForegroundColor Blue (Get-Date -Month $i).ToString('MMMM')

                # Get the second Tuesday of the months (patch Tuesday)
                $patchTuesday = Get-SecondTuesday -Month $i -Year $Year

                # Iterate through each window and get the date and duration
                foreach ($collection in $allMaintenanceCollections) {

                    #$allMaintenanceCollections.Name
                    #$collection = $allMaintenanceCollections[0]

                    # Extract the window name and time range
                    $windowCategory = ($collection.Name -split '\|')[2].Trim()
                    $windowTimeRange = ($collection.Name -split '\|')[-1].Trim()

                    # Split the time portion into start and end time
                    $startAndEndTime = ($windowTimeRange -split '-').Trim()
                    $startTime = [DateTime]::ParseExact($startAndEndTime[0], 'HH:mm', $null)
                    $endTime = [DateTime]::ParseExact($startAndEndTime[1], 'HH:mm', $null)

                    # Calculate the exact start date and time based on the offset days and start time
                    $patchDate = $patchTuesday.AddDays($patchWindows.$windowCategory.OffsetDays)
                    $patchTime = [DateTime]::ParseExact($patchWindows.$windowCategory.StartTime, 'HH:mm', $null)
                    $patchStart = Get-Date -Year $patchDate.Year -Month $patchDate.Month -Day $patchDate.Day -Hour $patchTime.Hour -Minute $patchTime.Minute -Second 0 -Millisecond 0

                    # Combine the window date with the start time
                    $windowStartTime = Get-Date -Year $patchStart.Year -Month $patchStart.Month -Day $patchStart.Day -Hour $startTime.Hour -Minute $startTime.Minute -Second 0 -Millisecond 0
                    $windowEndTime = Get-Date -Year $patchStart.Year -Month $patchStart.Month -Day $patchStart.Day -Hour $endTime.Hour -Minute $endTime.Minute -Second 0 -Millisecond 0

                    # If the start time is earlier than the patch window start time, it means it's the next day
                    if ($windowStartTime -lt $patchStart) {
                        $windowStartTime = $windowStartTime.AddDays(1)
                    }

                    # If the end time is earlier than the start time, it means it's the next day
                    if ($windowEndTime -lt $windowStartTime) {
                        $windowEndTime = $windowEndTime.AddDays(1)
                    }

                    # Generate a maintenance window name based on start date and time
                    $windowName = $windowStartTime.ToString('MMMM yyyy')

                    # Create new maintenance window on collection
                    try {

                        Write-Host -NoNewline -ForegroundColor Cyan 'Maintenance window: '
                        Write-Host -NoNewline -ForegroundColor Yellow $windowName
                        Write-Host -NoNewline -ForegroundColor Cyan ' on collection '
                        Write-Host -NoNewline -ForegroundColor Yellow $collection.Name
                        Write-Host -NoNewline -ForegroundColor Cyan '... '

                        $windowExists = Get-CMMaintenanceWindow -CollectionName $collection.Name -Verbose:$false | Where-Object Name -EQ $windowName
                        $windowSchedule = New-CMSchedule -Nonrecurring -Start $windowStartTime -End $windowEndTime

                        if ($windowExists) {
                            <#
                            $setMaintenanceWindow = @{
                                CollectionName        = $collection.Name
                                MaintenanceWindowName = $windowName
                                IsEnabled             = $true
                                Schedule              = $windowSchedule
                                ErrorAction           = 'Stop'
                            }
                            $null = Set-CMMaintenanceWindow @setMaintenanceWindow -ApplyTo SoftwareUpdatesOnly
                            #>
                            Write-Host -ForegroundColor Gray 'Already exists.'
                        } else {
                            $newMaintenanceWindow = @{
                                CollectionName = $collection.Name
                                Name           = $windowName
                                Schedule       = $windowSchedule
                                IsEnabled      = $true
                                ErrorAction    = 'Stop'
                            }
                            $null = New-CMMaintenanceWindow @newMaintenanceWindow -ApplyTo SoftwareUpdatesOnly
                            Write-Host -ForegroundColor Green 'Created.'
                        }
                    } catch {
                        Write-Host -NoNewline -ForegroundColor Red 'Failed. '
                        Write-Host -ForegroundColor Red $_.Exception.Message
                    }

                } #foreach_collection

                Write-Host
                Start-Sleep -Seconds 2
            } #for_month_loop


        } catch {
            # Catch all for unknown errors
            $PSCmdlet.ThrowTerminatingError($PSItem)
        }

        #endregion
    }
}

<#
#Testing
$Year = 2025
$i = 3
foreach ($collection in $allMaintenanceCollections) {$collection | Remove-CMMaintenanceWindow -MaintenanceWindowName '*' -Force -ErrorAction SilentlyContinue}
#>
