<#
.SYNOPSIS
    Parallel file copying utility using RoboCopy for faster transfers.

.DESCRIPTION
    This script dramatically speeds up large file transfers by running multiple RoboCopy operations simultaneously.

    How it works:
    - Instead of copying all files with a single RoboCopy command (which is slow for many small files)
    - This script breaks the job into smaller parts (one job per subfolder)
    - Each part runs as a separate background process
    - By working on multiple folders in parallel, the overall copy completes much faster

    The script handles both files in the source root directory and all subdirectories.

.PARAMETER NoConfirm
    If specified, skips the confirmation prompt that shows RoboCopy options before starting.
    Useful for scheduled or automated tasks where no user interaction is possible.

.EXAMPLE
    .\Robocopy_Faster.ps1
    Runs with interactive confirmation of RoboCopy options

.EXAMPLE
    .\Robocopy_Faster.ps1 -NoConfirm
    Runs without prompting for confirmation of RoboCopy options

.NOTES
    Author: Richard Tetlow
    Requirements: Windows PowerShell 5.1 or later
#>

param (
    [switch]$NoConfirm
)

# ==============================
# EDIT THESE SETTINGS AS NEEDED
# ==============================

# Source folder containing files to be copied
# Example: 'C:\Data\ProjectFiles'
$src = 'C:\Source'

# Target destination path where files will be copied to
# Example: '\\server\share\Backup'
$dest = '\\Path\To\Destination'

# Folders to exclude from copying (leave empty if none)
# Example: 'Temp', 'Cache', 'Old Projects'
$excludeDir = ''

# Directory where log files will be stored
# A new log is created for each folder being copied
$logDir = 'C:\Temp\Robocopy\Logs'

# Maximum number of parallel RoboCopy jobs
# Increase for faster copying but higher system load
# Decrease if experiencing performance issues
$maxJobs = 10

# Order folders by most recent modification date first
# Set to $true to copy newest folders first (recommended for prioritizing recent work)
# Set to $false to copy in alphabetical order
$newestFirst = $true

# ==============================
# ROBOCOPY COMMAND PARAMETERS
# ==============================
# This section defines the exact behavior of the RoboCopy commands
# Lines starting with '::' are inactive/commented out options
# Each option has an explanation after the '::'
# Advanced users can modify these parameters to customize copying behavior
$roboCopyOptions = @"
:: Do not remove these or the script will break.
	/NOSD			:: No Source directory is specified. To Be Specified on the command Line
	/NODD			:: No Destination directory is specified. To Be Specified on the command Line

:: Include These Files :
	/IF				:: Include only files matching the specified patterns
		*.*			:: In this case, include all files (wildcards are acceptable)
::	/IM				:: Include modified files (differing change times).
::	/IS				:: Includes the same files. Same files are identical in name, size, times, and all attributes.
::	/IT				:: Includes 'tweaked' files that differ by attributes.


:: Exclusions
	/XD				:: Exclude certain directories by name/path
		'DfsrPrivate'
		'System Volume Information'
		'Recycle Bin'
		'`$RECYCLE.BIN'
		'Sysmon'

	/XF				:: Exclude certain file names or patterns
		Thumbs.db	:: Directory thumbnails
		*.tmp		:: Temporary files
		*.lock		:: ArcGIS lock files
		*.isis_lock	:: Vulcan lock files
	/XJ				:: Exclude symbolic links/junction points to prevent unintended copies.
::	/XX				:: Excludes extra files and directories present in the destination but not the source.
::	/XO				:: Source directory files older than the destination are excluded from the copy.
::	/XN				:: Source directory files newer than the destination are excluded from the copy.
::	/XC				:: Excludes existing files with the same timestamp, but different file sizes.

:: Copy options
	/E				:: Copy subdirectories, including empty ones.
	/PURGE			:: Remove destination files/dirs that no longer exist in the source.
	/COPY:DATS		:: Specifies which file properties to copy. (D=Data, A=Attributes, T=Time, S=Security, O=Owner info, U=Audit info)
::	/COPYALL		:: Equivalent to /COPY:DATSOU (copies everything).
	/DCOPY:DT		:: Specifies which directory properties to copy. (D=Data, A=Attributes, T=Time)
::	/A-:SH			:: Removes the specified attributes from copied files.
::	/ZB				:: Use restartable mode; if access is denied, use backup mode.
	/B				:: Use backup mode to copy files (requires admin privileges).
	/MT:32			:: Use multi-threaded copies with n threads (default is 8).
::	/SECFIX			:: Fix file security on all files, even skipped files.
::	/TIMFIX			:: Fix file times on all files, even skipped files.
::	/FFT			:: Assumes FAT file times (2-second granularity) for compatibility with other systems.

:: Retry options
	/R:0			:: Retry failed copies up to 3 times.
	/W:1			:: Wait 1 second between retries.

:: Logging options
	/NDL			:: No Directory List - don't log directory names.
	/NFL			:: No File List - don't log file names.
	/NP				:: No Progress - don't display percentage progress.
"@

# ==============================
# UTILITY FUNCTIONS
# ==============================
# These functions support the main script operations

# Function: Get-FunctionScriptBlock
# Purpose: Captures function definitions so they can be used in background jobs
# This is needed because background jobs run in separate PowerShell instances
# and don't automatically have access to functions defined in the main script
function Get-FunctionScriptBlock {
    param(
        [string[]]$Function  # Names of functions to convert into transferable script blocks
    )

    # Create a string containing the full function definitions
    [string]$functionBlock = foreach ($func in $Function) {
        "function $((Get-Item function:\$func).Name) {$((Get-Item function:\$func).ScriptBlock)};"
    }
    $functionBlock
}


# Function: Select-RoboSummary
# Purpose: Extracts and parses the summary statistics from RoboCopy log files
# This makes it possible to display concise information about what was copied
function Select-RoboSummary {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$log, # The RoboCopy log content to parse

        [parameter(Mandatory = $false, ValueFromPipeline = $false)]
        [switch]$separateUnits  # Whether to separate numbers from their units (e.g., "10 MB" -> "10" and "MB")
    )
    begin {
        # Column headers from RoboCopy summary table (these match RoboCopy's output format)
        $cellHeaders = @('Total', 'Copied', 'Skipped', 'Mismatch', 'Failed', 'Extras')
        # Row types from RoboCopy summary table
        $rowTypes = @('Dirs', 'Files', 'Bytes')
    }
    process {
        # Find lines containing summary statistics using regular expression pattern matching
        $rows = $log | Select-String -Pattern '(Dirs|Files|Bytes)\s*:(\s*([0-9]+(\.[0-9]+)?( [a-zA-Z]+)?)+)+' -AllMatches

        # Validate that we found the expected summary information
        if ($rows.Count -eq 0) {
            throw 'Summary table not found in log {0}' -f $log
        }
        if ($rows.Matches.Count -ne $rowTypes.Count) {
            throw 'Unexpected number of rows in summary. Expected {0}, found {1}' -f $rowTypes.Count, $rowsMatch.Count
        }

        # Process each row in the summary table (Dirs, Files, Bytes)
        for ($x = 0; $x -lt $rows.Matches.Count; $x++) {
            $rowType = $rowTypes[$x]
            # Extract cell values from the current row
            $rowCells = $rows.Matches[$x].Groups[2].Captures | ForEach-Object { $_.ToString().Trim() }

            # Verify correct number of columns
            if ($cellHeaders.Length -ne $rowCells.Count) {
                throw 'Unexpected number of columns in the summary row. Expected {0} but found {1}' -f $cellHeaders.Length, $rowCells.Count
            }

            # Create an object to hold the parsed data
            $row = New-Object -TypeName PSObject
            $row | Add-Member -Type NoteProperty -Name Type -Value $rowType

            # Add each column value as a property to the object (Total, Copied, Skipped, etc.)
            for ($i = 0; $i -lt $rowCells.Count; $i++) {
                $header = $cellHeaders[$i]
                $cell = $rowCells[$i]

                # Split units from values if requested (e.g., "10 MB" -> "10" and "MB")
                if ($separateUnits -and ($cell -match ' ')) {
                    $cell = $cell -split ' '
                }

                $row | Add-Member -Type NoteProperty -Name $header -Value $cell
            }

            # Output the row object
            $row
        }
    }
}


# Function: Get-RoboCopyOptions
# Purpose: Parse RoboCopy options string and return structured option information
function Get-RoboCopyOptions {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$OptionsString,

        [Parameter(Mandatory = $false)]
        [switch]$ActiveOnly
    )

    process {
        $options = [System.Collections.Generic.List[PSCustomObject]]::new()

        # Process each line of the RoboCopy options
        foreach ($line in ($OptionsString.Split([Environment]::NewLine).Trim() | Where-Object { -not ([string]::IsNullOrWhiteSpace($_)) })) {
            # Skip header and non-option lines if they don't start with '::' or '/'
            if (-not ($line.StartsWith('::') -or $line.StartsWith('/')) -and -not $options.Count) {
                continue
            }

            # Check if option is active (not commented out)
            $isActive = -not $line.TrimStart().StartsWith('::')

            # Skip inactive options if ActiveOnly is specified
            if ($ActiveOnly -and -not $isActive) {
                continue
            }

            # Split the line into the option and its explanation
            [string[]]$splitLine = ($line -replace '^::\s*', '').Split('::').Trim() |
                Where-Object { -not ([string]::IsNullOrWhiteSpace($_)) }

            # If this is a RoboCopy command switch
            if ($splitLine[0].StartsWith('/')) {
                $optionObj = [PSCustomObject]@{
                    Option      = $splitLine[0]
                    Active      = $isActive
                    Explanation = if ($splitLine.Count -gt 1) { $splitLine[1] } else { '' }
                    Arguments   = [System.Collections.Generic.List[string]]::new()
                }
                $options.Add($optionObj)
            }
            # Otherwise it might be an additional argument for the previous option
            elseif ($options.Count -gt 0) {
                if ($isActive) {
                    $lastOption = $options[-1]
                    $lastOption.Arguments.Add($splitLine[0])
                }
            }
        }

        # For ActiveOnly mode, return a simple array of option strings with their arguments
        if ($ActiveOnly) {
            $activeOptions = [System.Collections.Generic.List[string]]::new()
            foreach ($opt in $options) {
                if ($opt.Arguments.Count -gt 0) {
                    $activeOptions.Add("$($opt.Option) $($opt.Arguments -join ' ')")
                } else {
                    $activeOptions.Add($opt.Option)
                }
            }
            $activeOptions.ToArray()
        } else {
            # Return the full structured options
            $options.ToArray()
        }
    }
}

# Function: Write-JobStats
# Purpose: Displays color-coded statistics about a completed RoboCopy operation
# Shows how many files were copied, skipped, failed, etc. with appropriate colors
function Write-JobStats {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [PSObject]$JobStats, # Statistics object from the RoboCopy operation

        [Parameter(Mandatory = $false, Position = 1)]
        [DateTime]$JobStart, # When the job started (for duration calculation)

        [Parameter(Mandatory = $false, Position = 2)]
        [string]$RoboOptionsString = $null   # RoboCopy options string to parse
    )

    Write-Host -NoNewline -ForegroundColor White '('

    # Calculate and display job duration (how long the copy took)
    if ($JobStart) {
        $jobDuration = New-TimeSpan -Start $JobStart -End (Get-Date)
        Write-Host -NoNewline -ForegroundColor Yellow $('{0:dd\d\.hh\:mm\:ss}' -f $jobDuration)
    }

    # Check if the PURGE option is active and XO is not active
    $isPurgeActive = $false
    $isXOActive = $false

    if ($RoboOptionsString) {
        $activeOptions = Get-RoboCopyOptions -OptionsString $RoboOptionsString -ActiveOnly
        $isPurgeActive = $activeOptions -contains '/PURGE' -or ($activeOptions | Where-Object { $_ -match '^/PURGE\b' })
        $isXOActive = $activeOptions -contains '/XO' -or ($activeOptions | Where-Object { $_ -match '^/XO\b' })
    }

    # Display file statistics with appropriate colors
    # - Green for successfully copied files
    if (($JobStats.Copied -as [int])) {
        Write-Host -NoNewline -ForegroundColor Green " Copied: $($JobStats.Copied)"
    }

    # Handle Extras (deleted files) appropriately based on options
    if (($JobStats.Extras -as [int])) {
        # Only show as "Deleted" if using /PURGE without /XO
        if ($isPurgeActive -and -not $isXOActive) {
            Write-Host -NoNewline -ForegroundColor DarkGreen " Deleted: $($JobStats.Extras)"
        } else {
            # Otherwise, add to skipped count
            $totalSkipped = ([int]$JobStats.Skipped) + ([int]$JobStats.Extras)
            $JobStats.Skipped = $totalSkipped
            # Don't display Extras separately since we've added them to Skipped
        }
    }

    # - Yellow for skipped files (already exist and are identical)
    if (($JobStats.Skipped -as [int])) {
        Write-Host -NoNewline -ForegroundColor Yellow " Skipped: $($JobStats.Skipped)"
    }

    # - Red for failed files (couldn't be copied due to errors)
    if (($JobStats.Failed -as [int])) {
        Write-Host -NoNewline -ForegroundColor Red " Failed: $($JobStats.Failed)"
    }

    Write-Host -ForegroundColor White ')'
}


# Function: Show-RunningJob
# Purpose: Displays information about currently running RoboCopy jobs
# Useful for long-running operations to see progress
function Show-RunningJob {
    # Get all jobs that are still running
    $runningJobs = Get-Job -State 'Running'
    $currentTime = Get-Date

    if ($runningJobs) {
        # Display a header showing how many jobs are still running
        Write-Host -ForegroundColor Yellow ('[{0:dd\/MM\/yy HH\:mm}] {1} robocopy still running:' -f $currentTime, $runningJobs.Count)

        # Show each running job with its name and duration
        foreach ($job in $runningJobs) {
            Write-Host -NoNewline -ForegroundColor Yellow $job.Name
            Write-Host (' (Duration: {0:dd\d\.hh\:mm\:ss})' -f (New-TimeSpan -Start $job.PSBeginTime -End $currentTime))
        }

        Write-Host
    }
}

# ==============================
# VALIDATE PATHS BEFORE STARTING
# ==============================
# This section checks that all directories exist before attempting to copy files

# Verify the source directory exists
if (-not (Test-Path -Path $src)) {
    Write-Host -ForegroundColor Red "Source directory invalid: $src"
    return
}

# Verify the destination directory exists
if (-not (Test-Path -Path $dest)) {
    Write-Host -ForegroundColor Red "Destination directory invalid: $dest"
    return
}

# Create the log directory if it doesn't exist
if (-not (Test-Path -Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    Write-Host -ForegroundColor Cyan "Created log directory: $logDir"
}

# ==============================
# CONFIRM SETTINGS WITH USER
# ==============================
# This section shows the user what options will be used and asks for confirmation

# Display RoboCopy options and request confirmation (unless -NoConfirm was specified)
if (-not ($NoConfirm)) {
    # Parse the RoboCopy options into a structured format for display
    $options = Get-RoboCopyOptions -OptionsString $roboCopyOptions

    # Show the information to the user
    Write-Host 'Please check that the RoboCopy options are correct before starting.'
    Write-Host "`nSource: $src"
    Write-Host "Destination: $dest"

    # Display options in a formatted way
    foreach ($opt in $options) {
        if ($opt.Active) {
            Write-Host -NoNewline -ForegroundColor Green 'ACTIVE: '
        } else {
            Write-Host -NoNewline -ForegroundColor Gray 'DISABLED: '
        }

        Write-Host -NoNewline $opt.Option

        # Add arguments if any exist
        if ($opt.Arguments.Count -gt 0) {
            Write-Host -NoNewline " $($opt.Arguments -join ' ')"
        }

        # Add explanation if available
        if (-not [string]::IsNullOrWhiteSpace($opt.Explanation)) {
            Write-Host -NoNewline ' :: '
            Write-Host -ForegroundColor Cyan $opt.Explanation
        } else {
            Write-Host
        }
    }

    # Prompt for confirmation to proceed
    if ((Read-Host 'Are these correct? Y/[N]') -ne 'Y') {
        Write-Host 'Operation cancelled by user.' -ForegroundColor Yellow
        return
    }
}

# ==============================
# PREPARE FOR COPY OPERATIONS
# ==============================
# Create a RoboCopy job file to be used by all copy operations
# A job file contains all the RoboCopy parameters so they don't need to be specified each time
$jobFile = Join-Path $logDir 'robocopy_options.rcj'
$roboCopyOptions | Set-Content -Path $jobFile

# Record start time for overall duration calculation
$startTime = Get-Date
Write-Host "Starting at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan

# ==============================
# ANALYZE SOURCE DIRECTORY
# ==============================
# This section identifies what needs to be copied

# Get all files in the source root directory (not in subfolders)
$files = Get-ChildItem -Path $src -File -Force -Attributes !System

# Process the exclude directories list if provided
$excludeDir = foreach ($dir in $excludeDir) {
    $dir.Split([Environment]::NewLine) |
        Where-Object { ($_) -and -not ($_.StartsWith('#')) } | # Skip empty lines and comments
        ForEach-Object { $_.Trim() }  # Remove whitespace
}

# Get all subdirectories in the source, excluding system folders
$folders = Get-ChildItem -Path $src -Directory -Attributes !System -Force

# Apply directory exclusion filter if directories to exclude were specified
if ($excludeDir) {
    $folders = $folders | Where-Object { -not $excludeDir.Contains($_.Name) }
}

# Sort folders by modification time if requested
if ($newestFirst) {
    $folders = $folders | Sort-Object LastWriteTime -Descending
    Write-Host 'Folders sorted by most recently modified first' -ForegroundColor Cyan
}

# ==============================
# DISPLAY INITIAL INFORMATION
# ==============================
# Show a summary of what will be copied
Write-Host -ForegroundColor Cyan 'Starting RoboCopies'
Write-Host -NoNewline 'Logging: '; Write-Host -ForegroundColor Yellow $logDir
Write-Host
Write-Host -NoNewline 'Source: '; Write-Host -ForegroundColor Yellow $src
Write-Host -NoNewline 'Destination: '; Write-Host -ForegroundColor Yellow $dest
Write-Host -NoNewline 'Total folders: '; Write-Host -ForegroundColor Yellow $folders.Count
Write-Host

# ==============================
# COPY FILES IN ROOT DIRECTORY
# ==============================
# Files in the root directory are copied separately from subdirectories

# Copy files from the root directory, if any exist
if ($files) {
    # Inform user that root files are being copied
    Write-Host -NoNewline -ForegroundColor Cyan ('Copying {0} files in root directory... ' -f $files.Count)

    # Create a log file specific to the root copy operation
    $logFile = Join-Path -Path $logDir -ChildPath "root_dir_files-$(Get-Date -f yyyy-MM-dd-mm-ss).log"

    # These options aren't needed or don't apply when copying files in the root directory
    $rootOptionsToExclude = '/NOSD', '/NODD', '/XD', '/E', '/DCOPY', '/NDL'

    # Build a custom options list for the root copy, filtering out excluded options
    $rootOptions = $options | Where-Object Active -EQ $true | ForEach-Object {
        $skipOption = $false
        foreach ($exclusion in $rootOptionsToExclude) {
            if ($_.Option.StartsWith($exclusion)) {
                $skipOption = $true
                break
            }
        }
        if (-not $skipOption) { $_.Option.Split(' ') }
    }

    # Create the final RoboCopy command for root files
    $rootCmdLine = @($src, $dest) + $rootOptions

    # Execute RoboCopy for root files and log the output
    $rootStart = Get-Date
    robocopy $rootCmdLine > $logFile

    # Display completion message and statistics
    Write-Host -NoNewline -ForegroundColor Green 'Completed '
    $jobStats = Get-Content $logFile -Raw | Select-RoboSummary | Where-Object Type -EQ Files
    Write-JobStats $jobStats $rootStart
}

# ==============================
# SETUP FOR PARALLEL PROCESSING
# ==============================
# This section prepares the script for running multiple RoboCopy processes in parallel

# Define the script block that will execute in each background job
# This is the code that will run for each folder copy operation
$ScriptBlock = {
    param($src, $dest, $jobFile, $logFile, $num, $functions)

    # Import the functions needed for log parsing and statistics display
    # (This is necessary because functions don't automatically transfer to background jobs)
    . ([scriptblock]::Create($functions))

    # Record start time for this job
    $jobStart = Get-Date

    # Execute RoboCopy using the job file for options
    # The job file contains all the RoboCopy parameters
    robocopy $src $dest /job:$jobFile > $logFile

    # Display completion information and statistics
    Write-Host -NoNewline -ForegroundColor Green "$num. "
    Write-Host -NoNewline -ForegroundColor Green 'Completed: '
    Write-Host -NoNewline -ForegroundColor White "$(Split-Path -Path $src -Leaf) "

    # Parse and display statistics from the log file
    $jobStats = Get-Content $logFile -Raw | Select-RoboSummary | Where-Object Type -EQ Files
    Write-JobStats $jobStats $jobStart
}

# Create a string containing function definitions to pass to background jobs
# (This allows the background jobs to use the same functions defined in this script)
$functions = Get-FunctionScriptBlock -Function 'Write-JobStats', 'Select-RoboSummary'

# ==============================
# COPY FOLDERS IN PARALLEL
# ==============================
# This is the main part of the script that processes multiple folders simultaneously

# Initialize job counter and define invalid characters for log filenames
$num = 0
$invalidLogNameChars = '[\[\]~"#%&*:<>?/\\{|}]+'

# Process each subfolder using background jobs
$folders | ForEach-Object {
    # If we've reached the maximum number of parallel jobs, wait for one to complete
    # This prevents overloading the system with too many simultaneous operations
    do {
        Start-Sleep -Milliseconds 500
        $j = Get-Job -State 'Running'
    } while ($j.count -ge $maxJobs)

    # Process any jobs that have completed
    # Remove completed jobs to free up system resources
    Get-Job -State 'Completed' | Receive-Job | Out-Null
    Remove-Job -State 'Completed' | Out-Null

    # Increment the job counter
    $num++

    # Define the destination path for this folder
    $destinationFolder = Join-Path -Path $dest -ChildPath $_.Name

    # Create a clean log filename for this folder
    # Remove invalid characters that could cause problems in filenames
    $logFile = Join-Path -Path $logDir -ChildPath "$($_.Name -replace $invalidLogNameChars,'')-$(Get-Date -f yyyy-MM-dd-mm-ss).log"

    # Define a job name for tracking
    $jobName = "$num. $($_.Name)"

    # Inform the user that this folder's job is starting
    Write-Host -NoNewline -ForegroundColor Cyan "$num. Started: "
    Write-Host -ForegroundColor White $_.Name

    # Start a background job for this folder
    # This runs the RoboCopy process in a separate PowerShell instance
    Start-Job -Name $jobName -ScriptBlock $ScriptBlock -ArgumentList $_.FullName, $destinationFolder, $jobFile, $logFile, $num, $functions | Out-Null
}

# ==============================
# WAIT FOR ALL JOBS TO FINISH
# ==============================
# This section monitors the progress of all jobs until completion

# Display status information
Write-Host ''
Write-Host -ForegroundColor Green 'All jobs started.'
Write-Host ''

# For large operations, show current job status
if ($folders.Count -gt 20) {
    Show-RunningJob
}

# Wait for all jobs to finish, showing status updates for long-running operations
$beginCheck = Get-Date
$checkHours = 1

while (Get-Job -State 'Running') {
    Start-Sleep -Seconds 1

    # Display running job status every hour
    # This helps track progress for long-running operations
    if ((Get-Date) -ge $beginCheck.AddHours($checkHours)) {
        $checkHours++
        Show-RunningJob
    }

    # Process completed jobs
    # Remove completed jobs to free up system resources
    Get-Job -State 'Completed' | Receive-Job | Out-Null
    Remove-Job -State 'Completed' | Out-Null
}

# ==============================
# DISPLAY COMPLETION SUMMARY
# ==============================
# This section shows the final results and timing information

# Display completion message and total duration
Write-Host
Write-Host ('All copies completed to {0}' -f $dest) -ForegroundColor Green

# Calculate and display the total operation duration
$duration = New-TimeSpan -Start $startTime -End (Get-Date)
Write-Host -ForegroundColor Yellow ('Duration: {0:dd\d\.hh\:mm\:ss}' -f $duration)

# Show completion time
Write-Host "Finished at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
#endregion FINAL SUMMARY
