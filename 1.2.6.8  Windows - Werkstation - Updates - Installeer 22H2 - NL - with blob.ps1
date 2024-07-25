[CmdletBinding()]
param (
    # Niks
)

begin {
    # Set execution policy to allow running external scripts
    try {
        Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Verbose 'Execution policy set to RemoteSigned for CurrentUser.'
    } catch {
        Write-Error "Failed to set execution policy: $($_.Exception.Message)"
        return
    }

    # Get the system language, and change ISO url accordingly
    $SystemLanguage = (Get-WinSystemLocale).Name
    if ($SystemLanguage -like '*nl*') {
        $IsoUrl = 'https://fourtop.blob.core.windows.net/beheerfiles/installatiebestanden/Win11_22H2_Dutch_x64v2.iso'
    } elseif ($SystemLanguage -like '*en*') {
        $IsoUrl = 'https://fourtop.blob.core.windows.net/beheerfiles/installatiebestanden/Win11_22H2_English_x64v2.iso'
    }

    # Check if MoSetup key exists, if not, create it
    $MoSetupPath = 'HKLM:\SYSTEM\Setup\MoSetup'
    if ( -not ( Test-Path -Path $MoSetupPath ) ) {
        try {
            New-Item -Path 'HKLM:\SYSTEM\Setup' -Name 'MoSetup'
            Write-Verbose 'MoSetup registry key created.'
        } catch {
            Write-Error "Failed to create MoSetup registry key: $($_.Exception.Message)"
        }
    }

    # Ensure C:\Temp directory exists
    $TempDirectory = 'C:\Temp'
    if ( -not ( Test-Path -Path $TempDirectory ) ) {
        try {
            New-Item -Path $TempDirectory -ItemType 'Directory'
            Write-Output "Created folder: $TempDirectory"
        } catch {
            throw "Failed to create Temp directory: $($_.Exception.Message)"
        }
    }

    # Add TPM bypass to the registry
    try {
        New-ItemProperty -Path $MoSetupPath -Name 'AllowUpgradesWithUnsupportedTPMOrCPU' -PropertyType DWord -Value 1 -Force -ErrorAction stop
        Write-Verbose 'Added TPM bypass registry key.'
    } catch {
        Write-Error $_.Exception.Message
    }
}

process {
    # Start BITS transfer for the download
    $IsoFilePath = "$($TempDirectory)\windows11-22H2-NL-NL.iso"
    try {
        Start-BitsTransfer -Source $IsoUrl -Destination $IsoFilePath -TransferType Download -ErrorAction Stop
        Write-Verbose "ISO downloaded to $IsoFilePath."
    } catch {
        throw "BITS transfer failed: $($_.Exception.Message)"
    }
    
    # Mount the ISO
    try {
        $DiskImage = Mount-DiskImage -ImagePath $IsoFilePath -PassThru -ErrorAction Stop
        Write-Verbose "Mounted ISO: $($IsoFilePath)."
    } catch {
        throw "Failed to mount ISO: $($_.Exception.Message)"
    }

    # Wait for the ISO to mount and get the volume
    $MountStatus = $false
    $Retries = 0
    $MaxRetries = 10
    while ( -not $MountStatus -and $Retries -lt $MaxRetries ) {
        try {
            $DriveLetter = ( Get-DiskImage -ImagePath $IsoFilePath | Get-Volume ).DriveLetter
            if ( $DriveLetter ) {
                $MountStatus = $true
                Write-Verbose "ISO mounted with drive letter: $($DriveLetter)."
            }
        } catch {
            Start-Sleep -Seconds 5
            $Retries++
            Write-Verbose "Retry $($Retries)/$($MaxRetries) to get the drive letter."
        }
    }

    if ( $MountStatus ) {
        # Parameters for Start-Process
        $SetupParams = @{
            FilePath     = "$($DriveLetter):\setup.exe"
            ArgumentList = '/auto upgrade /compat ignorewarning /dynamicupdate disable /showoobe none /eula accept /noreboot /quiet /BitLocker TryKeepActive'
            Wait         = $true
            NoNewWindow  = $true
        }
        
        # Start the installation and wait until it completes
        try {
            Start-Process @SetupParams
            Write-Verbose 'Windows setup started.'
            
            # Unmount the ISO
            Dismount-DiskImage -ImagePath $IsoFilePath
            Write-Verbose "ISO unmounted: $($IsoFilePath)."

            # Remove the TPM bypass from the registry after setup completes
            Remove-ItemProperty -Path $MoSetupPath -Name 'AllowUpgradesWithUnsupportedTPMOrCPU' -Force
            Write-Verbose 'Removed TPM bypass registry key.'
            
            Write-Output 'Upgrade to Windows 11 22H2 was successful.'
            return
        } catch {
            throw "Setup process failed: $($_.Exception.Message)"
        }
    } else {
        throw 'Failed to mount the ISO.'
    }
}