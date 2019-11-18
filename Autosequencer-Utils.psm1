# Lots of useful info about remoting to VMs here: https://docs.microsoft.com/en-us/virtualization/hyper-v-on-windows/user-guide/powershell-direct

Function Get-SequencerDataFile () {
    # Return the data file for the autosequencer VM
    # If more than one autosequencer VM, take an optional VMName parameter to specify which one
    
    param(
        [Parameter(Mandatory = $false)][string]$VMName
    )

    # Autosequencer VM credentials are stored as secure string in a file in $SequencerDataFOlder
    $SequencerDataFolder = "C:\ProgramData\Microsoft Application Virtualization\AutoSequencer\SequencerMachines"
    
    $SequencerDataFiles = (Get-ChildItem -file $SequencerDataFolder)
    if ($SequencerDataFiles.Count -eq 0) {
        Throw "No Sequencer data files found in $SequencerDataFolder.."
    } elseif ($VMName) {
        $SeqFile = $SequencerDataFiles | where { $_.Name -eq $VMName }
        if (-not $SeqFile) {
            Throw "No sequencer data file matching $VMName in $SequencerDataFolder"
        }
    } elseif ($SequencerDataFiles.Count -eq 1) {
        $SeqFile = $SequencerDataFiles[0]
    } else {
        $SequencerNames = $SequencerDataFiles.name -join ", "
        Throw "Multiple sequencer data files please pick one: $SequencerNames"
    }

    return $SeqFile
}

Function Get-SequencerVMCred () {
    # Return a login credential for the autosequencer VM
    # If more than one autosequencer VM, take an optional VMName parameter to specify which one
    
    param(
        [Parameter(Mandatory = $false)][string]$VMName
    )

    $SeqFile = Get-SequencerDataFile $VMName
    $CredInfo = (Get-Content $SeqFile.FullName) -split "\n"
    $Cred = New-Object System.Management.Automation.PSCredential $CredInfo[0], ( $CredInfo[1] | ConvertTo-SecureString)

    Return $Cred
}

Function New-SequencerPSSession () {
    # Create PSSession to Autosequencer VM
    # If more than one autosequencer VM, take an optional VMName parameter to specify which one
    
    param(
        [Parameter(Mandatory = $false)][string]$VMName
    )

    $SeqFile = Get-SequencerDataFile $VMName
    $Cred = Get-SequencerVMCred $VMName

    if (-not ($VMName)) {
        $VMName = $SeqFile.Name
    }

    $Sess = new-pssession -vmname $VMName -credential $Cred

    return $Sess
}
Export-ModuleMember New-SequencerPSSession

Function Add-SequencerToHosts () {
    # Add Sequencer VM to hosts and fix any network issues that might affect connecting to it
    # If more than one autosequencer VM, take an optional VMName parameter to specify which one
    
    param(
        [Parameter(Mandatory = $false)][string]$VMName
    )

    $Hostsfile = "C:\windows\system32\drivers\etc\hosts"
    $SeqFile = Get-SequencerDataFile $VMName
    $Sess = New-SequencerPSSession $VMName
    
    if (-not ($VMName)) {
        $VMName = $SeqFile.Name
    }

    Write-Host "* Disabling Firewall on VM"
    Write-Host "* Disabling TCP/IPv6 on VM"
    Write-Host "* Setting ExecutionPolicy to remotesigned on VM"
    Invoke-command -session $Sess -scriptblock {
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False
        disable-netadapterbinding -name "Ethernet" -componentid ms_tcpip6 -passthru
        Set-executionPolicy RemoteSigned
    } | select systemname, ifAlias, DisplayName, Enabled > $Null

    $IP = get-vm $VMName | select -expand NetworkAdapters | select -expand ipaddresses
    if (-not ($IP)) {
        Throw ("Unable to get IP address for $VMName")
    }
    Write-Host "Adding entry for VM to your hosts file: $VMName -> $IP"
    $SeqHostsEntry = "$IP`t`t$VMName"
    Copy $hostsfile $hostsfile.backup
    (get-content $hostsfile) -replace "^.+$VMName.*$",  $SeqHostsEntry | set-content  $hostsfile

    Write-Host "* Setting timezone of VM to be the same as your computer to lower-cognitive dissonance quotient"
    $MyTimezone = (Get-TimeZone)
    invoke-command -Session $Sess -scriptblock {
        param($timezone)
        set-timezone -id $timezone.id
    } -argumentlist $MyTimeZone
}
Export-ModuleMember Add-SequencerToHosts

Function Enter-SequencerPSSession () {
    # Enter a PSSession on the Sequencer VM
    # If more than one autosequencer VM, take an optional VMName parameter to specify which one
    
    param(
        [Parameter(Mandatory = $false)][string]$VMName
    )

    $Sess = New-SequencerPSSession $VMName
    Enter-PSSession -Session $Sess
}
Export-ModuleMember Enter-SequencerPSSession

Function Get-SequencerLogin () {
    # Create a terminal session to Autosequencer VM
    # If more than one autosequencer VM, take an optional VMName parameter to specify which one

    param(
        [Parameter(Mandatory = $false)][string]$VMName
    )

    $SeqFile = Get-SequencerDataFile $VMName
    $Cred = Get-SequencerVMCred $VMName

    if (-not ($VMName)) {
        $VMName = $SeqFile.Name
    }

    $Username = $Cred.UserName
    $Password = $Cred.GetNetworkCredential().Password

    [PSCustomObject]@{
        VMName = $VMName
        UserName = $UserName
        Password = $PassWord
    }
}

Function Connect-Sequencer () {

    # Create a terminal session to Autosequencer VM
    # If more than one autosequencer VM, take an optional VMName parameter to specify which one

    param(
        [Parameter(Mandatory = $false)][String]$VMName,
        [Parameter(Mandatory = $false)][Int16]$w,
        [Parameter(Mandatory = $false)][Int16]$h
    )

    $SeqFile = Get-SequencerDataFile $VMName
    $Cred = Get-SequencerVMCred $VMName
    Add-SequencerToHosts $VMName

    if (-not ($VMName)) {
        $VMName = $SeqFile.Name
    }

    $Username = $Cred.UserName
    $Password = $Cred.GetNetworkCredential().Password

    $geom=""
    if (-not ($h) -and ($env:TSHeight)) { $h = $env:TSHeight }
    if (-not ($w) -and ($env:TSWidth)) { $w = $env:TSWidth }
    if ($h) {
        $env:TSHeight = $h
        $geom += " /h:$h "
    }
    if ($w) {
        $env:TSWidth = $w
        $geom += " /w:$w "
    }

    Cmdkey /generic:$VMName /user:$Username /pass:$Password
    $P = Start-process mstsc "/v:$VMName $geom" -PassThru
    if (!$P)
    {
        Throw "Unable to start terminal services session to $VMName"
    }
    return

    # Woulda been nice to optionally open powershell and explorer in the build folder but no luck so far:
    $l = Get-SequencerLogin $VMName
    $s = New-SequencerPSSession $VMName
    $consolesession = (invoke-command -session $p -scriptblock { quser }) -split "\s" -match "rdp-tcp.[0-9]*"
    PsExec.exe \\seq1903 -u $l.username -p $l.password -d  "calc.exe" -i $ConsoleSession
}
Export-ModuleMember Connect-Sequencer

Function Reset-Sequencer () {
    # Reset sequencer to checkpoint 'sequencer-base'
    # If more than one autosequencer VM, take an optional VMName parameter to specify which one
    param(
        [Parameter(Mandatory = $false)][string]$VMName
    )

    $Snapshot = 'sequencer-base'

    $SeqFile = Get-SequencerDataFile $VMName

    if (-not ($VMName)) {
        $VMName = $SeqFile.Name
    }

    Restore-VMSnapshot -Name $snapshot -VMName $VMName -Confirm:$false
}
Export-ModuleMember Reset-Sequencer

function Get-AllAppVs () {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][String]$Path
    )

    if (-not ($Path)) {
        $SearchPath = "*.appv"
    } elseif ($Path -like "*.appv") {
        $SearchPath = $Path
    } elseif ($Path[-1] -ne "\") {
        $SearchPath = $Path + "\*.appv"
    } else {
        $SearchPath = $Path + "*.appv"
    }

    get-childitem -Recurse $SearchPath
}

Function Test-AppV () {
    # If a single .appv file exists in this folder or a subfolder, add it (or if it is already added, remove it)
    # If there is more than one .appv you'll have to specify which one.
    # You can specify a fully qualified .appv, or a folder containing a .appv, and you can use wildcards.
    # You can also use command completion

    param(
        [Parameter(Position = 0, Mandatory = $False)]
        [ArgumentCompleter(
            {
                param($AppVPath)
                Get-AllAppVs
            }
        )][string]
        $Path,
        [Parameter(Mandatory = $false)][Switch]$User
    )

    $AppVFile = Get-AllAppvs $Path
    if (($AppVFile).count -eq 0) {
        Write-Warning "No .appv files found in current directory $((get-location).path))"
        return
    } elseif (($AppVFile).count -gt 1) {
        $AppVList = $AppVFile | ForEach { resolve-path -relative $_.fullname }
        $AppVList = "`n" + ($AppVList -join "`n")
        Write-Warning "$(($AppVFile).Count) .appv files found in this directory $((get-location).path)), pick one: $AppVList"
        return
    }

    $pkg = Get-AppvClientPackage -all | Where-Object Path -eq $AppVFile.FullName
    if ($pkg) {
        Write-Warning "Unpublishing and removing $pkg.Name"
        # $pkg | Unpublish-AppVClientPackage
        if ($pkg.IsPublishedGlobally) {
            UnPublish-AppVClientPackage -global -Package $pkg
        } else {
            UnPublish-AppVClientPackage -Package $pkg
        }
        Write-Warning "Removed"
        return
    } else {
        Write-Verbose "Adding App-V package file $(Resolve-path -Relative $AppVFile.FullName) "
        $pkg = Add-AppVClientPackage -path $AppVFile.FullName
        if ($User) {
            $pkg | Publish-AppVClientPackage
        } else {
            $pkg | Publish-AppVClientPackage -Global
        }
    }
}

# WIP
Function Create-APPVPackagePlaceHolder () {
#New-CMApplication
@{
    Name = ""                                       #<String>
    Description = ""                                #<String>
    Publisher = ""                                  #<String>
    SoftwareVersion = ""                            #<String>
    OptionalReference = ""                          #<String>
    ReleaseDate = (get-date)                        #<DateTime>
    AutoInstall = $False                            #<Boolean>
    Owner = ""                                      #<String>
    SupportContact = ""                             #<String>
    LocalizedName = ""                              #<String>
    UserDocumentation = ""                          #<String>
    LinkText = ""                                   #<String>
    LocalizedDescription = ""                       #<String>
    Keyword = ""                                    #<String>
    PrivacyUrl = ""                                 #<String>
    IsFeatured = $False                             #<Boolean>
    IconLocationFile = ""                           #<String>
    DisplaySupersedenceInApplicationCatalog = $True #<Boolean>
}

# Add-CMAppv5XDeploymentType 
@{
    ContentFallback = ""
    FastNetworkDeploymentMode = ""                  # <ContentHandlingMode>
    SlowNetworkDeploymentMode = ""                  # <ContentHandlingMode>
    DeploymentTypeName = ""                         # <String>
    AddRequirement = ""                             # <Rule[]>
    ApplicationName = ""                            # <String>
    RemoveLanguage = ""                             # <String[]>
    RemoveRequirement = ""                          # <Rule[]>
    AddLanguage = "en","en-US","en_GB"              # <Rule[]>
    Comment = ""                                    # <String>
    ContentLocation = ""                            # <String>
}
}
# new-batchappvsequencerpackages -configfile "C:\scratch\packaging\Blender\blender_config.xml" -vmname seq1903 -outputpath "C:\scratch\packaging\Blender\output"
