function Prepare-IwrCommand {
    # create reg key to skip IExplore/Edge's first run wizard so that Invoke-WebRequest works properly
    
    $keyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main' 
    try {
        New-Item -Path $keyPath -Force -ErrorAction Stop | Out-Null
        Set-ItemProperty -Path $keyPath -Name 'DisableFirstRunCustomize' -Value 1
    } catch [System.Security.SecurityException],[System.UnauthorizedAccessException] {
        
        if ($MyInvocation.MyCommand.Definition.Replace("`r",'') -match 'try {\n(?<adminCommands>(.|\n)+)\n\s*} catch') {
            $adminCommands = $Matches.adminCommands.Replace('    ','').Replace('$keyPath',"'$keyPath'")
            Start-Process -FilePath powershell.exe -Verb RunAs -Wait -ArgumentList '-Command',$adminCommands
        }
    }
    
}

function Install-VirtualBox {
    # TODO: check if virtualbox and ext pack are alreadu installed before downlaoding
    # TODO: check if Hyper-V hypervisor is enabled and throw error
    $downloadPageUri = 'https://download.virtualbox.org/virtualbox'
    $global:downloadDestination = New-Item -Path $PSScriptRoot -Name Files -ItemType Directory | Select-Object -ExpandProperty FullName
    try {
        $latestVersion = Invoke-WebRequest -Uri "$downloadPageUri/LATEST.TXT" -Method Get
    } catch [System.NotSupportedException] {
        Prepare-IwrCommand
        $latestVersion = Invoke-WebRequest -Uri "$downloadPageUri/LATEST.TXT" -Method Get
    }

    $version = $latestVersion.Content.Trim()
    $versionFilesPage = Invoke-WebRequest -Uri "$downloadPageUri/$version/"

    # download the installer file to the configured destination
    if ($versionFilesPage.Content -match '<a href="(?<filename>.*Win\.exe)"') {
        $installerUri = "$downloadPageUri/$version/$($Matches['filename'])"
        $installerFile = Join-Path -Path $downloadDestination -ChildPath $Matches['filename']
        Invoke-WebRequest -Uri $installerUri -Method Get -OutFile $installerFile
    }

    # download the extension pack to the configured destination
    if ($versionFilesPage.Content -match '<a href="(?<extPack>.*\.vbox-extpack)"') {
        $extPackUri = "$downloadPageUri/$version/$($Matches['extPack'])"
        $extPackFile = Join-Path -Path $downloadDestination -ChildPath $Matches['extPack']
        Invoke-WebRequest -Uri $extPackUri -Method Get -OutFile $extPackFile
    }

    # silently extract the MSI from the installer EXE
    Start-Process -FilePath $installerFile -Verb runas -Wait -ArgumentList '-extract','-silent','-path',$downloadDestination

    # silently install by passing the extracted msi to msiexec with correct arguments
    $installerMsi = Get-ChildItem -Path $downloadDestination | Where-Object -Property Name -Value 'msi$' -Match | Select-Object -ExpandProperty FullName
    Start-Process -FilePath msiexec.exe -Verb runas -Wait -ArgumentList '/i',$installerMsi,'/qn','VBOX_INSTALLDESKTOPSHORTCUT=0'

    # add virtualbox install directory to PATH
    $Env:Path += ';C:\Program Files\Oracle\VirtualBox'

    # install extension pack
    'y' | vboxmanage extpack install $extPackFile

    # create VirtualBox VMs root folder
    # TODO: make this customizable
    $global:VMroot = New-Item -Path $PSScriptRoot -Name VMs -ItemType Directory | Select-Object -ExpandProperty FullName
}