param([string]$ComputerName = $env:COMPUTERNAME)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

if (-not ('DwmApi' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DwmApi {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    public static void SetDarkMode(IntPtr handle) { int value = 1; DwmSetWindowAttribute(handle, 20, ref value, 4); }
}
"@
}

$script:TargetPC = $ComputerName
$script:IsRemote = $ComputerName -ne $env:COMPUTERNAME
$script:UsersPath = if ($script:IsRemote) { "\\$ComputerName\C$\Users" } else { "C:\Users" }
$script:profileData = [System.Collections.ArrayList]::new()
$script:SystemAccounts = @('Administrator','Default','Default User','Public','All Users','defaultuser0','SYSTEM','LocalService','NetworkService','systemprofile','LocalService','NetworkService')

function Get-FolderSizeMB($path) {
    if (!(Test-Path $path -ErrorAction SilentlyContinue)) { return 0 }
    try {
        $size = 0
        $dirInfo = New-Object System.IO.DirectoryInfo($path)
        foreach ($file in $dirInfo.GetFiles("*", [System.IO.SearchOption]::AllDirectories)) {
            $size += $file.Length
        }
        return [math]::Round($size / 1MB, 2)
    } catch { return 0 }
}

function Get-ADUserDisplayName($samAccountName) {
    try {
        $user = Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -Properties DisplayName -ErrorAction Stop
        if ($user) { return $user.DisplayName }
    } catch {}
    return ""
}

$script:DomainName = try { (Get-WmiObject Win32_ComputerSystem).Domain -replace '\..*$', '' } catch { "" }

function Is-OrphanedProfile($folderName) {
    if ($folderName -match '\.temp\d*$') { return $true }
    if ($folderName -match '\.bak\d*$') { return $true }
    if ($folderName -match '\.old\d*$') { return $true }
    if ($script:DomainName -and $folderName -match ("\." + [regex]::Escape($script:DomainName) + '\d*$')) { return $true }
    if ($folderName -match '\.\d{6}$') { return $true }
    return $false
}

function Is-LocalAccount($folderName) {
    $baseName = $folderName
    if (Is-OrphanedProfile $folderName) {
        $baseName = $folderName -replace '\.temp\d*$', '' -replace '\.bak\d*$', '' -replace '\.old\d*$', '' -replace '\.\d{6}$', ''
        if ($script:DomainName) {
            $baseName = $baseName -replace ("\." + [regex]::Escape($script:DomainName) + '\d*$'), ''
        }
    }
    if ($baseName -in $script:SystemAccounts) { return $true }
    try {
        $user = Get-ADUser -Filter "SamAccountName -eq '$baseName'" -ErrorAction Stop
        return $false
    } catch {
        return $true
    }
}

function Get-ProfileLastLogon($sid, $folderPath) {
    try {
        if ($script:IsRemote) {
            $ntuser = Join-Path $folderPath "NTUSER.DAT"
            if (Test-Path $ntuser -ErrorAction SilentlyContinue) {
                return (Get-Item $ntuser -Force -ErrorAction SilentlyContinue).LastWriteTime
            }
        } else {
            $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
            if (Test-Path $regPath) {
                $localTime = (Get-ItemProperty $regPath -ErrorAction SilentlyContinue).LocalProfileLoadTimeLow
                $highTime = (Get-ItemProperty $regPath -ErrorAction SilentlyContinue).LocalProfileLoadTimeHigh
                if ($localTime -and $highTime) {
                    $fileTime = ([long]$highTime -shl 32) -bor [long]$localTime
                    if ($fileTime -gt 0) { return [DateTime]::FromFileTime($fileTime) }
                }
            }
            $ntuser = Join-Path $folderPath "NTUSER.DAT"
            if (Test-Path $ntuser -ErrorAction SilentlyContinue) {
                return (Get-Item $ntuser -Force -ErrorAction SilentlyContinue).LastWriteTime
            }
        }
    } catch {}
    return $null
}

function Get-RemoteRegistry {
    $profilePaths = @{}
    $regSIDs = @{}
    if ($script:IsRemote) {
        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $script:TargetPC)
            $profileList = $reg.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList")
            foreach ($sid in $profileList.GetSubKeyNames()) {
                $subKey = $profileList.OpenSubKey($sid)
                $imgPath = $subKey.GetValue("ProfileImagePath")
                if ($imgPath) {
                    $uncPath = $imgPath -replace '^C:', "\\$script:TargetPC\C$"
                    $profilePaths[$uncPath] = $sid
                    $regSIDs[$sid] = $uncPath
                }
            }
            $reg.Close()
        } catch {}
    } else {
        $regProfiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" -ErrorAction SilentlyContinue
        foreach ($p in $regProfiles) {
            if ($p.ProfileImagePath) {
                $profilePaths[$p.ProfileImagePath] = $p.PSChildName
                $regSIDs[$p.PSChildName] = $p.ProfileImagePath
            }
        }
    }
    return @{ Paths = $profilePaths; SIDs = $regSIDs }
}

function Convert-ToLocalPath($uncPath) {
    if ($script:IsRemote) {
        return $uncPath.Replace("\\$script:TargetPC\C$", "C:")
    }
    return $uncPath
}

function Get-BasicProfileInfo {
    $profiles = [System.Collections.ArrayList]::new()
    $regData = Get-RemoteRegistry
    $profilePaths = $regData.Paths
    $regSIDs = $regData.SIDs
    $excludeFolders = @('Public','Default','Default User','All Users')
    try {
        $userFolders = [System.IO.Directory]::GetDirectories($script:UsersPath) | ForEach-Object { [System.IO.DirectoryInfo]::new($_) } | Where-Object { $_.Name -notin $excludeFolders }
    } catch { return $profiles }
    foreach ($folder in $userFolders) {
        $fullPath = $folder.FullName
        $folderName = $folder.Name
        $isOrphaned = Is-OrphanedProfile $folderName
        $isLocal = Is-LocalAccount $folderName
        $hasRegKey = $profilePaths.ContainsKey($fullPath)
        $sid = ""
        if ($hasRegKey) { $sid = $profilePaths[$fullPath] }
        $lastLogon = Get-ProfileLastLogon $sid $fullPath
        $localPath = Convert-ToLocalPath $fullPath
        $isSystem = $folderName -in $script:SystemAccounts
        $null = $profiles.Add([PSCustomObject]@{
            FolderName = $folderName
            DisplayName = ""
            DesktopMB = 0
            DocumentsMB = 0
            DownloadsMB = 0
            TotalGB = 0
            Path = $fullPath
            SID = $sid
            HasRegKey = $hasRegKey
            IsOrphaned = $isOrphaned
            IsLocal = $isLocal
            IsSystem = $isSystem
            HasFiles = $false
            LocalPath = $localPath
            LastLogon = $lastLogon
            SizesLoaded = $false
            CanDelete = (-not $isSystem)
        })
    }
    foreach ($sid in $regSIDs.Keys) {
        $path = $regSIDs[$sid]
        if (!(Test-Path $path -ErrorAction SilentlyContinue)) {
            $localPath = Convert-ToLocalPath $path
            $folderName = Split-Path $path -Leaf
            $isSystem = $folderName -in $script:SystemAccounts
            $null = $profiles.Add([PSCustomObject]@{
                FolderName = "[ORPHAN KEY] $folderName"
                DisplayName = ""
                DesktopMB = 0
                DocumentsMB = 0
                DownloadsMB = 0
                TotalGB = 0
                Path = $path
                SID = $sid
                HasRegKey = $true
                IsOrphaned = $true
                IsLocal = $false
                IsSystem = $isSystem
                HasFiles = $false
                LocalPath = $localPath
                LastLogon = $null
                SizesLoaded = $true
                CanDelete = (-not $isSystem)
            })
        }
    }
    return $profiles
}

function Remove-UserProfile($profile) {
    $errors = @()
    if ($profile.SID) {
        try {
            if ($script:IsRemote) {
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $script:TargetPC)
                $reg.DeleteSubKeyTree("SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($profile.SID)", $false)
                $reg.Close()
            } else {
                Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($profile.SID)" -Recurse -Force -ErrorAction Stop
            }
        } catch { $errors += "Registro: $_" }
    }
    if (Test-Path $profile.Path -ErrorAction SilentlyContinue) {
        try {
            if ($script:IsRemote) {
                $scriptBlock = { param($p) cmd /c "takeown /F `"$p`" /R /A /D Y >nul 2>&1 & icacls `"$p`" /grant Administrators:F /T /C >nul 2>&1 & rd /s /q `"$p`"" }
                Invoke-Command -ComputerName $script:TargetPC -ScriptBlock $scriptBlock -ArgumentList $profile.LocalPath -ErrorAction Stop
            } else {
                $null = takeown /F $profile.Path /R /A /D Y 2>&1
                $null = icacls $profile.Path /grant Administrators:F /T /C 2>&1
                Remove-Item $profile.Path -Recurse -Force -ErrorAction Stop
            }
        } catch { $errors += "Cartella: $_" }
    }
    return $errors
}

function Remove-ProfilesParallel($profiles) {
    $results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
    $targetPC = $script:TargetPC
    $isRemote = $script:IsRemote
    $profiles | ForEach-Object -ThrottleLimit 5 -Parallel {
        $p = $_
        $errors = @()
        $targetPC = $using:targetPC
        $isRemote = $using:isRemote
        if ($p.SID) {
            try {
                if ($isRemote) {
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $targetPC)
                    $reg.DeleteSubKeyTree("SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($p.SID)", $false)
                    $reg.Close()
                } else {
                    Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($p.SID)" -Recurse -Force -ErrorAction Stop
                }
            } catch { $errors += "Registro: $_" }
        }
        if (Test-Path $p.Path -ErrorAction SilentlyContinue) {
            try {
                if ($isRemote) {
                    $scriptBlock = { param($path) cmd /c "takeown /F `"$path`" /R /A /D Y >nul 2>&1 & icacls `"$path`" /grant Administrators:F /T /C >nul 2>&1 & rd /s /q `"$path`"" }
                    Invoke-Command -ComputerName $targetPC -ScriptBlock $scriptBlock -ArgumentList $p.LocalPath -ErrorAction Stop
                } else {
                    $null = takeown /F $p.Path /R /A /D Y 2>&1
                    $null = icacls $p.Path /grant Administrators:F /T /C 2>&1
                    Remove-Item $p.Path -Recurse -Force -ErrorAction Stop
                }
            } catch { $errors += "Cartella: $_" }
        }
        ($using:results).Add([PSCustomObject]@{ Name = $p.FolderName; Errors = $errors })
    }
    return $results
}

function New-Win11Button($text, $x, $y, $w, $h, $accent, $danger) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Location = New-Object System.Drawing.Point($x,$y)
    $btn.Size = New-Object System.Drawing.Size($w,$h)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor = "Hand"
    $btn.Font = New-Object System.Drawing.Font("Segoe UI",10)
    if ($danger) {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(196,43,28)
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.Tag = "danger"
    } elseif ($accent) {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(0,103,192)
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.Tag = "accent"
    } else {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(55,55,55)
        $btn.ForeColor = [System.Drawing.Color]::White
    }
    $btn.Add_MouseEnter({
        $r = [math]::Min($this.BackColor.R + 20, 255)
        $g = [math]::Min($this.BackColor.G + 20, 255)
        $b = [math]::Min($this.BackColor.B + 20, 255)
        $this.BackColor = [System.Drawing.Color]::FromArgb($r,$g,$b)
    })
    $btn.Add_MouseLeave({
        if ($this.Tag -eq "danger") { $this.BackColor = [System.Drawing.Color]::FromArgb(196,43,28) }
        elseif ($this.Tag -eq "accent") { $this.BackColor = [System.Drawing.Color]::FromArgb(0,103,192) }
        else { $this.BackColor = [System.Drawing.Color]::FromArgb(55,55,55) }
    })
    return $btn
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "User Manager"
$form.Size = New-Object System.Drawing.Size(1400,800)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(32,32,32)
$form.ForeColor = [System.Drawing.Color]::White
$form.FormBorderStyle = "Sizable"
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)
$form.Add_Shown({ [DwmApi]::SetDarkMode($form.Handle) })

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = "Top"
$pnlHeader.Height = 90
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(39,39,39)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "USER MANAGER - ZAICK.NET"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI",20)
$lblTitle.Location = New-Object System.Drawing.Point(25,12)
$lblTitle.AutoSize = $true
$lblTitle.ForeColor = [System.Drawing.Color]::White

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Selected Computer: $script:TargetPC"
$lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI",10)
$lblSubtitle.Location = New-Object System.Drawing.Point(27,50)
$lblSubtitle.AutoSize = $true
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)

$lblRemote = New-Object System.Windows.Forms.Label
$lblRemote.Text = "Connect to:"
$lblRemote.Location = New-Object System.Drawing.Point(920,20)
$lblRemote.Size = New-Object System.Drawing.Size(80,25)
$lblRemote.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)

$txtPC = New-Object System.Windows.Forms.TextBox
$txtPC.Text = $script:TargetPC
$txtPC.Location = New-Object System.Drawing.Point(1000,17)
$txtPC.Size = New-Object System.Drawing.Size(200,30)
$txtPC.BackColor = [System.Drawing.Color]::FromArgb(55,55,55)
$txtPC.ForeColor = [System.Drawing.Color]::White
$txtPC.BorderStyle = "FixedSingle"
$txtPC.Font = New-Object System.Drawing.Font("Segoe UI",10)

$btnConnect = New-Win11Button "Connetti" 1210 15 120 32 $true $false
$btnConnect.Add_Click({
    $script:TargetPC = $txtPC.Text.Trim()
    $script:IsRemote = $script:TargetPC -ne $env:COMPUTERNAME
    if ($script:IsRemote) {
        $script:UsersPath = "\\$script:TargetPC\C$\Users"
    } else {
        $script:UsersPath = "C:\Users"
    }
    $form.Text = "Gestione Profili Utente - $script:TargetPC"
    $lblSubtitle.Text = "Computer: $script:TargetPC"
    Refresh-Grid
})

$pnlHeader.Controls.AddRange(@($lblTitle,$lblSubtitle,$lblRemote,$txtPC,$btnConnect))

$pnlLegend = New-Object System.Windows.Forms.Panel
$pnlLegend.Location = New-Object System.Drawing.Point(20,100)
$pnlLegend.Size = New-Object System.Drawing.Size(1340,45)
$pnlLegend.BackColor = [System.Drawing.Color]::FromArgb(44,44,44)

$pnlDarkRed = New-Object System.Windows.Forms.Panel
$pnlDarkRed.Location = New-Object System.Drawing.Point(20,15)
$pnlDarkRed.Size = New-Object System.Drawing.Size(14,14)
$pnlDarkRed.BackColor = [System.Drawing.Color]::FromArgb(80,30,30)

$lblLegDarkRedTxt = New-Object System.Windows.Forms.Label
$lblLegDarkRedTxt.Text = "Local/System"
$lblLegDarkRedTxt.Location = New-Object System.Drawing.Point(40,13)
$lblLegDarkRedTxt.AutoSize = $true
$lblLegDarkRedTxt.ForeColor = [System.Drawing.Color]::FromArgb(200,200,200)

$pnlRed = New-Object System.Windows.Forms.Panel
$pnlRed.Location = New-Object System.Drawing.Point(160,15)
$pnlRed.Size = New-Object System.Drawing.Size(14,14)
$pnlRed.BackColor = [System.Drawing.Color]::FromArgb(232,72,60)

$lblLegRedTxt = New-Object System.Windows.Forms.Label
$lblLegRedTxt.Text = "Orphans/Duplicates"
$lblLegRedTxt.Location = New-Object System.Drawing.Point(180,13)
$lblLegRedTxt.AutoSize = $true
$lblLegRedTxt.ForeColor = [System.Drawing.Color]::FromArgb(200,200,200)

$pnlOrange = New-Object System.Windows.Forms.Panel
$pnlOrange.Location = New-Object System.Drawing.Point(310,15)
$pnlOrange.Size = New-Object System.Drawing.Size(14,14)
$pnlOrange.BackColor = [System.Drawing.Color]::FromArgb(255,185,0)

$lblLegOrangeTxt = New-Object System.Windows.Forms.Label
$lblLegOrangeTxt.Text = "Non-empty"
$lblLegOrangeTxt.Location = New-Object System.Drawing.Point(330,13)
$lblLegOrangeTxt.AutoSize = $true
$lblLegOrangeTxt.ForeColor = [System.Drawing.Color]::FromArgb(200,200,200)

$pnlGreen = New-Object System.Windows.Forms.Panel
$pnlGreen.Location = New-Object System.Drawing.Point(410,15)
$pnlGreen.Size = New-Object System.Drawing.Size(14,14)
$pnlGreen.BackColor = [System.Drawing.Color]::FromArgb(108,203,95)

$lblLegGreenTxt = New-Object System.Windows.Forms.Label
$lblLegGreenTxt.Text = "Empty"
$lblLegGreenTxt.Location = New-Object System.Drawing.Point(430,13)
$lblLegGreenTxt.AutoSize = $true
$lblLegGreenTxt.ForeColor = [System.Drawing.Color]::FromArgb(200,200,200)

$pnlLegend.Controls.AddRange(@($pnlDarkRed,$lblLegDarkRedTxt,$pnlRed,$lblLegRedTxt,$pnlOrange,$lblLegOrangeTxt,$pnlGreen,$lblLegGreenTxt))

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(20,155)
$dgv.Size = New-Object System.Drawing.Size(1340,520)
$dgv.BackgroundColor = [System.Drawing.Color]::FromArgb(44,44,44)
$dgv.GridColor = [System.Drawing.Color]::FromArgb(60,60,60)
$dgv.BorderStyle = "None"
$dgv.CellBorderStyle = "SingleHorizontal"
$dgv.ColumnHeadersBorderStyle = "None"
$dgv.RowHeadersVisible = $false
$dgv.EnableHeadersVisualStyles = $false
$dgv.AllowUserToAddRows = $false
$dgv.AllowUserToDeleteRows = $false
$dgv.AllowUserToResizeRows = $false
$dgv.SelectionMode = "FullRowSelect"
$dgv.AutoSizeColumnsMode = "Fill"
$dgv.RowTemplate.Height = 36
$dgv.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(44,44,44)
$dgv.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0,103,192)
$dgv.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$dgv.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI",9)
$dgv.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8,0,8,0)
$dgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50,50,50)
$dgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(200,200,200)
$dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$dgv.ColumnHeadersHeight = 36

$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.Name = "Sel"
$colCheck.HeaderText = ""
$colCheck.Width = 35
$dgv.Columns.Add($colCheck)
$dgv.Columns.Add("Cartella","User Folder")
$dgv.Columns.Add("NomeCognome","Name and Surname")
$dgv.Columns.Add("UltimoAccesso","Last Login")
$dgv.Columns.Add("Desktop","Desktop")
$dgv.Columns.Add("Documenti","Documents")
$dgv.Columns.Add("Download","Download")
$dgv.Columns.Add("Totale","Total")
$dgv.Columns.Add("Registro","Reg")
$dgv.Columns["UltimoAccesso"].DefaultCellStyle.Alignment = "MiddleCenter"
$dgv.Columns["Desktop"].DefaultCellStyle.Alignment = "MiddleRight"
$dgv.Columns["Documenti"].DefaultCellStyle.Alignment = "MiddleRight"
$dgv.Columns["Download"].DefaultCellStyle.Alignment = "MiddleRight"
$dgv.Columns["Totale"].DefaultCellStyle.Alignment = "MiddleRight"
$dgv.Columns["Registro"].DefaultCellStyle.Alignment = "MiddleCenter"
$dgv.Columns["Registro"].Width = 45

$dgv.Add_CellContentClick({
    param($sender, $e)
    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
        $p = $script:profileData[$e.RowIndex]
        if (-not $p.CanDelete) {
            $dgv.Rows[$e.RowIndex].Cells[0].Value = $false
        }
    }
})

function Format-Size($mb) {
    if ($mb -eq 0) { return "0 KB" }
    $kb = $mb * 1024
    if ($kb -lt 1024) { return "{0:N0} KB" -f $kb }
    elseif ($mb -lt 1024) { return "{0:N1} MB" -f $mb }
    else { return "{0:N2} GB" -f ($mb/1024) }
}

function Format-Date($dt) {
    if ($dt) { return $dt.ToString("dd/MM/yyyy HH:mm") }
    else { return "-" }
}

function Update-RowColor($row, $idx) {
    $p = $script:profileData[$idx]
    if ($p.IsLocal -or $p.IsSystem) {
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50,25,25)
        $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(70,35,35)
        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(140,140,140)
    } elseif ($p.IsOrphaned -or !$p.HasRegKey) {
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(70,35,35)
        $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(140,50,50)
        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    } elseif ($p.HasFiles) {
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(70,55,30)
        $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(140,110,40)
        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    } else {
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(35,60,35)
        $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(50,120,50)
        $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    }
}

function Load-ProfileDetails($idx) {
    $p = $script:profileData[$idx]
    if ($p.SizesLoaded) { return }
    $desktop = Join-Path $p.Path "Desktop"
    $documents = Join-Path $p.Path "Documents"
    $downloads = Join-Path $p.Path "Downloads"
    $p.DesktopMB = Get-FolderSizeMB $desktop
    $p.DocumentsMB = Get-FolderSizeMB $documents
    $p.DownloadsMB = Get-FolderSizeMB $downloads
    $p.TotalGB = [math]::Round(($p.DesktopMB + $p.DocumentsMB + $p.DownloadsMB) / 1024, 2)
    $p.HasFiles = ($p.DesktopMB + $p.DocumentsMB + $p.DownloadsMB) -gt 0
    $baseName = $p.FolderName
    if (Is-OrphanedProfile $p.FolderName) {
        $baseName = $p.FolderName -replace '\.temp\d*$', '' -replace '\.bak\d*$', '' -replace '\.old\d*$', '' -replace '\.\d{6}$', ''
        if ($script:DomainName) {
            $baseName = $baseName -replace ("\." + [regex]::Escape($script:DomainName) + '\d*$'), ''
        }
    }
    $p.DisplayName = Get-ADUserDisplayName $baseName
    $p.SizesLoaded = $true
    $row = $dgv.Rows[$idx]
    $row.Cells["NomeCognome"].Value = $p.DisplayName
    $row.Cells["Desktop"].Value = Format-Size $p.DesktopMB
    $row.Cells["Documenti"].Value = Format-Size $p.DocumentsMB
    $row.Cells["Download"].Value = Format-Size $p.DownloadsMB
    $row.Cells["Totale"].Value = Format-Size ($p.DesktopMB + $p.DocumentsMB + $p.DownloadsMB)
    Update-RowColor $row $idx
}

$script:loadTimer = New-Object System.Windows.Forms.Timer
$script:loadTimer.Interval = 100
$script:currentLoadIndex = 0
$script:bgRunspace = $null
$script:bgPowerShell = $null
$script:bgResult = $null

$script:loadTimer.Add_Tick({
    if ($script:bgResult -eq $null) { return }
    if ($script:bgResult.IsCompleted) {
        $script:loadTimer.Stop()
        try {
            $results = $script:bgPowerShell.EndInvoke($script:bgResult)
            if ($results) {
                foreach ($r in $results) {
                    $idx = $r.Index
                    if ($idx -lt $script:profileData.Count) {
                        $p = $script:profileData[$idx]
                        $p.DesktopMB = $r.DesktopMB
                        $p.DocumentsMB = $r.DocumentsMB
                        $p.DownloadsMB = $r.DownloadsMB
                        $p.HasFiles = ($r.DesktopMB + $r.DocumentsMB + $r.DownloadsMB) -gt 0
                        $p.DisplayName = $r.DisplayName
                        $p.SizesLoaded = $true
                        $row = $dgv.Rows[$idx]
                        $row.Cells["NomeCognome"].Value = $p.DisplayName
                        $row.Cells["Desktop"].Value = Format-Size $p.DesktopMB
                        $row.Cells["Documenti"].Value = Format-Size $p.DocumentsMB
                        $row.Cells["Download"].Value = Format-Size $p.DownloadsMB
                        $row.Cells["Totale"].Value = Format-Size ($p.DesktopMB + $p.DocumentsMB + $p.DownloadsMB)
                        Update-RowColor $row $idx
                    }
                }
            }
        } catch {}
        $script:bgPowerShell.Dispose()
        $script:bgRunspace.Close()
        $lblCount.Text = "$($script:profileData.Count) profili"
    }
})

function Start-BackgroundLoad {
    $script:bgRunspace = [runspacefactory]::CreateRunspace()
    $script:bgRunspace.Open()
    $script:bgPowerShell = [powershell]::Create()
    $script:bgPowerShell.Runspace = $script:bgRunspace
    
    $profiles = $script:profileData | ForEach-Object {
        [PSCustomObject]@{
            Index = $script:profileData.IndexOf($_)
            Path = $_.Path
            FolderName = $_.FolderName
            IsOrphaned = $_.IsOrphaned
        }
    }
    $domainName = $script:DomainName
    
    $null = $script:bgPowerShell.AddScript({
        param($profiles, $domainName)
        
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        
        function Get-Size($path) {
            if (!(Test-Path $path -ErrorAction SilentlyContinue)) { return 0 }
            try {
                $size = 0
                $di = New-Object System.IO.DirectoryInfo($path)
                foreach ($f in $di.GetFiles("*", [System.IO.SearchOption]::AllDirectories)) {
                    $size += $f.Length
                }
                return [math]::Round($size / 1MB, 2)
            } catch { return 0 }
        }
        
        function Get-DisplayName($samName) {
            try {
                $u = Get-ADUser -Filter "SamAccountName -eq '$samName'" -Properties DisplayName -ErrorAction Stop
                if ($u) { return $u.DisplayName }
            } catch {}
            return ""
        }
        
        function Is-Orphaned($name, $domain) {
            if ($name -match '\.temp\d*$') { return $true }
            if ($name -match '\.bak\d*$') { return $true }
            if ($name -match '\.old\d*$') { return $true }
            if ($domain -and $name -match ("\." + [regex]::Escape($domain) + '\d*$')) { return $true }
            if ($name -match '\.\d{6}$') { return $true }
            return $false
        }
        
        function Get-BaseName($name, $domain) {
            $base = $name
            if (Is-Orphaned $name $domain) {
                $base = $name -replace '\.temp\d*$', '' -replace '\.bak\d*$', '' -replace '\.old\d*$', '' -replace '\.\d{6}$', ''
                if ($domain) { $base = $base -replace ("\." + [regex]::Escape($domain) + '\d*$'), '' }
            }
            return $base
        }
        
        $results = @()
        foreach ($p in $profiles) {
            $desktop = Join-Path $p.Path "Desktop"
            $documents = Join-Path $p.Path "Documents"
            $downloads = Join-Path $p.Path "Downloads"
            $baseName = Get-BaseName $p.FolderName $domainName
            $results += [PSCustomObject]@{
                Index = $p.Index
                DesktopMB = Get-Size $desktop
                DocumentsMB = Get-Size $documents
                DownloadsMB = Get-Size $downloads
                DisplayName = Get-DisplayName $baseName
            }
        }
        return $results
    }).AddArgument($profiles).AddArgument($domainName)
    
    $script:bgResult = $script:bgPowerShell.BeginInvoke()
    $script:loadTimer.Start()
}

function Refresh-Grid {
    $script:loadTimer.Stop()
    if ($script:bgPowerShell) { try { $script:bgPowerShell.Stop(); $script:bgPowerShell.Dispose() } catch {} }
    if ($script:bgRunspace) { try { $script:bgRunspace.Close() } catch {} }
    $dgv.Rows.Clear()
    $script:profileData.Clear()
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblCount.Text = "Scansione..."
    $form.Refresh()
    try {
        $script:profileData = Get-BasicProfileInfo
        foreach ($p in $script:profileData) {
            $regStatus = "No"
            if ($p.HasRegKey) { $regStatus = "Yes" }
            $idx = $dgv.Rows.Add($false, $p.FolderName, "...", (Format-Date $p.LastLogon), "...", "...", "...", "...", $regStatus)
            $row = $dgv.Rows[$idx]
            $row.Cells[0].ReadOnly = (-not $p.CanDelete)
            if ($p.IsLocal -or $p.IsSystem) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50,25,25)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(70,35,35)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(140,140,140)
            } elseif ($p.IsOrphaned -or !$p.HasRegKey) {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(70,35,35)
                $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(140,50,50)
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(44,44,44)
            }
        }
        $lblCount.Text = "$($script:profileData.Count) profiles - loading..."
        Start-BackgroundLoad
    } catch {
        $lblCount.Text = "Errore"
        [System.Windows.Forms.MessageBox]::Show("Errore: $_","Errore","OK","Error")
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

$pnlBottom = New-Object System.Windows.Forms.Panel
$pnlBottom.Dock = "Bottom"
$pnlBottom.Height = 70
$pnlBottom.BackColor = [System.Drawing.Color]::FromArgb(39,39,39)

$btnRefresh = New-Win11Button "Update" 20 15 100 40 $true $false
$btnRefresh.Add_Click({ Refresh-Grid })

$btnSelectAll = New-Win11Button "All" 135 15 80 40 $false $false
$btnSelectAll.Add_Click({
    for ($i = 0; $i -lt $dgv.Rows.Count; $i++) {
        if ($script:profileData[$i].CanDelete) {
            $dgv.Rows[$i].Cells[0].Value = $true
        }
    }
})

$btnDeselectAll = New-Win11Button "Nobody" 230 15 90 40 $false $false
$btnDeselectAll.Add_Click({ foreach ($row in $dgv.Rows) { $row.Cells[0].Value = $false } })

$btnSelectOrphaned = New-Win11Button "Sel. Orphans" 335 15 110 40 $false $false
$btnSelectOrphaned.ForeColor = [System.Drawing.Color]::FromArgb(232,72,60)
$btnSelectOrphaned.Add_Click({
    for ($i = 0; $i -lt $dgv.Rows.Count; $i++) {
        $p = $script:profileData[$i]
        if ($p.CanDelete -and ($p.IsOrphaned -or !$p.HasRegKey)) {
            $dgv.Rows[$i].Cells[0].Value = $true
        }
    }
})

$btnSelectEmpty = New-Win11Button "Sel. Empty" 460 15 100 40 $false $false
$btnSelectEmpty.ForeColor = [System.Drawing.Color]::FromArgb(108,203,95)
$btnSelectEmpty.Add_Click({
    for ($i = 0; $i -lt $dgv.Rows.Count; $i++) {
        $p = $script:profileData[$i]
        if ($p.CanDelete -and $p.SizesLoaded -and !$p.HasFiles -and !$p.IsOrphaned -and $p.HasRegKey -and !$p.IsLocal) {
            $dgv.Rows[$i].Cells[0].Value = $true
        }
    }
})

$btnSelectOld = New-Win11Button "Sel. >90gg" 575 15 110 40 $false $false
$btnSelectOld.ForeColor = [System.Drawing.Color]::FromArgb(255,185,0)
$btnSelectOld.Add_Click({
    $cutoff = (Get-Date).AddDays(-90)
    for ($i = 0; $i -lt $dgv.Rows.Count; $i++) {
        $p = $script:profileData[$i]
        if ($p.CanDelete -and $p.LastLogon -and $p.LastLogon -lt $cutoff -and !$p.IsOrphaned -and !$p.IsLocal) {
            $dgv.Rows[$i].Cells[0].Value = $true
        }
    }
})

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "0 profiles"
$lblCount.Location = New-Object System.Drawing.Point(950,23)
$lblCount.AutoSize = $true
$lblCount.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
$lblCount.Font = New-Object System.Drawing.Font("Segoe UI",10)

$btnDelete = New-Win11Button "Delete selected" 1170 15 160 40 $false $true
$btnDelete.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$btnDelete.Add_Click({
    $selected = @()
    for ($i = 0; $i -lt $dgv.Rows.Count; $i++) {
        if ($dgv.Rows[$i].Cells[0].Value -eq $true -and $script:profileData[$i].CanDelete) {
            $selected += $script:profileData[$i]
        }
    }
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No profile selected","Warning","OK","Warning")
        return
    }
    $msg = "Delete $($selected.Count) profiles from $script:TargetPC?`n`nFast parallel deletion.`nThis operation is IRREVERSIBLE!"
    $confirm = [System.Windows.Forms.MessageBox]::Show($msg,"Confirm deletion","YesNo","Warning")

    if ($confirm -eq "Yes") {
        $script:loadTimer.Stop()
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Removing profiles"
        $progressForm.Size = New-Object System.Drawing.Size(500,180)
        $progressForm.StartPosition = "CenterParent"
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.BackColor = [System.Drawing.Color]::FromArgb(32,32,32)
        $progressForm.ForeColor = [System.Drawing.Color]::White
        $progressForm.MaximizeBox = $false
        $progressForm.MinimizeBox = $false
        $progressForm.ControlBox = $false
        
        $lblProgress = New-Object System.Windows.Forms.Label
        $lblProgress.Text = "Deletion in progress..."
        $lblProgress.Location = New-Object System.Drawing.Point(20,20)
        $lblProgress.Size = New-Object System.Drawing.Size(450,25)
        $lblProgress.Font = New-Object System.Drawing.Font("Segoe UI",10)
        $lblProgress.ForeColor = [System.Drawing.Color]::White
        
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Location = New-Object System.Drawing.Point(20,55)
        $progressBar.Size = New-Object System.Drawing.Size(445,30)
        $progressBar.Style = "Marquee"
        $progressBar.MarqueeAnimationSpeed = 30
        
        $lblDetail = New-Object System.Windows.Forms.Label
        $lblDetail.Text = "$($selected.Count) profiles processing in parallel..."
        $lblDetail.Location = New-Object System.Drawing.Point(20,95)
        $lblDetail.Size = New-Object System.Drawing.Size(450,25)
        $lblDetail.Font = New-Object System.Drawing.Font("Segoe UI",9)
        $lblDetail.ForeColor = [System.Drawing.Color]::FromArgb(160,160,160)
        
        $progressForm.Controls.AddRange(@($lblProgress,$progressBar,$lblDetail))
        $progressForm.Show()
        $progressForm.Refresh()
        
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, 10)
        $runspacePool.Open()
        $jobs = @()
        $targetPC = $script:TargetPC
        $isRemote = $script:IsRemote
        
        $scriptBlock = {
            param($prof, $targetPC, $isRemote)
            $errors = @()
            if ($prof.SID) {
                try {
                    if ($isRemote) {
                        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $targetPC)
                        $reg.DeleteSubKeyTree("SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($prof.SID)", $false)
                        $reg.Close()
                    } else {
                        Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($prof.SID)" -Recurse -Force -ErrorAction Stop
                    }
                } catch { $errors += "Reg: $_" }
            }
            if (Test-Path $prof.Path -ErrorAction SilentlyContinue) {
                try {
                    cmd /c "rd /s /q `"$($prof.Path)`"" 2>$null
                    if (Test-Path $prof.Path -ErrorAction SilentlyContinue) {
                        cmd /c "takeown /F `"$($prof.Path)`" /R /A /D Y >nul 2>&1"
                        cmd /c "icacls `"$($prof.Path)`" /grant Administrators:F /T /C /Q >nul 2>&1"
                        cmd /c "rd /s /q `"$($prof.Path)`"" 2>$null
                    }
                } catch { $errors += "Dir: $_" }
                if (Test-Path $prof.Path -ErrorAction SilentlyContinue) {
                    $errors += "Folder not deleted"
                }
            }
            return [PSCustomObject]@{ Name = $prof.FolderName; Errors = $errors }
        }
        
        foreach ($p in $selected) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $runspacePool
            $null = $ps.AddScript($scriptBlock).AddArgument($p).AddArgument($targetPC).AddArgument($isRemote)
            $jobs += [PSCustomObject]@{ Pipe = $ps; Result = $ps.BeginInvoke() }
        }
        
        $errors = @()
        $success = 0
        foreach ($job in $jobs) {
            $result = $job.Pipe.EndInvoke($job.Result)
            if ($result.Errors.Count -gt 0) {
                $errors += "$($result.Name): $($result.Errors -join '; ')"
            } else {
                $success++
            }
            $job.Pipe.Dispose()
        }
        $runspacePool.Close()
        $runspacePool.Dispose()
        
        $progressForm.Close()
        $progressForm.Dispose()
        
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        if ($errors.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("Completati: $success`nErrori: $($errors.Count)`n`n$($errors -join "`n")","Risultato","OK","Warning")
        } else {
            [System.Windows.Forms.MessageBox]::Show("$success profiles deleted","Completed","OK","Information")
        }
        Refresh-Grid
    }
})

$lblZaick = New-Object System.Windows.Forms.Label
$lblZaick.Text = "zaick.net"
$lblZaick.Location = New-Object System.Drawing.Point(1270,25)
$lblZaick.AutoSize = $true
$lblZaick.ForeColor = [System.Drawing.Color]::FromArgb(100,100,100)
$lblZaick.Font = New-Object System.Drawing.Font("Segoe UI",9)
$lblZaick.Anchor = "Right"

$pnlBottom.Controls.AddRange(@($btnRefresh,$btnSelectAll,$btnDeselectAll,$btnSelectOrphaned,$btnSelectEmpty,$btnSelectOld,$lblCount,$btnDelete,$lblZaick))

$form.Controls.AddRange(@($pnlHeader,$pnlLegend,$dgv,$pnlBottom))

$form.Add_Shown({ Refresh-Grid })

$form.Add_Resize({
    $pnlLegend.Width = $form.ClientSize.Width - 40
    $dgv.Width = $form.ClientSize.Width - 40
    $dgv.Height = $form.ClientSize.Height - $pnlHeader.Height - $pnlLegend.Height - $pnlBottom.Height - 25
})

$form.Add_FormClosing({
    $script:loadTimer.Stop()
    if ($script:bgPowerShell) { try { $script:bgPowerShell.Stop(); $script:bgPowerShell.Dispose() } catch {} }
    if ($script:bgRunspace) { try { $script:bgRunspace.Close() } catch {} }
})

[void]$form.ShowDialog()
