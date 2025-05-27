Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization # For future charting if needed

#region Color Scheme Definition
$global:colorScheme = @{
    FormBackground        = [System.Drawing.Color]::FromArgb(245, 245, 245) # Light Gray (almost White Smoke)
    PanelBackground       = [System.Drawing.Color]::FromArgb(225, 225, 225) # Slightly adjusted gray for button panel
    OutputBackground      = [System.Drawing.Color]::FromArgb(248, 248, 248) # Very light gray for output, off-white
    OutputForeground      = [System.Drawing.Color]::FromArgb(30, 30, 30)    # Dark Gray for text
    ButtonBackground      = [System.Drawing.Color]::FromArgb(0, 120, 215)   # Modern Blue
    ButtonHoverBackground = [System.Drawing.Color]::FromArgb(0, 100, 185)   # Darker Blue for hover
    ButtonForeground      = [System.Drawing.Color]::White
    LabelForeground       = [System.Drawing.Color]::FromArgb(0, 51, 102)    # Dark Blue for labels/headers
    Title                 = [System.Drawing.Color]::FromArgb(0, 102, 204)   # Blue for titles in output
    Warning               = [System.Drawing.Color]::FromArgb(255, 153, 0)   # Orange/Amber for warnings
    Success               = [System.Drawing.Color]::FromArgb(34, 139, 34)   # ForestGreen for success
    Error                 = [System.Drawing.Color]::FromArgb(205, 0, 0)     # Strong Red for errors
    Info                  = [System.Drawing.Color]::FromArgb(0, 120, 215)   # Blue for general info
    DefaultText           = [System.Drawing.Color]::FromArgb(50, 50, 50)    # Default text color for output
}
#endregion
#region Global Variables and Helper Functions for GUI

$Global:OutputRichTextBox = $null

function Write-GuiOutput {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = $global:colorScheme.DefaultText,
        [bool]$AppendNewLine = $true,
        [bool]$Bold = $false
    )
    if ($Global:OutputRichTextBox) {
        $Global:OutputRichTextBox.SelectionStart = $Global:OutputRichTextBox.TextLength
        $Global:OutputRichTextBox.SelectionLength = 0
        $Global:OutputRichTextBox.SelectionColor = $Color
        $Global:OutputRichTextBox.SelectionFont = New-Object System.Drawing.Font($Global:OutputRichTextBox.Font.FontFamily, $Global:OutputRichTextBox.Font.Size, ($Bold ? [System.Drawing.FontStyle]::Bold : [System.Drawing.FontStyle]::Regular))
        $Global:OutputRichTextBox.AppendText($Text + ($AppendNewLine ? "`n" : ""))
        $Global:OutputRichTextBox.ScrollToCaret()
    } else {
        Write-Host $Text # Fallback if GUI not ready
    }
}

function Clear-GuiOutput {
    if ($Global:OutputRichTextBox) {
        $Global:OutputRichTextBox.Clear()
    }
}

function Show-GuiMessage {
    param(
        [string]$Message,
        [string]$Title = "Thông báo",
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    return [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon)
}

function Confirm-Action-GUI {
    param (
        [string]$Message = "Bạn có chắc chắn muốn tiếp tục?"
    )
    $result = Show-GuiMessage -Message $Message -Title "Xác nhận" -Buttons YesNo -Icon Question
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

#endregion

#region Original Script Functions (Modified for GUI)

# Administrator Check
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Show-GuiMessage -Message "Vui lòng chạy ứng dụng với quyền ADMINISTRATOR để hoạt động đầy đủ!" -Title "Yêu cầu quyền Administrator" -Icon Error
    exit
}

function Get-SystemInfo-GUI {
    Clear-GuiOutput
    try {
        Write-GuiOutput "Đang thu thập thông tin hệ thống, vui lòng đợi..." -Color $global:colorScheme.Info -Bold $true
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor # Keep this at the beginning

        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine("--- THÔNG TIN HỆ THỐNG ---`n")

        # Thông tin Chung
        $sb.AppendLine("**Thông tin Chung:**")
        
        # Get-ComputerInfo replacement for specific properties
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue

        $csName = if ($cs) { $cs.Name } else { $env:COMPUTERNAME }
        $sb.AppendLine("  Tên máy tính: $csName")

        $chassisTypeNumber = (Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ChassisTypes | Where-Object {$_ -ne 0 -and $_ -ne 1 -and $_ -ne 2} | Select-Object -First 1)
        $loaiMay = "Không xác định"
        if ($chassisTypeNumber -in (3,4,5,6,7,13,15,17,24)) { $loaiMay = "PC (Desktop)" }
        elseif ($chassisTypeNumber -in (8,9,10,11,14,30,31,32)) { $loaiMay = "PC (Laptop/Portable)" }
        $sb.AppendLine("  Loại máy: $loaiMay")

        $osName = if ($os) { $os.Caption } else { "Không xác định" }
        $sb.AppendLine("  Hệ điều hành: $osName")
        $osVersion = if ($os) { $os.Version } else { "Không xác định" } # Version includes build number
        $sb.AppendLine("  Phiên bản HĐH: $osVersion")

        # IP and MAC address (existing logic seems fine for PS3+)
        $activeAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object {$_.Status -eq 'Up'} | Select-Object -First 1
        $ipAddress = "Không có kết nối mạng"
        $macAddress = "Không có"
        if ($activeAdapter) {
            $macAddress = $activeAdapter.MacAddress
            $ipConfig = Get-NetIPConfiguration -InterfaceIndex $activeAdapter.ifIndex -ErrorAction SilentlyContinue | Where-Object {$_.IPv4DefaultGateway -ne $null -and $_.NetObjectStatus -eq 'Healthy'} | Select-Object -First 1
            if ($ipConfig.IPv4Address.IPAddress) {
                # Ensure AddressFamily property exists before trying to filter by it, for older systems/adapters
                if ($ipConfig.IPv4Address[0].PSObject.Properties.Name -contains 'AddressFamily') {
                    $ipAddress = ($ipConfig.IPv4Address | Where-Object {$_.AddressFamily -eq 'InterNetwork'} | Select-Object -ExpandProperty IPAddress) -join ", "
                } else { # Fallback if AddressFamily is not present, take the first one
                    $ipAddress = ($ipConfig.IPv4Address | Select-Object -ExpandProperty IPAddress -First 1) -join ", "
                }
            } else { # Fallback if no gateway or primary IP found via Get-NetIPConfiguration
                $ipAddress = (Get-NetIPAddress -InterfaceIndex $activeAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress | Select-Object -First 1)
                if (-not $ipAddress) {$ipAddress = "Không có địa chỉ IPv4"}
            }
        }
        $sb.AppendLine("  Địa chỉ IP: $ipAddress")
        $sb.AppendLine("  Địa chỉ MAC: $macAddress")

        # Cấu hình Phần cứng
        $sb.AppendLine("`n**Cấu hình Phần cứng:**")

        # CPU
        $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        $sb.AppendLine("  CPU:")
        if ($cpu) {
            $sb.AppendLine("    Kiểu máy: $($cpu.Name)")
            $sb.AppendLine("    Số lõi: $($cpu.NumberOfCores)")
            $sb.AppendLine("    Số luồng: $($cpu.NumberOfLogicalProcessors)")
        } else { $sb.AppendLine("    Không thể lấy thông tin CPU.")}


        # Bộ nhớ RAM (using $os from above)
        $totalRamGB = if ($os) { [math]::Round($os.TotalVisibleMemorySize / 1MB, 2) } else { "Không xác định" }
        $sb.AppendLine("  Bộ nhớ RAM: $totalRamGB GB")

        # Mainboard
        $mainboard = Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction SilentlyContinue | Select-Object -First 1
        $sb.AppendLine("  Mainboard:")
        if ($mainboard) {
            $sb.AppendLine("    Nhà sản xuất: $($mainboard.Manufacturer)")
            $sb.AppendLine("    Kiểu máy: $($mainboard.Product)")
            $sb.AppendLine("    Số Sê-ri: $($mainboard.SerialNumber)")
        } else { $sb.AppendLine("    Không thể lấy thông tin Mainboard.")}


        # Ổ đĩa - Modified to include fallback
        $sb.AppendLine("  Ổ đĩa:")
        if (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue) {
            $physicalDisks = Get-PhysicalDisk -ErrorAction SilentlyContinue
            if ($physicalDisks) {
                $diskIndex = 1
                foreach ($disk in $physicalDisks) {
                    $sb.AppendLine("    - Ổ đĩa ${diskIndex}:")
                    $sb.AppendLine("        Kiểu máy: $($disk.FriendlyName)")
                    $sb.AppendLine("        Dung lượng (GB): $([math]::Round($disk.Size / 1GB, 0))")
                    $sb.AppendLine("        Giao tiếp: $($disk.BusType)")
                    $sb.AppendLine("        Loại phương tiện: $($disk.MediaType)")
                    $diskIndex++
                }
            } else {
                $sb.AppendLine("    Không tìm thấy thông tin ổ đĩa vật lý (sử dụng Get-PhysicalDisk).")
            }
        } else { # Fallback for older PowerShell
            $wmiDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
            if ($wmiDisks) {
                $diskIndex = 1
                foreach ($disk in $wmiDisks) {
                    $sb.AppendLine("    - Ổ đĩa ${diskIndex}:")
                    $sb.AppendLine("        Kiểu máy: $($disk.Model)")
                    $sb.AppendLine("        Dung lượng (GB): $([math]::Round($disk.Size / 1GB, 0))")
                    $sb.AppendLine("        Giao tiếp: $($disk.InterfaceType)") # Less descriptive than BusType
                    $sb.AppendLine("        Loại phương tiện: $($disk.MediaType) (Thông tin chi tiết hơn cần Get-PhysicalDisk)")
                    $diskIndex++
                }
            } else { # Else này dành cho 'if ($wmiDisks)'
                $sb.AppendLine("    Không tìm thấy thông tin ổ đĩa vật lý (sử dụng Win32_DiskDrive).")
            }
        } # Dấu ngoặc này đóng khối 'else' cho 'if (Get-Command Get-PhysicalDisk ...)'

        # Card đồ họa (GPU)
        $sb.AppendLine("  Card đồ họa (GPU):")
        $gpus = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
        if ($gpus) {
            $gpuIndex = 1
            foreach ($gpu in $gpus) {
                $sb.AppendLine("    - GPU ${gpuIndex}:")
                $sb.AppendLine("        Tên: $($gpu.Name)")
                $sb.AppendLine("        Nhà sản xuất: $($gpu.AdapterCompatibility)")
                $adapterRamMB = if ($gpu.AdapterRAM) { [math]::Round($gpu.AdapterRAM / 1MB, 0) } else { "Không xác định" }
                $sb.AppendLine("        Tổng bộ nhớ (MB): $adapterRamMB")
                $sb.AppendLine("        Độ phân giải hiện tại: $($gpu.CurrentHorizontalResolution)x$($gpu.CurrentVerticalResolution)")
                $gpuIndex++
            }
        } else {
            $sb.AppendLine("    Không tìm thấy thông tin card đồ họa.")
        }

        # Màn hình
        $sb.AppendLine("  Màn hình:")
        $monitors = Get-CimInstance Win32_DesktopMonitor -ErrorAction SilentlyContinue
        if ($monitors) {
            $monitorIndex = 1
            foreach ($monitor in $monitors) {
                $sb.AppendLine("    - Màn hình ${monitorIndex}:")
                $monitorName = if ($monitor.Description -ne "Generic PnP Monitor" -and $monitor.Description) { $monitor.Description } elseif ($monitor.Name) { $monitor.Name } else { "Không xác định" }
                $sb.AppendLine("        Tên: $monitorName")
                $sb.AppendLine("        Độ phân giải: $($monitor.ScreenWidth)x$($monitor.ScreenHeight)")
                $sb.AppendLine("        Trạng thái: $($monitor.Status)")
                $monitorIndex++
            }
        } else {
            $sb.AppendLine("    Không tìm thấy thông tin màn hình.")
        }

        Write-GuiOutput $sb.ToString() -Color $global:colorScheme.DefaultText

    } catch {
        Write-GuiOutput "Lỗi khi lấy thông tin hệ thống: $($_.Exception.Message)" -Color $global:colorScheme.Error
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Get-DiskInfo-GUI {
    Clear-GuiOutput
    try {
        Write-GuiOutput "------ DUNG LƯỢNG Ổ ĐĨA ------" -Color $global:colorScheme.Title -Bold $true
        $diskInfo = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | 
                    Select-Object DeviceID, 
                                  @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}, 
                                  @{Name="TotalSize(GB)";Expression={[math]::Round($_.Size/1GB,2)}}, 
                                  @{Name="FreeSpace(%)";Expression={[math]::Round($_.FreeSpace/$_.Size*100,2)}}
        Write-GuiOutput ($diskInfo | Format-Table -AutoSize | Out-String) -Color $global:colorScheme.DefaultText
    } catch {
        Write-GuiOutput "Lỗi khi lấy thông tin ổ đĩa: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Invoke-MpScanAndReport-GUI {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScanAction,
        [string]$InitiatingMessage = "Đang bắt đầu quét...",
        [string]$SuccessMessage = "Quét hoàn tất!",
        [string]$FailureMessage = "Lỗi khi thực hiện quét:"
    )
    Write-GuiOutput $InitiatingMessage -Color $global:colorScheme.Info
    try {
        # This will still open its own progress window for Start-MpScan
        & $ScanAction 
        Write-GuiOutput $SuccessMessage -Color $global:colorScheme.Success
    } catch {
        Write-GuiOutput "$FailureMessage $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Check-Malware-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ KIỂM TRA & QUÉT PHẦN MỀM ĐỘC HẠI ------" -Color $global:colorScheme.Title -Bold $true
    Write-GuiOutput "Chức năng này yêu cầu một giao diện người dùng phức tạp hơn (ví dụ: một cửa sổ dialog riêng với các tùy chọn quét)." -Color $global:colorScheme.Warning
    Write-GuiOutput "Hiện tại, bạn có thể chọn một hành động quét mặc định hoặc triển khai một dialog mới." -Color $global:colorScheme.Warning

    # Example: Directly offer a quick scan or full scan via confirmation
    if (Confirm-Action-GUI -Message "Bạn có muốn thực hiện Quét nhanh (Quick Scan)?") {
        Invoke-MpScanAndReport-GUI -ScanAction { Start-MpScan -ScanType QuickScan }
    } elseif (Confirm-Action-GUI -Message "Bạn có muốn thực hiện Quét toàn bộ (Full Scan)?") {
        Invoke-MpScanAndReport-GUI -ScanAction { Start-MpScan -ScanType FullScan }
    } elseif (Confirm-Action-GUI -Message "Bạn có muốn Cập nhật định nghĩa virus?") {
        Write-GuiOutput "Đang cập nhật định nghĩa virus cho Windows Defender..." -Color $global:colorScheme.Info
        try {
            Update-MpSignature
            Write-GuiOutput "Cập nhật định nghĩa virus hoàn tất." -Color $global:colorScheme.Success
        } catch { Write-GuiOutput "Lỗi khi cập nhật định nghĩa virus: $($_.Exception.Message)" -Color $global:colorScheme.Error }
    } else {
        Write-GuiOutput "Không có hành động quét nào được chọn." -Color $global:colorScheme.Warning
    }
    # For custom scan, file scan, offline scan, a more dedicated UI (new form/dialog) would be needed
    # to get path inputs etc.
}

function Get-InstalledSoftware { # Keep original logic, GUI will call and format
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $allSoftwareEntries = @()
    foreach ($path in $registryPaths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.DisplayName -and $_.DisplayVersion) {
                $allSoftwareEntries += [PSCustomObject]@{
                    Name        = $_.DisplayName.Trim()
                    Version     = $_.DisplayVersion.Trim()
                    Publisher   = $_.Publisher
                    InstallDate = $_.InstallDate
                }
            }
        }
    }
    if ($allSoftwareEntries.Count -gt 0) {
        return $allSoftwareEntries | 
               Sort-Object Name, @{Expression='Publisher'; Descending=$true} | 
               Group-Object Name, Version | 
               ForEach-Object { $_.Group[0] } | 
               Sort-Object Name
    } else {
        return @()
    }
}

function Check-SoftwareVersion-GUI {
    Clear-GuiOutput
    try {
        Write-GuiOutput "------ PHIÊN BẢN PHẦN MỀM ------" -Color $global:colorScheme.Title -Bold $true
        Write-GuiOutput "Đang lấy danh sách phần mềm, việc này có thể mất một chút thời gian..." -Color $global:colorScheme.Info
        $software = Get-InstalledSoftware
        if ($software) {
            Write-GuiOutput ($software | Select-Object Name, Version, Publisher | Format-Table -AutoSize | Out-String) -Color $global:colorScheme.DefaultText
        } else {
            Write-GuiOutput "Không tìm thấy thông tin phần mềm đã cài đặt hoặc không có quyền truy cập Registry." -Color $global:colorScheme.Warning
        }
    } catch {
        Write-GuiOutput "Lỗi khi lấy phiên bản phần mềm: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Check-Activation-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ KIỂM TRA KÍCH HOẠT WINDOWS VÀ OFFICE ------" -Color $global:colorScheme.Title -Bold $true
    Write-GuiOutput "--- KIỂM TRA KÍCH HOẠT WINDOWS ---" -Color $global:colorScheme.Title
    # cscript output will go to a new console window. Capturing it is more complex.
    # For GUI, it's better to parse the WMI object if possible, or accept the console pop-up.
    try {
        $windowsActivation = cscript //nologo c:\windows\system32\slmgr.vbs /xpr
        Write-GuiOutput ($windowsActivation | Out-String) -Color $global:colorScheme.DefaultText
    } catch {
        Write-GuiOutput "Lỗi khi chạy slmgr.vbs: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
    Write-GuiOutput ""
    Write-GuiOutput "--- KIỂM TRA KÍCH HOẠT OFFICE ---" -Color $global:colorScheme.Title
    Write-GuiOutput "Lưu ý: Kiểm tra Office qua ospp.vbs áp dụng cho bản cũ. Với M365/Office mới, kiểm tra trong ứng dụng." -Color $global:colorScheme.Warning

    $officePaths = @{
        "Office16" = @("$env:ProgramFiles\Microsoft Office\Office16\ospp.vbs", "$env:ProgramFiles(x86)\Microsoft Office\Office16\ospp.vbs")
        "Office15" = @("$env:ProgramFiles\Microsoft Office\Office15\ospp.vbs", "$env:ProgramFiles(x86)\Microsoft Office\Office15\ospp.vbs")
        "Office14" = @("$env:ProgramFiles\Microsoft Office\Office14\ospp.vbs", "$env:ProgramFiles(x86)\Microsoft Office\Office14\ospp.vbs")
    }
    $foundOspp = $false
    foreach ($version in $officePaths.Keys) {
        foreach ($path in $officePaths[$version]) {
            if (Test-Path $path) { # The fix for ParserError is here
                Write-GuiOutput "Đang thử kiểm tra Office (phiên bản ${version}) tại: $path" -Color $global:colorScheme.Info
                try {
                    $officeActivation = cscript //nologo $path /dstatus
                    Write-GuiOutput ($officeActivation | Out-String) -Color $global:colorScheme.DefaultText
                } catch {
                    Write-GuiOutput "Lỗi khi chạy ospp.vbs cho ${version}: $($_.Exception.Message)" -Color $global:colorScheme.Error
                }
                $foundOspp = $true
                break 
            }
        }
        if ($foundOspp) { break }
    }
    if (-not $foundOspp) {
        Write-GuiOutput "Không tìm thấy tệp ospp.vbs cho Office 2010-2016." -Color $global:colorScheme.Warning
    }
}

function Check-BatteryHealth-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ KIỂM TRA ĐỘ CHAI PIN ------" -Color $global:colorScheme.Title -Bold $true
    try {
        $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($null -eq $battery) {
            Write-GuiOutput "Không tìm thấy pin trên thiết bị này (có thể là máy tính để bàn)." -Color $global:colorScheme.Warning
            return
        }
        $reportPath = "$env:USERPROFILE\battery-report.html"
        powercfg /batteryreport /output "$reportPath" /Duration 1 # Shorten duration for faster report
        Write-GuiOutput "Báo cáo độ chai pin đã được tạo tại '$reportPath'" -Color $global:colorScheme.Success
        Write-GuiOutput "Vui lòng mở tệp này bằng trình duyệt để xem chi tiết." -Color $global:colorScheme.Info
        if (Confirm-Action-GUI -Message "Bạn có muốn mở báo cáo pin ngay bây giờ không?") {
            try { Start-Process $reportPath } catch { Write-GuiOutput "Không thể tự động mở báo cáo: $($_.Exception.Message)" -Color $global:colorScheme.Error }
        }
    } catch {
        Write-GuiOutput "Lỗi khi kiểm tra pin: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Check-WifiConnection-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ KIỂM TRA KẾT NỐI WIFI ------" -Color $global:colorScheme.Title -Bold $true
    try {
        $allWifiAdapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object {$_.MediaType -eq "Native 802.11"}
        if ($null -eq $allWifiAdapters -or $allWifiAdapters.Count -eq 0) {
            Write-GuiOutput "Không tìm thấy card mạng Wi-Fi nào." -Color $global:colorScheme.Warning
            return
        }
        $activeWifiAdapter = $allWifiAdapters | Where-Object {$_.Status -eq "Up"}
        if ($null -eq $activeWifiAdapter -or $activeWifiAdapter.Count -eq 0) {
            Write-GuiOutput "Không có kết nối Wi-Fi nào đang hoạt động." -Color $global:colorScheme.Warning
            Write-GuiOutput "Thông tin các card Wi-Fi có sẵn:" -Color $global:colorScheme.Info
            Write-GuiOutput ($allWifiAdapters | Select-Object Name, InterfaceDescription, Status, MacAddress | Format-List | Out-String) -Color $global:colorScheme.DefaultText
            Write-GuiOutput "`nCác mạng Wi-Fi đã lưu (profiles):" -Color $global:colorScheme.Info
            Write-GuiOutput (netsh wlan show profiles | Out-String) -Color $global:colorScheme.DefaultText
        } else {
            Write-GuiOutput "Thông tin kết nối Wi-Fi đang hoạt động:" -Color $global:colorScheme.Success
            Write-GuiOutput (netsh wlan show interfaces | Out-String) -Color $global:colorScheme.DefaultText
        }
    } catch {
        Write-GuiOutput "Lỗi khi kiểm tra Wifi: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Check-DriveAndTemperature-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ KIỂM TRA Ổ CỨNG ------" -Color $global:colorScheme.Title -Bold $true
    try {
        if (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue) {
            $physicalDisks = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, HealthStatus, Size 
            Write-GuiOutput ($physicalDisks | Format-Table -AutoSize | Out-String) -Color $global:colorScheme.DefaultText
        } else {
            Write-GuiOutput "Thông tin chi tiết về ổ cứng (Get-PhysicalDisk) yêu cầu PowerShell 4.0+. Hiển thị thông tin cơ bản từ Win32_DiskDrive:" -Color $global:colorScheme.Warning
            $wmiDisks = Get-CimInstance -ClassName Win32_DiskDrive | Select-Object DeviceID, Model, @{Name="Size(GB)";Expression={[math]::Round($_.Size/1GB,0)}}, Status, InterfaceType, MediaType
            Write-GuiOutput ($wmiDisks | Format-Table -AutoSize | Out-String) -Color $global:colorScheme.DefaultText
            Write-GuiOutput "(Lưu ý: 'MediaType' và 'HealthStatus' từ Win32_DiskDrive có thể kém chi tiết hơn Get-PhysicalDisk)" -Color $global:colorScheme.Warning
        }
        Write-GuiOutput ""
        Write-GuiOutput "------ KIỂM TRA NHIỆT ĐỘ CPU/HỆ THỐNG (Thử nghiệm) ------" -Color $global:colorScheme.Title -Bold $true
        try {
            $tempObjects = Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ErrorAction Stop
            if ($tempObjects) {
                Write-GuiOutput "Nhiệt độ CPU/Hệ thống (Celsius):" -Color $global:colorScheme.Success
                $tempData = $tempObjects | ForEach-Object {
                    $deviceName = $_.InstanceName -replace ".*ThermalZone$", "ThermalZone" -replace ".*\\", ""
                    $tempCelsius = [Math]::Round(($_.CurrentTemperature / 10) - 273.15, 1)
                    "$deviceName : $tempCelsius °C"
                }
                Write-GuiOutput ($tempData -join "`n") -Color $global:colorScheme.DefaultText
            } else {
                Write-GuiOutput "Không thể lấy thông tin nhiệt độ qua WMI." -Color $global:colorScheme.Warning
            }
        } catch {
            Write-GuiOutput "Lỗi lấy nhiệt độ WMI: $($_.Exception.Message)" -Color $global:colorScheme.Warning
            Write-GuiOutput "(Ghi chú: Kiểm tra nhiệt độ CPU chính xác thường cần phần mềm của bên thứ ba)" -Color $global:colorScheme.Warning
        }
    } catch {
        Write-GuiOutput "Lỗi khi kiểm tra ổ cứng hoặc nhiệt độ: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Manage-BackgroundAppsAndProcesses-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ QUẢN LÝ ỨNG DỤNG & PHẦN MỀM CHẠY NGẦM ------" -Color $global:colorScheme.Title -Bold $true
    Write-GuiOutput "Chức năng này cần một giao diện người dùng tương tác hơn để chọn và dừng tiến trình." -Color $global:colorScheme.Warning
    Write-GuiOutput "Ví dụ: Hiển thị danh sách các tiến trình trong một ListBox hoặc DataGridView." -Color $global:colorScheme.Warning
    
    # Simplified: Just list processes for now
    $bgapps = Get-Process | Sort-Object CPU -Descending | Select-Object ProcessName, Id, @{Name="CPU(s)"; Expression={$_.CPU.TotalSeconds.ToString('F0')}}, MainWindowTitle
    Write-GuiOutput "DANH SÁCH TOÀN BỘ PHẦN MỀM ĐANG CHẠY NGẦM (Sắp xếp theo CPU):" -Color $global:colorScheme.Info
    Write-GuiOutput ($bgapps | Format-Table -AutoSize | Out-String) -Color $global:colorScheme.DefaultText
    Write-GuiOutput "`nTổng cộng: $($bgapps.Count) tiến trình đang chạy." -Color $global:colorScheme.DefaultText
    Show-GuiMessage "Để dừng một tiến trình, bạn có thể sử dụng Task Manager hoặc một công cụ quản lý tiến trình chuyên dụng. Việc tích hợp chức năng dừng vào GUI này cần phát triển thêm."
}

function Reset-Network-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ RESET KẾT NỐI INTERNET ------" -Color $global:colorScheme.Title -Bold $true
    if (Confirm-Action-GUI -Message "Bạn có chắc chắn muốn reset cài đặt mạng? Điều này có thể yêu cầu khởi động lại.") {
        try {
            ipconfig /flushdns; Write-GuiOutput "Đã xóa cache DNS." -Color $global:colorScheme.Success
            ipconfig /release; Write-GuiOutput "Đã giải phóng địa chỉ IP." -Color $global:colorScheme.Success
            ipconfig /renew; Write-GuiOutput "Đã làm mới địa chỉ IP." -Color $global:colorScheme.Success
            netsh winsock reset; Write-GuiOutput "Đã reset Winsock catalog. (Cần khởi động lại)" -Color $global:colorScheme.Warning
            netsh int ip reset; Write-GuiOutput "Đã reset TCP/IP stack. (Cần khởi động lại)" -Color $global:colorScheme.Warning
            Write-GuiOutput "Hoàn tất! Vui lòng khởi động lại máy tính." -Color $global:colorScheme.Success
        } catch {
            Write-GuiOutput "Lỗi khi reset mạng: $($_.Exception.Message)" -Color $global:colorScheme.Error
        }
    } else {
        Write-GuiOutput "Hành động reset mạng đã bị hủy." -Color $global:colorScheme.Warning
    }
}

function Clear-TempFiles-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ XÓA FILE TẠM, DỌN DẸP PREFETCH VÀ THÙNG RÁC ------" -Color $global:colorScheme.Title -Bold $true
    if (Confirm-Action-GUI -Message "Bạn có chắc chắn muốn xóa các file tạm, dọn dẹp Prefetch và Thùng rác?") {
        try {
            Write-GuiOutput "Đang xóa $env:TEMP..." -Color $global:colorScheme.Info
            Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-GuiOutput "Đã xóa $env:TEMP." -Color $global:colorScheme.Success

            Write-GuiOutput "Đang xóa C:\Windows\Temp..." -Color $global:colorScheme.Info
            Remove-Item "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-GuiOutput "Đã xóa C:\Windows\Temp." -Color $global:colorScheme.Success
            
            Write-GuiOutput "Đang xóa C:\Windows\SoftwareDistribution\Download..." -Color $global:colorScheme.Info
            Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-GuiOutput "Đã xóa C:\Windows\SoftwareDistribution\Download." -Color $global:colorScheme.Success

            Write-GuiOutput "Đang dọn dẹp Prefetch..." -Color $global:colorScheme.Info
            try { Remove-Item "C:\Windows\Prefetch\*" -Recurse -Force -ErrorAction Stop; Write-GuiOutput "Đã dọn dẹp Prefetch." -Color $global:colorScheme.Success } 
            catch { Write-GuiOutput "Lỗi dọn Prefetch: $($_.Exception.Message)" -Color $global:colorScheme.Warning } # Changed to Warning as it might be due to permissions/in-use files

            Write-GuiOutput "Đang dọn dẹp Thùng rác..." -Color $global:colorScheme.Info
            try { 
                if (Get-Command Clear-RecycleBin -ErrorAction SilentlyContinue) {
                    Clear-RecycleBin -Force -ErrorAction Stop
                    Write-GuiOutput "Đã dọn dẹp Thùng rác." -Color $global:colorScheme.Success
                } else {
                    Write-GuiOutput "Lệnh 'Clear-RecycleBin' không khả dụng trên phiên bản PowerShell này (cần PS 5.0+). Bỏ qua dọn dẹp Thùng rác." -Color $global:colorScheme.Warning
                }
            } 
            catch { Write-GuiOutput "Lỗi dọn Thùng rác: $($_.Exception.Message)" -Color $global:colorScheme.Error } # Keep Error for actual cmdlet failure
            
            Write-GuiOutput "Hoàn tất dọn dẹp." -Color $global:colorScheme.Success
        } catch {
            Write-GuiOutput "Lỗi trong quá trình dọn dẹp: $($_.Exception.Message)" -Color $global:colorScheme.Error
        }
    } else {
        Write-GuiOutput "Hành động dọn dẹp đã bị hủy." -Color $global:colorScheme.Warning
    }
}

function Run-SFCScan-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ KIỂM TRA VÀ SỬA LỖI TỆP HỆ THỐNG (SFC SCAN) ------" -Color $global:colorScheme.Title -Bold $true
    Write-GuiOutput "Đang thực hiện SFC Scan. Quá trình này có thể mất một lúc. Vui lòng đợi..." -Color $global:colorScheme.Info
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $tempSfcLogPath = $null
        $tempSfcLogObject = $null # For New-TemporaryFile object if used

        if (Get-Command New-TemporaryFile -ErrorAction SilentlyContinue) {
            $tempSfcLogObject = New-TemporaryFile
            $tempSfcLogPath = $tempSfcLogObject.FullName
        } else {
            $tempSfcLogPath = [System.IO.Path]::GetTempFileName()
        }

        $process = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Verb RunAs -Wait -NoNewWindow -PassThru -RedirectStandardOutput $tempSfcLogPath -ErrorAction Stop
        
        $sfcOutput = Get-Content $tempSfcLogPath -Raw -ErrorAction SilentlyContinue
        
        if ($tempSfcLogObject -is [System.IO.FileInfo]) { # If New-TemporaryFile was used
             Remove-Item $tempSfcLogObject.FullName -Force -ErrorAction SilentlyContinue
        } elseif ($tempSfcLogPath) { # If GetTempFileName was used
             Remove-Item $tempSfcLogPath -Force -ErrorAction SilentlyContinue
        }

        if ($process.ExitCode -eq 0) {
            Write-GuiOutput "SFC Scan hoàn tất thành công (theo mã thoát)." -Color $global:colorScheme.Success
        } else {
            Write-GuiOutput "SFC Scan hoàn tất với mã lỗi: $($process.ExitCode)." -Color $global:colorScheme.Warning
        }
        Write-GuiOutput "Kết quả SFC Scan (nếu có output trực tiếp):" -Color $global:colorScheme.Info
        if ($sfcOutput) {
            Write-GuiOutput $sfcOutput -Color $global:colorScheme.DefaultText
        } else {
            Write-GuiOutput "Không có output trực tiếp từ SFC.EXE. Vui lòng kiểm tra log tại C:\Windows\Logs\CBS\CBS.log để biết chi tiết." -Color $global:colorScheme.Warning
        }
    } catch {
        Write-GuiOutput "Lỗi khi chạy SFC Scan: $($_.Exception.Message)" -Color $global:colorScheme.Error
        Write-GuiOutput "Vui lòng kiểm tra log tại C:\Windows\Logs\CBS\CBS.log để biết chi tiết." -Color $global:colorScheme.Warning
    } finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        if ($tempSfcLogPath -and (Test-Path $tempSfcLogPath)) { # Ensure cleanup if error occurred mid-way
            Remove-Item $tempSfcLogPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Create-RestorePoint-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ TẠO ĐIỂM KHÔI PHỤC HỆ THỐNG ------" -Color $global:colorScheme.Title -Bold $true
    # For GUI, input should come from an input box.
    # Using a simple default for now or prompt.
    $description = "HealthCheckTool Restore Point $(Get-Date -Format 'yyyy-MM-dd HH:mm')" 
    # Add InputBox for description later if needed
    # $description = [Microsoft.VisualBasic.Interaction]::InputBox("Nhập mô tả cho điểm khôi phục:", "Tạo điểm khôi phục", "Trước khi thay đổi hệ thống")
    # if ([string]::IsNullOrWhiteSpace($description)) { Write-GuiOutput "Hủy tạo điểm khôi phục." -Color $global:colorScheme.Warning; return }

    if (Confirm-Action-GUI -Message "Bạn có muốn tạo điểm khôi phục với mô tả '$description'?") {
        try {
            Checkpoint-Computer -Description $description -ErrorAction Stop
            Write-GuiOutput "Đã tạo điểm khôi phục hệ thống thành công: $description" -Color $global:colorScheme.Success
        } catch {
            Write-GuiOutput "Lỗi khi tạo điểm khôi phục: $($_.Exception.Message)" -Color $global:colorScheme.Error
        }
    } else {
         Write-GuiOutput "Hủy tạo điểm khôi phục." -Color $global:colorScheme.Warning
    }
}

function Check-DriverErrors-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ KIỂM TRA DRIVER BỊ LỖI ------" -Color $global:colorScheme.Title -Bold $true
    try {
        $errorDrivers = Get-CimInstance Win32_PnPEntity | Where-Object { $_.Status -ne "OK" -or $_.ConfigManagerErrorCode -ne 0 }
        if ($errorDrivers) {
            Write-GuiOutput "Các driver bị lỗi hoặc gặp sự cố:" -Color $global:colorScheme.Warning
            Write-GuiOutput ($errorDrivers | Select-Object Name, DeviceID, Status, ConfigManagerErrorCode | Format-Table -AutoSize | Out-String) -Color $global:colorScheme.DefaultText
        } else {
            Write-GuiOutput "Tất cả driver đều đang hoạt động bình thường!" -Color $global:colorScheme.Success
        }
    } catch {
        Write-GuiOutput "Lỗi khi kiểm tra driver: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Install-PSWindowsUpdateModuleIfNeeded {
    if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
        Write-GuiOutput "Lệnh 'Install-Module' không khả dụng. Module PSWindowsUpdate cần được cài đặt thủ công hoặc bạn cần cập nhật PowerShell (PowerShellGet module)." -Color $global:colorScheme.Error
        Show-GuiMessage "Chức năng này yêu cầu module PowerShellGet (thường có trong PowerShell 5.0+) để tự động cài đặt PSWindowsUpdate. Vui lòng cài đặt PSWindowsUpdate thủ công hoặc cập nhật PowerShell của bạn." -Title "Yêu cầu Module" -Icon Warning
        return $false
    }

    if (!(Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-GuiOutput "Module PSWindowsUpdate chưa được cài đặt. Đang tiến hành cài đặt..." -Color $global:colorScheme.Info
        try {
            # Added -Scope CurrentUser to avoid needing admin for Install-Module itself if PS is run as admin for other things
            # Confirm:$false to avoid interactive prompts from Install-Module
            Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Scope CurrentUser -Confirm:$false -ErrorAction Stop
            Write-GuiOutput "Đã cài đặt module PSWindowsUpdate." -Color $global:colorScheme.Success
        } catch {
            Write-GuiOutput "Lỗi khi cài đặt PSWindowsUpdate: $($_.Exception.Message)" -Color $global:colorScheme.Error
            Write-GuiOutput "Vui lòng thử cài đặt thủ công: Install-Module PSWindowsUpdate -Scope CurrentUser" -Color $global:colorScheme.Warning
            return $false
        }
    }
    # Ensure module is imported if already installed or just installed
    if (-not (Get-Module -Name PSWindowsUpdate)) {
        try {
            Import-Module PSWindowsUpdate -ErrorAction Stop
        } catch {
            Write-GuiOutput "Lỗi khi import module PSWindowsUpdate: $($_.Exception.Message)" -Color $global:colorScheme.Error
            return $false
        }
    }
    return $true
}


function Update-Drivers-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ CẬP NHẬT DRIVER TỰ ĐỘNG ------" -Color $global:colorScheme.Title -Bold $true
    if (-not (Install-PSWindowsUpdateModuleIfNeeded)) { return }
    
    if (Confirm-Action-GUI -Message "Bạn có chắc chắn muốn tìm và cài đặt các bản cập nhật driver?") {
        try {
            Write-GuiOutput "Đang tìm và cài đặt cập nhật driver..." -Color $global:colorScheme.Info
            Get-WindowsUpdate -MicrosoftUpdate -Category "Drivers" -AcceptAll -Install # -AutoReboot removed
            Write-GuiOutput "Hoàn tất kiểm tra và cập nhật driver! Một số cập nhật có thể yêu cầu khởi động lại." -Color $global:colorScheme.Success
        } catch {
            Write-GuiOutput "Lỗi khi cập nhật driver: $($_.Exception.Message)" -Color $global:colorScheme.Error
        }
    } else {
        Write-GuiOutput "Hành động cập nhật driver đã bị hủy." -Color $global:colorScheme.Warning
    }
}

function Update-Software-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ CẬP NHẬT PHẦN MỀM (qua winget) ------" -Color $global:colorScheme.Title -Bold $true
    if (!(Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-GuiOutput "Winget chưa được cài đặt! Vui lòng cài đặt hoặc cập nhật Windows." -Color $global:colorScheme.Warning
        return
    }
    if (Confirm-Action-GUI -Message "Bạn có chắc chắn muốn tìm và cài đặt các bản cập nhật phần mềm qua winget? (Sẽ mở cửa sổ console riêng)") {
        try {
            Write-GuiOutput "Đang chạy 'winget upgrade --all'. Vui lòng theo dõi cửa sổ console được mở..." -Color $global:colorScheme.Info
            Start-Process cmd -ArgumentList "/c winget upgrade --all & pause" -Verb RunAs -Wait
            Write-GuiOutput "Winget upgrade hoàn tất." -Color $global:colorScheme.Success
        } catch {
            Write-GuiOutput "Lỗi khi cập nhật phần mềm: $($_.Exception.Message)" -Color $global:colorScheme.Error
        }
    } else {
        Write-GuiOutput "Hành động cập nhật phần mềm đã bị hủy." -Color $global:colorScheme.Warning
    }
}

function Update-Windows-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ CẬP NHẬT WINDOWS TỰ ĐỘNG ------" -Color $global:colorScheme.Title -Bold $true
    if (-not (Install-PSWindowsUpdateModuleIfNeeded)) { return }

    if (Confirm-Action-GUI -Message "Bạn có chắc chắn muốn tìm và cài đặt các bản cập nhật Windows?") {
        try {
            Write-GuiOutput "Đang tìm và cài đặt cập nhật Windows..." -Color $global:colorScheme.Info
            Get-WindowsUpdate -AcceptAll -Install 
            Write-GuiOutput "Hoàn tất! Một số cập nhật có thể yêu cầu khởi động lại." -Color $global:colorScheme.Success
        } catch {
            Write-GuiOutput "Lỗi khi cập nhật Windows: $($_.Exception.Message)" -Color $global:colorScheme.Error
        }
    } else {
        Write-GuiOutput "Hành động cập nhật Windows đã bị hủy." -Color $global:colorScheme.Warning
    }
}

function View-RecentEventLogErrors-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ XEM CÁC LỖI NGHIÊM TRỌNG GẦN ĐÂY (EVENT LOG - 24 GIỜ QUA) ------" -Color $global:colorScheme.Title -Bold $true
    try {
        $startTime = (Get-Date).AddDays(-1)
        $sb = New-Object System.Text.StringBuilder

        $sb.AppendLine("--- Lỗi trong System Log ---")
        $systemErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2; StartTime=$startTime} -MaxEvents 20 -ErrorAction SilentlyContinue
        if ($systemErrors) {
            $sb.AppendLine(($systemErrors | Format-Table TimeCreated, Id, LevelDisplayName, Message -AutoSize -Wrap | Out-String))
        } else { $sb.AppendLine("Không có lỗi nghiêm trọng nào trong System Log 24 giờ qua.") }

        $sb.AppendLine("`n--- Lỗi trong Application Log ---")
        $appErrors = Get-WinEvent -FilterHashtable @{LogName='Application'; Level=1,2; StartTime=$startTime} -MaxEvents 20 -ErrorAction SilentlyContinue
        if ($appErrors) {
            $sb.AppendLine(($appErrors | Format-Table TimeCreated, Id, LevelDisplayName, Message -AutoSize -Wrap | Out-String))
        } else { $sb.AppendLine("Không có lỗi nghiêm trọng nào trong Application Log 24 giờ qua.") }
        
        Write-GuiOutput $sb.ToString() -Color $global:colorScheme.DefaultText

    } catch {
        Write-GuiOutput "Lỗi khi xem Event Log: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

#region New Utility Functions
function Open-ResourceMonitor-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ RESOURCE MONITOR ------" -Color $global:colorScheme.Title -Bold $true
    try {
        Start-Process resmon.exe
        Write-GuiOutput "Đã khởi chạy Resource Monitor." -Color $global:colorScheme.Success
    } catch {
        Write-GuiOutput "Lỗi khi mở Resource Monitor: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Open-DiskCleanup-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ DISK CLEANUP ------" -Color $global:colorScheme.Title -Bold $true
    try {
        Start-Process cleanmgr.exe
        Write-GuiOutput "Đã khởi chạy Disk Cleanup. Vui lòng chọn ổ đĩa và tùy chọn trong cửa sổ mới." -Color $global:colorScheme.Info
    } catch {
        Write-GuiOutput "Lỗi khi mở Disk Cleanup: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Clear-PrintSpooler-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ XÓA HÀNG ĐỢI IN (CLEAR PRINT SPOOLER) ------" -Color $global:colorScheme.Title -Bold $true
    if (Confirm-Action-GUI -Message "Bạn có chắc chắn muốn dừng dịch vụ Print Spooler, xóa các tệp trong hàng đợi và khởi động lại dịch vụ không? Thao tác này sẽ xóa tất cả các lệnh in đang chờ.") {
        try {
            Write-GuiOutput "Đang dừng dịch vụ Print Spooler..." -Color $global:colorScheme.Info
            Stop-Service -Name Spooler -Force -ErrorAction Stop
            Write-GuiOutput "Dịch vụ Print Spooler đã dừng." -Color $global:colorScheme.Success

            Write-GuiOutput "Đang xóa các tệp trong hàng đợi in (C:\Windows\System32\spool\PRINTERS\)..." -Color $global:colorScheme.Info
            $printQueuePath = "C:\Windows\System32\spool\PRINTERS\*"
            Remove-Item -Path $printQueuePath -Force -ErrorAction SilentlyContinue # SilentlyContinue as folder might be empty or access issues
            Write-GuiOutput "Đã xóa các tệp trong hàng đợi." -Color $global:colorScheme.Success

            Write-GuiOutput "Đang khởi động lại dịch vụ Print Spooler..." -Color $global:colorScheme.Info
            Start-Service -Name Spooler -ErrorAction Stop
            Write-GuiOutput "Dịch vụ Print Spooler đã được khởi động lại." -Color $global:colorScheme.Success
            Write-GuiOutput "Hàng đợi in đã được dọn dẹp." -Color $global:colorScheme.Success
        } catch {
            Write-GuiOutput "Lỗi khi dọn dẹp hàng đợi in: $($_.Exception.Message)" -Color $global:colorScheme.Error
            Write-GuiOutput "Hãy thử khởi động lại dịch vụ 'Print Spooler' thủ công nếu cần (services.msc)." -Color $global:colorScheme.Warning
        }
    } else {
        Write-GuiOutput "Hành động dọn dẹp hàng đợi in đã bị hủy." -Color $global:colorScheme.Warning
    }
}

function Open-DeviceManager-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ DEVICE MANAGER ------" -Color $global:colorScheme.Title -Bold $true
    try {
        Start-Process devmgmt.msc
        Write-GuiOutput "Đã khởi chạy Device Manager." -Color $global:colorScheme.Success
    } catch {
        Write-GuiOutput "Lỗi khi mở Device Manager: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Open-SystemConfiguration-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ SYSTEM CONFIGURATION (MSCONFIG) ------" -Color $global:colorScheme.Title -Bold $true
    try {
        Start-Process msconfig.exe
        Write-GuiOutput "Đã khởi chạy System Configuration." -Color $global:colorScheme.Success
    } catch {
        Write-GuiOutput "Lỗi khi mở System Configuration: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Open-TaskScheduler-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ TASK SCHEDULER ------" -Color $global:colorScheme.Title -Bold $true
    try {
        Start-Process taskschd.msc
        Write-GuiOutput "Đã khởi chạy Task Scheduler." -Color $global:colorScheme.Success
    } catch {
        Write-GuiOutput "Lỗi khi mở Task Scheduler: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Open-EventViewer-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ EVENT VIEWER ------" -Color $global:colorScheme.Title -Bold $true
    try {
        Start-Process eventvwr.msc
        Write-GuiOutput "Đã khởi chạy Event Viewer." -Color $global:colorScheme.Success
    } catch {
        Write-GuiOutput "Lỗi khi mở Event Viewer: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Open-RegistryEditor-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ REGISTRY EDITOR ------" -Color $global:colorScheme.Title -Bold $true
    if (Confirm-Action-GUI -Message "CẢNH BÁO: Chỉnh sửa Registry không đúng cách có thể gây ra sự cố nghiêm trọng cho hệ thống. Bạn có chắc chắn muốn tiếp tục và mở Registry Editor không?" -Title "Cảnh báo Registry" -Icon Warning) {
        try {
            Start-Process regedit.exe
            Write-GuiOutput "Đã khởi chạy Registry Editor. Hãy cẩn thận khi thực hiện thay đổi." -Color $global:colorScheme.Warning
        } catch {
            Write-GuiOutput "Lỗi khi mở Registry Editor: $($_.Exception.Message)" -Color $global:colorScheme.Error
        }
    } else {
        Write-GuiOutput "Hành động mở Registry Editor đã bị hủy." -Color $global:colorScheme.Warning
    }
}

function Open-ControlPanel-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ CONTROL PANEL ------" -Color $global:colorScheme.Title -Bold $true
    try {
        Start-Process control.exe
        Write-GuiOutput "Đã khởi chạy Control Panel." -Color $global:colorScheme.Success
    } catch {
        Write-GuiOutput "Lỗi khi mở Control Panel: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

function Open-WindowsSettings-GUI {
    Clear-GuiOutput
    Write-GuiOutput "------ MỞ WINDOWS SETTINGS ------" -Color $global:colorScheme.Title -Bold $true
    try {
        Start-Process "ms-settings:"
        Write-GuiOutput "Đã khởi chạy Windows Settings." -Color $global:colorScheme.Success
    } catch {
        Write-GuiOutput "Lỗi khi mở Windows Settings: $($_.Exception.Message)" -Color $global:colorScheme.Error
    }
}

#endregion

#region GUI Setup

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Công Cụ Kiểm Tra Sức Khỏe Máy Tính"
$mainForm.Size = New-Object System.Drawing.Size(800, 700) # Increased width for better layout
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = 'Sizable' # Allow resizing
$mainForm.MaximizeBox = $true       # Show maximize button
$mainForm.MinimizeBox = $true       # Show minimize button
$mainForm.BackColor = $global:colorScheme.FormBackground

$defaultFont = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$outputFont = New-Object System.Drawing.Font("Consolas", 9.5) # Slightly larger Consolas
$mainForm.Font = $defaultFont

# Output RichTextBox
$Global:OutputRichTextBox = New-Object System.Windows.Forms.RichTextBox
$Global:OutputRichTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$Global:OutputRichTextBox.Font = $outputFont
$Global:OutputRichTextBox.ReadOnly = $true
$Global:OutputRichTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$Global:OutputRichTextBox.HideSelection = $false # Keep selection visible even when not focused
$Global:OutputRichTextBox.BackColor = $global:colorScheme.OutputBackground
$Global:OutputRichTextBox.ForeColor = $global:colorScheme.OutputForeground
$Global:OutputRichTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

# Panel for buttons
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Left
$buttonPanel.Width = 250
$buttonPanel.AutoScroll = $true
$buttonPanel.BackColor = $global:colorScheme.PanelBackground 

# SplitContainer for resizable layout
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.Panel1.Controls.Add($buttonPanel)
$splitContainer.Panel2.Controls.Add($Global:OutputRichTextBox)
$splitContainer.SplitterDistance = 250
$splitContainer.IsSplitterFixed = $false # Allow user to drag the splitter
$mainForm.Controls.Add($splitContainer)

# Script-level variables for managing menu layout with GroupBoxes
$script:activeGroupBox = $null
$script:buttonYOffsetInGroup = 0 # Y-offset for buttons inside the current GroupBox
$script:overallMenuYOffset = 10  # Y-offset for the next GroupBox in the buttonPanel

function Add-ButtonToCurrentGroup {
    param (
        [string]$Text,
        [scriptblock]$OnClick
    )
    
    if (-not $script:activeGroupBox) {
        Write-Error "Lỗi: Không có nhóm menu nào đang hoạt động. Hãy gọi Start-NewMenuGroup trước."
        return
    }

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Width = $script:activeGroupBox.ClientSize.Width - 20 # Width relative to active GroupBox
    $button.Height = 38 # Increased height for better spacing
    $button.Location = New-Object System.Drawing.Point(10, $script:buttonYOffsetInGroup)
    $button.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 0
    $button.BackColor = $global:colorScheme.ButtonBackground
    $button.ForeColor = $global:colorScheme.ButtonForeground
    $button.Font = $buttonFont
    $button.Padding = New-Object System.Windows.Forms.Padding(10,0,0,0) # Adjusted left padding
    $button.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

    $button.Add_Click($OnClick)
    $button.Add_MouseEnter({$this.BackColor = $global:colorScheme.ButtonHoverBackground})
    $button.Add_MouseLeave({$this.BackColor = $global:colorScheme.ButtonBackground})

    $script:activeGroupBox.Controls.Add($button)
    $script:buttonYOffsetInGroup += $button.Height + 7 # Increased spacing between buttons
}

function Start-NewMenuGroup {
    param([string]$Title)

    # Finalize the previous group's height and update overall Y offset
    if ($script:activeGroupBox) {
        $requiredHeight = $script:buttonYOffsetInGroup + 10 # Content height + bottom padding
        $script:activeGroupBox.Height = [Math]::Max(50, $requiredHeight) # Ensure a minimum height for the groupbox
        $script:overallMenuYOffset += $script:activeGroupBox.Height + 15 # Increased spacing for the next group
    }

    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Text = $Title
    $groupBox.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold) # Slightly larger/bolder GroupBox title
    $groupBox.ForeColor = $global:colorScheme.LabelForeground
    $groupBox.Width = $buttonPanel.ClientSize.Width - 12 # Slightly less than panel width for margin
    $groupBox.Location = New-Object System.Drawing.Point(6, $script:overallMenuYOffset)
    $groupBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    # $groupBox.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220) # Optional: slightly different background for groupbox

    $buttonPanel.Controls.Add($groupBox)
    $script:activeGroupBox = $groupBox
    $script:buttonYOffsetInGroup = 30 # Reset Y for buttons inside the new group (increased padding from groupbox top border + title)
}

function FinalizeLastMenuGroup {
    if ($script:activeGroupBox) {
        $requiredHeight = $script:buttonYOffsetInGroup + 10
        $script:activeGroupBox.Height = [Math]::Max(50, $requiredHeight)
    }
}

# --- Menu Structure ---
Start-NewMenuGroup "KIỂM TRA & GIÁM SÁT"
Add-ButtonToCurrentGroup "1. Thông tin hệ thống" { Get-SystemInfo-GUI }
Add-ButtonToCurrentGroup "2. Dung lượng ổ đĩa" { Get-DiskInfo-GUI }
Add-ButtonToCurrentGroup "3. Kiểm tra phần mềm độc hại" { Check-Malware-GUI }
Add-ButtonToCurrentGroup "4. Xem lỗi Event Log gần đây" { View-RecentEventLogErrors-GUI }
Add-ButtonToCurrentGroup "5. Kiểm tra phiên bản phần mềm" { Check-SoftwareVersion-GUI }
Add-ButtonToCurrentGroup "6. Kiểm tra kích hoạt Win/Office" { Check-Activation-GUI }
Add-ButtonToCurrentGroup "7. Kiểm tra độ chai pin" { Check-BatteryHealth-GUI }
Add-ButtonToCurrentGroup "8. Kiểm tra kết nối Wifi" { Check-WifiConnection-GUI }
Add-ButtonToCurrentGroup "9. Kiểm tra ổ cứng và nhiệt độ" { Check-DriveAndTemperature-GUI }
Add-ButtonToCurrentGroup "10. Quản lý ứng dụng chạy ngầm" { Manage-BackgroundAppsAndProcesses-GUI }
Add-ButtonToCurrentGroup "11. Mở Resource Monitor" { Open-ResourceMonitor-GUI }

Start-NewMenuGroup "SỬA LỖI & TỐI ƯU"
Add-ButtonToCurrentGroup "12. Reset kết nối Internet" { Reset-Network-GUI }
Add-ButtonToCurrentGroup "13. Xóa file tạm, dọn dẹp" { Clear-TempFiles-GUI }
Add-ButtonToCurrentGroup "14. SFC Scan (Kiểm tra file hệ thống)" { Run-SFCScan-GUI }
Add-ButtonToCurrentGroup "15. Tạo điểm khôi phục" { Create-RestorePoint-GUI }
Add-ButtonToCurrentGroup "16. Kiểm tra driver bị lỗi" { Check-DriverErrors-GUI }
Add-ButtonToCurrentGroup "17. Cập nhật driver tự động" { Update-Drivers-GUI }
Add-ButtonToCurrentGroup "18. Cập nhật phần mềm (winget)" { Update-Software-GUI }
Add-ButtonToCurrentGroup "19. Cập nhật Windows tự động" { Update-Windows-GUI }
Add-ButtonToCurrentGroup "20. Mở Disk Cleanup" { Open-DiskCleanup-GUI }
Add-ButtonToCurrentGroup "21. Xóa hàng đợi In" { Clear-PrintSpooler-GUI }

Start-NewMenuGroup "CÔNG CỤ HỆ THỐNG"
Add-ButtonToCurrentGroup "22. Mở Device Manager" { Open-DeviceManager-GUI }
Add-ButtonToCurrentGroup "23. Mở System Configuration" { Open-SystemConfiguration-GUI }
Add-ButtonToCurrentGroup "24. Mở Task Scheduler" { Open-TaskScheduler-GUI }
Add-ButtonToCurrentGroup "25. Mở Event Viewer" { Open-EventViewer-GUI }
Add-ButtonToCurrentGroup "26. Mở Registry Editor (Cảnh báo!)" { Open-RegistryEditor-GUI }
Add-ButtonToCurrentGroup "27. Mở Control Panel" { Open-ControlPanel-GUI }
Add-ButtonToCurrentGroup "28. Mở Windows Settings" { Open-WindowsSettings-GUI }

Start-NewMenuGroup "KHÁC"
Add-ButtonToCurrentGroup "29. Thoát" { $mainForm.Close() }

FinalizeLastMenuGroup # Finalize the layout of the last group

# Initial message
Write-GuiOutput "Chào mừng bạn đến với Tiện ích máy tính!" -Color $global:colorScheme.Info -Bold $true
Write-GuiOutput "Vui lòng chọn một chức năng từ menu bên trái." -Color $global:colorScheme.Info

$mainForm.Add_Shown({$mainForm.Activate()})
[void]$mainForm.ShowDialog()
$mainForm.Dispose()

#endregion
