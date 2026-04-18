# Windows Management Toolbox

A menu-driven PowerShell script that automates four common IT tasks on Windows 10/11:

- Software installation via **Winget**
- **WiFi profile** backup and restore
- **Device driver** export and import
- **Windows activation** using the BIOS OEM key

---

## Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 or Windows 11 |
| PowerShell | 5.1 or 7+ |
| Privileges | **Administrator** (mandatory) |
| Winget | Auto-installed/updated by the script |
| Internet | Required for Winget installs and OSD driver downloads |

---

## Files

```
c:\Winget\
├── Windows-Toolbox.ps1            ← Main script (run this)
├── Winget-Install-Update_v1.4.ps1 ← Winget prereq installer (called automatically)
├── WiFi-Profiles\                 ← WiFi XML exports are saved here
├── CLAUDE.md                      ← Project coding standards (for developers)
├── SKILL.md                       ← Technical command reference (for developers)
└── README.md                      ← This file
```

---

## How to Run

Right-click **PowerShell** → **Run as Administrator**, then:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Winget\Windows-Toolbox.ps1"
```

Or navigate to the folder first:

```powershell
cd C:\Winget
.\Windows-Toolbox.ps1
```

---

## Menu Options

```
[1] Install / Update Software (Winget)
[2] Edit Winget Package List
[3] Backup WiFi Profiles
[4] Restore WiFi Profiles
[5] Export Drivers  (Native)
[6] Import Drivers  (Native)
[7] Export Drivers  (OSD Module)
[8] Activate Windows (BIOS OEM Key)
[9] Check Activation Status
[0] Exit
```

---

## Managing the Software List

The list of apps installed by option **[1]** is defined in the `$WingetPackages` array near the top of `Windows-Toolbox.ps1`.

**Option A — Use the built-in editor (recommended):**
1. Run the script and choose **[2] Edit Winget Package List**
2. Notepad opens the script directly
3. Find the `$WingetPackages` array and add or remove IDs
4. Save the file, then run option **[1]**

**Option B — Edit manually:**
Open `Windows-Toolbox.ps1` in any text editor and modify the array:

```powershell
$WingetPackages = @(
    "7zip.7zip",
    "Brave.Brave",
    # Add your IDs here:
    "VideoLAN.VLC",
    "Spotify.Spotify"
)
```

**Finding a Winget ID:**

```powershell
winget search <AppName>
# Example:
winget search vlc
```

---

## Pre-loaded Software (22 packages)

| Application | Winget ID |
|---|---|
| 7-Zip | `7zip.7zip` |
| AMD Chipset Software | `AMD.AmdChipsetSoftware` |
| AMD Software: Adrenalin | `AMD.AmdSoftwareAdrenalin` |
| Brave Browser | `Brave.Brave` |
| Git | `Git.Git` |
| GitHub CLI | `GitHub.cli` |
| LM Studio | `LMStudio.LMStudio` |
| Logi Options+ | `Logitech.LogiOptionsPlus` |
| .NET SDK 10 | `Microsoft.DotNet.SDK.10` |
| ASP.NET Core Runtime 10 | `Microsoft.DotNet.AspNetCore.10` |
| Microsoft OneDrive | `Microsoft.OneDrive` |
| Microsoft Teams | `Microsoft.Teams` |
| Visual C++ Redist x64 | `Microsoft.VCRedist.2015+.x64` |
| Visual C++ Redist x86 | `Microsoft.VCRedist.2015+.x86` |
| Visual Studio Code | `Microsoft.VisualStudioCode` |
| Windows Desktop Runtime 10 | `Microsoft.DotNet.DesktopRuntime.10` |
| Node.js LTS | `OpenJS.NodeJS.LTS` |
| PowerShell 7 | `Microsoft.PowerShell` |
| Proton Drive | `Proton.ProtonDrive` |
| Proton Pass | `Proton.ProtonPass` |
| Python 3 | `Python.Python.3` |
| Vulkan SDK | `KhronosGroup.VulkanSDK` |

> **Note:** Microsoft Office Professional Plus 2021 is a volume license product and cannot be installed via Winget. Install it separately using your volume license media.

---

## WiFi Profiles

- **Backup [3]:** Exports all saved WiFi profiles as XML files to `WiFi-Profiles\` in the script folder. Passwords are included in plain text — store the folder securely.
- **Restore [4]:** Imports all `*.xml` files from `WiFi-Profiles\`. Profiles that already exist are skipped automatically (no errors).

To restore from a custom folder, edit the `$script:WiFiFolder` variable at the top of the script.

---

## Driver Export & Import

### Option [5] — Native Export (`Export-WindowsDriver`)
Exports all currently installed third-party drivers to a folder. You will be prompted for a destination path (default: `C:\Drivers\<ComputerName>`).

### Option [6] — Native Import (`Add-WindowsDriver`)
Installs drivers from a previously exported folder. You will be prompted for the source path. **Note:** This targets the offline Windows image. For installing into the running OS, use Device Manager or `pnputil` after export.

### Option [7] — OSD Module (`Save-MyDriverPack`)
Downloads the official driver pack for your device from the manufacturer's catalog. Requires internet. If the OSD module is not installed, the script will ask permission before downloading it from PSGallery.

---

## Windows Activation

### Option [8] — Activate via BIOS OEM Key
Reads the embedded OEM product key from the BIOS/UEFI firmware and activates Windows online. Works on most OEM devices (HP, Dell, Lenovo, etc.).

- If Windows is already activated, the script will report that and take no action.
- If no BIOS key is found (e.g., custom-built PCs), the script reports an error and stops.

### Option [9] — Check Activation Status
Displays the current license status, product name, and partial product key.

| Status Code | Meaning |
|---|---|
| 1 | Activated |
| 2 | Out-of-Box Grace Period |
| 5 | Not Activated |

---

## Log File

Every operation is logged to:

```
%TEMP%\Windows-Toolbox_YYYYMMDD.log
```

Typical path: `C:\Users\<YourName>\AppData\Local\Temp\Windows-Toolbox_20260401.log`

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Must be run as Administrator" | Right-click PowerShell → Run as Administrator |
| Winget not found | Run option [1] — the prereq script installs it automatically |
| WiFi restore says "already exists" | Normal — the profile is already saved. No action needed. |
| BIOS key not found | Device may not have an embedded OEM key. Use a retail key with `slmgr /ipk <key>` |
| OSD module install fails | Check internet connection and PSGallery availability |
| Driver import has no effect | Native `Add-WindowsDriver` targets offline images. Use `pnputil` for the live OS. |
