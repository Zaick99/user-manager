
<img width="1536" height="1024" alt="user-manager" src="https://github.com/user-attachments/assets/23de586d-6c36-495f-85a6-ee54243225ae" />

# 🖥️ Windows Domain Profile Manager

A powerful PowerShell GUI tool for managing user profiles on Windows domain computers with advanced filtering, bulk operations, and Active Directory integration.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)
![Windows](https://img.shields.io/badge/Windows-10%2F11%2FServer-green.svg)
![License](https://img.shields.io/badge/license-MIT-orange.svg)

---

## ✨ Key Features

### 🔍 **Profile Analysis**
- **Smart Detection**: Automatically identifies orphaned, empty, and inactive profiles
- **Size Calculation**: Real-time calculation of Desktop, Documents, and Downloads folder sizes
- **AD Integration**: Fetches user display names from Active Directory
- **Registry Validation**: Cross-references profiles with Windows registry entries
- **Last Logon Tracking**: Shows the last login date for each profile

### 🎯 **Intelligent Filtering**
- **Select Orphaned**: Profiles with missing registry keys or temporary suffixes (.temp, .bak, .old)
- **Select Empty**: Domain profiles with no files in main folders
- **Select Old**: Profiles not accessed in the last 90 days
- **Manual Selection**: Check/uncheck individual profiles with visual feedback

### 🚀 **Bulk Operations**
- **Parallel Deletion**: Multi-threaded profile removal for maximum speed
- **Safety Locks**: System accounts are automatically protected
- **Progress Tracking**: Real-time deletion progress with detailed feedback
- **Error Handling**: Robust error reporting for failed operations

### 🎨 **Modern UI**
- **Windows 11 Style**: Dark theme with rounded corners and modern aesthetics
- **Color-Coded Indicators**: Visual feedback for profile status
- **Responsive Design**: Adapts to different screen sizes
- **Live Updates**: Background loading of profile data without UI freezing

---

## 🛠️ Requirements

### **Mandatory**
- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- Administrator privileges

### **Optional (for full functionality)**
- Active Directory PowerShell module (RSAT tools)
- Domain membership (for AD user lookups)
- Remote registry access (for remote computer management)

---

## 📥 Installation

1. **Download the script**
   ```powershell
   # Clone or download user-manager.ps1
   ```

2. **Unblock the file** (if downloaded from internet)
   ```powershell
   Unblock-File -Path .\user-manager.ps1
   ```

3. **Install RSAT tools** (optional, for AD integration)
   ```powershell
   # Windows 10/11
   Get-WindowsCapability -Online | Where-Object Name -like "Rsat.ActiveDirectory*" | Add-WindowsCapability -Online
   
   # Windows Server
   Install-WindowsFeature -Name RSAT-AD-PowerShell
   ```

---

## 🚀 Usage

### **Local Computer**
Run the script without parameters to manage the local machine:
```powershell
.\user-manager.ps1
```

### **Remote Computer**
Specify a remote computer name:
```powershell
.\user-manager.ps1 -ComputerName "PC-NAME"
```

### **Advanced Usage**
```powershell
# Manage remote computer with explicit credentials (use PsExec or remote PowerShell)
Enter-PSSession -ComputerName "PC-NAME" -Credential (Get-Credential)
.\user-manager.ps1
```

---

## 🎮 How to Use

### **Main Window Overview**

<img width="1386" height="793" alt="user-manager-preview" src="https://github.com/user-attachments/assets/4cee1dd3-9719-41ee-9854-77775d558144" />

### **Step-by-Step Guide**

#### 1️⃣ **Launch and Load**
- Run the script (local or remote mode)
- The window opens with a loading animation
- Profile data loads in the background

#### 2️⃣ **Analyze Profiles**
- **Green circles (🟢)**: Active domain profiles with registry keys
- **Yellow circles (🟡)**: Orphaned profiles (no registry key)
- **Red circles (🔴)**: Local accounts (non-domain)
- **Black circles (⚫)**: Empty profiles (no files)

#### 3️⃣ **Use Smart Filters**

**🟡 Select Orphaned**
- Automatically checks all orphaned profiles
- These typically include:
  - Profiles ending in `.temp`, `.bak`, `.old`
  - Profiles with domain suffixes (e.g., `user.DOMAIN000`)
  - Profiles without registry entries

**⚫ Select Empty**
- Checks domain profiles with:
  - No files in Desktop, Documents, or Downloads
  - Valid registry entries (not orphaned)
  - Zero size in main folders

**📅 Select >90gg**
- Checks profiles not accessed in 90+ days
- Based on last logon timestamp
- Excludes local and orphaned profiles

#### 4️⃣ **Manual Selection**
- Click checkboxes to manually select/deselect profiles
- Use **Select All** / **Deselect All** for bulk operations
- System accounts cannot be selected (protected)

#### 5️⃣ **Delete Profiles**
- Click **Delete selected** button
- Review the confirmation dialog
- Watch the progress bar during deletion
- Check the results summary

---

## 🔧 Technical Details

### **Functions Explained**

#### 📊 `Get-FolderSizeMB($path)`
Calculates folder size in megabytes using recursive file enumeration. Returns 0 for inaccessible paths.

#### 👤 `Get-ADUserDisplayName($samAccountName)`
Queries Active Directory for user's display name. Returns empty string if user not found or AD unavailable.

#### 🔍 `Is-OrphanedProfile($folderName)`
Detects orphaned profiles by checking for:
- Temporary suffixes: `.temp`, `.temp0`, `.temp1`, etc.
- Backup suffixes: `.bak`, `.bak0`, etc.
- Old suffixes: `.old`, `.old0`, etc.
- Domain suffixes: `.DOMAIN`, `.DOMAIN000`, etc.
- Numeric suffixes: `.000001`, etc.

#### 🏠 `Is-LocalAccount($folderName)`
Determines if a profile belongs to a local account:
1. Checks against system account list
2. Queries Active Directory
3. Returns `true` if user not found in AD

#### 📅 `Get-ProfileLastLogon($sid, $folderPath)`
Retrieves last logon time using:
- **Registry method**: Reads `LocalProfileLoadTime` from registry (local mode)
- **File method**: Uses NTUSER.DAT last write time (remote mode)
- Returns `$null` if unable to determine

#### 🌐 `Get-RemoteRegistry()`
Connects to remote or local registry to:
- Enumerate profile list subkeys
- Map ProfileImagePath to SIDs
- Build profile-to-SID associations
- Returns hashtable with Paths and SIDs

#### 🔄 `Convert-ToLocalPath($uncPath)`
Converts UNC paths to local paths:
- `\\PC-NAME\C$\Users\john` → `C:\Users\john`
- Used for registry operations and display

#### 📋 `Get-BasicProfileInfo()`
Main data collection function that:
1. Scans user folders
2. Queries registry
3. Checks Active Directory
4. Detects orphaned profiles
5. Returns profile objects with metadata

#### ⏱️ `Load-ProfileSizesBackground($profileList)`
Asynchronous background worker that:
- Calculates folder sizes without blocking UI
- Updates DataGridView rows progressively
- Uses separate runspace for parallel execution

#### 🔄 `Refresh-Grid()`
Reloads all profile data:
- Clears existing data
- Fetches fresh profile information
- Restarts background size calculation
- Updates profile counter

#### 🎨 `New-Win11Button()`
Creates modern Windows 11-style buttons with:
- Flat design with hover effects
- Custom colors and rounded corners
- Consistent sizing and spacing

### **Color Indicators**

| Color | Meaning | Condition |
|-------|---------|-----------|
| 🟢 Green | Active domain profile | Has registry key, not orphaned, not local |
| 🟡 Yellow | Orphaned profile | No registry key or temporary suffix |
| 🔴 Red | Local account | Not in Active Directory |
| ⚫ Black | Empty profile | No files in main folders |

### **Deletion Process**

The script uses parallel runspaces for fast deletion:

1. **Registry Cleanup**
   - Removes profile key from `ProfileList` registry
   - Uses remote or local registry API

2. **Folder Removal**
   - Attempts quick deletion with `rd /s /q`
   - If failed, takes ownership with `takeown`
   - Grants full permissions with `icacls`
   - Retries deletion

3. **Error Handling**
   - Tracks individual profile errors
   - Reports success/failure counts
   - Shows detailed error messages

---

## ⚠️ Important Notes

### **🔒 Safety Features**
- **System Account Protection**: Administrator, System, Default, Public, and other system accounts cannot be deleted
- **Confirmation Dialog**: Warns before deletion with profile count
- **Irreversible Warning**: Clearly states that deletion cannot be undone

### **🌐 Remote Computer Requirements**
- **Admin Share Access**: Requires `\\COMPUTER\C$` access
- **Remote Registry**: Service must be running
- **Firewall Rules**: Allow File and Printer Sharing
- **Permissions**: Administrator rights on target computer

### **💡 Best Practices**
1. **Test First**: Run on a test machine before production
2. **Backup Important Data**: Ensure critical files are backed up
3. **Check Twice**: Review selected profiles before deletion
4. **Use Filters Carefully**: Understand what each filter selects
5. **Monitor Results**: Check the deletion summary for errors

### **❌ Limitations**
- Profiles currently logged in cannot be deleted
- Some profiles may require manual cleanup if files are locked
- AD lookups require network connectivity and RSAT tools
- Remote operations require appropriate network permissions

---

## 🐛 Troubleshooting

### **Issue: "Access Denied" errors**
**Solution:** Ensure you're running PowerShell as Administrator and have permissions on the target computer.

### **Issue: Active Directory users show as local**
**Solution:** Install RSAT tools and ensure domain connectivity:
```powershell
Import-Module ActiveDirectory
```

### **Issue: Remote computer not accessible**
**Solution:** Check network connectivity and enable remote registry:
```powershell
Set-Service -Name RemoteRegistry -StartupType Automatic -Status Running
```

### **Issue: Profile deletion fails**
**Solution:** 
- Check if user is currently logged in
- Verify file locks with Sysinternals tools
- Manually take ownership if needed

### **Issue: Size calculation is slow**
**Solution:** This is normal for large profiles. The script loads sizes in the background without freezing the UI.

---

## 📝 Changelog

### **Version 1.0** (Current)
- Initial release
- Dark theme Windows 11 UI
- Smart profile detection
- Parallel deletion support
- Active Directory integration
- Background size calculation
- Multiple selection filters

---

## 🤝 Contributing

Contributions are welcome! If you have suggestions, bug reports, or improvements:

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Open a pull request

---

## 📄 License

This project is licensed under the MIT License - feel free to use, modify, and distribute as needed.

---

## 👨‍💻 Author

**Zaick** - [zaick.net](https://zaick.net)

---

## ⭐ Support

If this tool helped you manage Windows profiles more efficiently, consider:
- ⭐ Starring the repository
- 🐛 Reporting bugs or issues
- 💡 Suggesting new features
- 📢 Sharing with other sysadmins

---

## 🔗 Related Tools

- **PsExec**: For remote PowerShell execution
- **Sysinternals Suite**: For advanced troubleshooting
- **Active Directory Users and Computers**: For user management

---

**Made with ❤️ for System Administrators**
