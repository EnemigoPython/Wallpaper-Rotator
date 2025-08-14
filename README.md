# Windows 11 Multi-Desktop Wallpaper Rotator

A Python-based wallpaper slideshow solution that actually works with Windows 11 virtual desktops.

## The Problem

Windows 11's built-in wallpaper slideshow feature is broken when using multiple virtual desktops. While Microsoft advertises the ability to set different wallpapers for each virtual desktop, enabling the slideshow functionality silently reverts all desktops to a single static image, completely breaking the multi-desktop wallpaper experience.

This is a known issue with no official fix from Microsoft, frustrating users who use virtual desktops and want their wallpapers to rotate consistently across all workspaces.

## The Solution

This project provides a custom wallpaper rotation system that:

- ✅ **Works across ALL virtual desktops simultaneously**
- ✅ **Runs completely silently in the background** (no window flashing)
- ✅ **Maintains rotation state** between system restarts
- ✅ **Supports both sequential and random rotation**
- ✅ **Handles dynamic folder changes** (add/remove images anytime)
- ✅ **Integrates with Windows Task Scheduler** for automated execution
- ✅ **Provides status monitoring and control options**

## How It Works

The solution combines multiple technologies to overcome Windows limitations:

1. **Python**: Handles the core logic (file management, state persistence, rotation algorithms)
2. **PowerShell VirtualDesktop Module**: Enables wallpaper changes across all virtual desktops
3. **VBScript**: Provides truly silent execution when run via Task Scheduler
4. **Windows Task Scheduler**: Automates the rotation at your preferred intervals

## Dependencies

### Required
- **Windows 11** (or Windows 10 with virtual desktop support)
- **Python 3.6+** with `pythonw.exe` available
- **PowerShell 5.1+**

### PowerShell Module (Required for Multi-Desktop Support)
The script requires the VirtualDesktop PowerShell module to work across all virtual desktops:

```powershell
# Run PowerShell as Administrator
Install-Module VirtualDesktop -Scope AllUsers -Force
```

Without this module, the script will fall back to single-desktop mode (current desktop only).

## Installation & Setup

### 1. Download the Files
Download these files to a folder (e.g., `C:\WallpaperRotator\`):
- `wallpaper_rotator.py`
- `run_wallpaper_rotator.vbs`

### 2. Install PowerShell Module
```powershell
# Open PowerShell as Administrator
Install-Module VirtualDesktop -Scope AllUsers -Force

# Test installation
Import-Module VirtualDesktop
Get-Command Set-AllDesktopWallpapers
```

### 3. Configure Your Wallpaper Folder
Edit the default folder path in `wallpaper_rotator.py` (line 196) or use the `--folder` parameter:

```python
default=r"C:\Users\YourName\Pictures\Wallpapers"
```

### 4. Test the Script
```bash
# Test basic functionality
python wallpaper_rotator.py --status

# Test multi-desktop support
python wallpaper_rotator.py --check-support

# Rotate wallpaper once
python wallpaper_rotator.py
```

### 5. Set Up Automated Rotation

#### Option A: Task Scheduler with VBScript (Recommended - Silent)
1. Open Task Scheduler (`taskschd.msc`)
2. Create Basic Task
3. **Name**: "Wallpaper Rotator"
4. **Trigger**: Daily, repeat every 5 minutes
5. **Action**: Start a program
6. **Program**: `C:\WallpaperRotator\run_wallpaper_rotator.vbs`
7. **Settings**: Enable "Run whether user is logged on or not"

#### Option B: Task Scheduler with Python Directly
1. **Program**: `pythonw.exe`
2. **Arguments**: `"C:\WallpaperRotator\wallpaper_rotator.py" --quiet`
3. **Start in**: `C:\WallpaperRotator\`

## Usage

### Command Line Options
```bash
# Basic rotation
python wallpaper_rotator.py

# Custom folder
python wallpaper_rotator.py --folder "C:\MyWallpapers"

# Random order instead of sequential
python wallpaper_rotator.py --order random

# Check status
python wallpaper_rotator.py --status

# Reset rotation to beginning
python wallpaper_rotator.py --reset

# Quiet mode (no output)
python wallpaper_rotator.py --quiet

# Check multi-desktop support
python wallpaper_rotator.py --check-support
```

### Supported Image Formats
- `.jpg`, `.jpeg`
- `.png`
- `.bmp`
- `.gif`
- `.tiff`
- `.webp`

## Troubleshooting

### "Multi-desktop support not available"
Install the VirtualDesktop PowerShell module:
```powershell
Install-Module VirtualDesktop -Scope AllUsers -Force
```

### Task Scheduler shows windows briefly
Use the VBScript wrapper (`run_wallpaper_rotator.vbs`) instead of calling Python directly.

### Wallpaper not changing
1. Check that images exist in your folder
2. Verify Python can access the folder path
3. Run with `--status` to see current state
4. Check Task Scheduler execution history

### Permission errors
- Run PowerShell as Administrator when installing modules
- Ensure the wallpaper folder is accessible to your user account

## Technical Notes

- **State persistence**: Rotation state is stored in `.wallpaper_state.json` in your wallpaper folder
- **Multi-desktop detection**: Automatically detects and uses VirtualDesktop module when available
- **Graceful fallback**: Works on single desktop if multi-desktop support isn't available
- **Error handling**: Continues working even if individual images can't be loaded

## Credits

**Created by**: [Basil Eagle](https://github.com/basileagle)

**AI Development Partner**: Claude (Anthropic) - Contributed significantly to the technical implementation, including:
- Solving the Windows 11 virtual desktop compatibility issues
- Implementing the PowerShell VirtualDesktop module integration
- Creating the VBScript wrapper for silent Task Scheduler execution
- Debugging subprocess window visibility problems
- Architecture design for multi-technology integration
- Writing this very README (because apparently I write documentation now too)

## License

MIT License - Feel free to use, modify, and distribute.

## Contributing

Issues and pull requests welcome! This project solves a real Windows 11 limitation, and improvements benefit everyone frustrated by Microsoft's broken slideshow feature.

---

*Finally, a wallpaper slideshow that works with Windows 11 virtual desktops!* 🎨