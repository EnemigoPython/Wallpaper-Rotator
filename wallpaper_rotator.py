#!/usr/bin/env python3
"""
Wallpaper Slideshow Rotator
Rotates through wallpapers in a specified folder, designed to work with multiple desktops.

SETUP FOR MULTI-DESKTOP SUPPORT:
To make wallpaper changes apply to ALL virtual desktops in Windows 11, install the VirtualDesktop PowerShell module:

1. Open PowerShell as Administrator
2. Run: Install-Module VirtualDesktop -Scope AllUsers -Force
3. Test: Import-Module VirtualDesktop; Get-Command Set-AllDesktopWallpapers

Without this module, wallpaper changes only apply to the current virtual desktop.
"""

import ctypes
import os
import json
import random
from pathlib import Path
from typing import List, Optional

DEFAULT_WALLPAPER_DIR = os.environ.get("WallpaperDir") or ""


class WallpaperRotator:
    def __init__(self, wallpaper_folder: str, state_file: Optional[str] = None):
        self.wallpaper_folder = Path(wallpaper_folder)
        self.state_file = (
            Path(state_file)
            if state_file
            else self.wallpaper_folder / ".wallpaper_state.json"
        )

        # Supported image formats
        self.supported_formats = {
            ".jpg",
            ".jpeg",
            ".png",
            ".bmp",
            ".gif",
            ".tiff",
            ".webp",
        }

        # Windows API constants
        self.SPI_SETDESKWALLPAPER = 20
        self.SPIF_UPDATEINIFILE = 0x01
        self.SPIF_SENDCHANGE = 0x02

        if not self.wallpaper_folder.exists():
            raise FileNotFoundError(
                f"Wallpaper folder not found: {self.wallpaper_folder}"
            )

    def get_image_files(self) -> List[Path]:
        """Get all supported image files from the wallpaper folder."""
        image_files = []
        for file_path in self.wallpaper_folder.iterdir():
            if (
                file_path.is_file()
                and file_path.suffix.lower() in self.supported_formats
            ):
                image_files.append(file_path)

        # Sort for consistent ordering
        return sorted(image_files)

    def load_state(self) -> dict:
        """Load the current state from the state file."""
        default_state = {
            "current_index": -1,
            "last_wallpaper": None,
            "image_count": 0,
            "order": "sequential",  # or "random"
        }

        if not self.state_file.exists():
            return default_state

        try:
            with open(self.state_file, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return default_state

    def save_state(self, state: dict) -> None:
        """Save the current state to the state file."""
        try:
            with open(self.state_file, "w") as f:
                json.dump(state, f, indent=2)
        except IOError as e:
            print(f"Warning: Could not save state file: {e}")

    def set_wallpaper(self, image_path: Path) -> bool:
        """Set the desktop wallpaper on all virtual desktops using PowerShell."""
        try:
            # Convert to absolute path string
            wallpaper_path = str(image_path.resolve())

            # First try PowerShell VirtualDesktop module for all desktops
            if self._set_wallpaper_all_desktops_powershell(wallpaper_path):
                return True

            # Fallback to traditional API (current desktop only)
            print("PowerShell method failed, falling back to single desktop method...")
            result = ctypes.windll.user32.SystemParametersInfoW(
                self.SPI_SETDESKWALLPAPER,
                0,
                wallpaper_path,
                self.SPIF_UPDATEINIFILE | self.SPIF_SENDCHANGE,
            )

            return result != 0
        except Exception as e:
            print(f"Error setting wallpaper: {e}")
            return False

    def _set_wallpaper_all_desktops_powershell(self, wallpaper_path: str) -> bool:
        """Set wallpaper on all virtual desktops using PowerShell VirtualDesktop module."""
        try:
            import subprocess

            # PowerShell command to set wallpaper on all virtual desktops
            ps_command = f'''
            # Import VirtualDesktop module if available
            if (Get-Module -ListAvailable -Name VirtualDesktop) {{
                Import-Module VirtualDesktop -ErrorAction SilentlyContinue
                if (Get-Command Set-AllDesktopWallpapers -ErrorAction SilentlyContinue) {{
                    Set-AllDesktopWallpapers -Path "{wallpaper_path}"
                    Write-Output "SUCCESS"
                }} else {{
                    Write-Output "COMMAND_NOT_FOUND"
                }}
            }} else {{
                Write-Output "MODULE_NOT_FOUND"
            }}
            '''

            # Execute PowerShell command hidden (no window flash)
            result = subprocess.run(
                ["powershell", "-WindowStyle", "Hidden", "-Command", ps_command],
                capture_output=True,
                text=True,
                timeout=30,
                creationflags=subprocess.CREATE_NO_WINDOW
                if hasattr(subprocess, "CREATE_NO_WINDOW")
                else 0,
            )

            if result.returncode == 0 and "SUCCESS" in result.stdout:
                return True
            elif "MODULE_NOT_FOUND" in result.stdout:
                print(
                    "VirtualDesktop PowerShell module not installed. Install with: Install-Module VirtualDesktop"
                )
                return False
            elif "COMMAND_NOT_FOUND" in result.stdout:
                print(
                    "Set-AllDesktopWallpapers command not found. Update VirtualDesktop module."
                )
                return False
            else:
                print(f"PowerShell error: {result.stderr}")
                return False

        except subprocess.TimeoutExpired:
            print("PowerShell command timed out")
            return False
        except Exception as e:
            print(f"PowerShell execution error: {e}")
            return False

    def get_next_wallpaper(self, order: str = "sequential") -> Optional[Path]:
        """Get the next wallpaper in the rotation."""
        image_files = self.get_image_files()

        if not image_files:
            print("No image files found in wallpaper folder.")
            return None

        state = self.load_state()
        current_index = state.get("current_index", -1)

        # Check if the folder contents have changed
        if len(image_files) != state.get("image_count", 0):
            print(f"Folder contents changed. Found {len(image_files)} images.")
            current_index = -1  # Reset to start from beginning

        if order == "random":
            # Random selection (avoid repeating the same image immediately)
            if len(image_files) > 1:
                available_indices = [
                    i for i in range(len(image_files)) if i != current_index
                ]
                next_index = random.choice(available_indices)
            else:
                next_index = 0
        else:
            # Sequential order
            next_index = (current_index + 1) % len(image_files)

        # Update state
        next_wallpaper = image_files[next_index]
        state.update(
            {
                "current_index": next_index,
                "last_wallpaper": str(next_wallpaper),
                "image_count": len(image_files),
                "order": order,
            }
        )

        self.save_state(state)
        return next_wallpaper

    def rotate_wallpaper(self, order: str = "sequential", verbose: bool = True) -> bool:
        """Rotate to the next wallpaper."""
        next_wallpaper = self.get_next_wallpaper(order)

        if next_wallpaper is None:
            return False

        if verbose:
            print(f"Setting wallpaper to: {next_wallpaper.name}")

        success = self.set_wallpaper(next_wallpaper)

        if success and verbose:
            print("Wallpaper changed successfully!")
        elif verbose:
            print("Failed to change wallpaper.")

        return success

    def set_order(self, order: str) -> None:
        """Set the rotation order (sequential or random)."""
        if order not in ["sequential", "random"]:
            raise ValueError("Order must be 'sequential' or 'random'")

        state = self.load_state()
        state["order"] = order
        self.save_state(state)

    def reset_rotation(self) -> None:
        """Reset the rotation to start from the beginning."""
        state = self.load_state()
        state["current_index"] = -1
        self.save_state(state)
        print("Rotation reset to start from beginning.")

    def check_virtualdesktop_support(self) -> bool:
        """Check if VirtualDesktop module and commands are available."""
        try:
            import subprocess

            ps_command = """
            if (Get-Module -ListAvailable -Name VirtualDesktop) {
                Import-Module VirtualDesktop -ErrorAction SilentlyContinue
                if (Get-Command Set-AllDesktopWallpapers -ErrorAction SilentlyContinue) {
                    Write-Output "SUPPORTED"
                } else {
                    Write-Output "COMMAND_MISSING"
                }
            } else {
                Write-Output "MODULE_MISSING"
            }
            """

            result = subprocess.run(
                ["powershell", "-WindowStyle", "Hidden", "-Command", ps_command],
                capture_output=True,
                text=True,
                timeout=30,
                creationflags=subprocess.CREATE_NO_WINDOW
                if hasattr(subprocess, "CREATE_NO_WINDOW")
                else 0,
            )

            return result.returncode == 0 and "SUPPORTED" in result.stdout

        except Exception:
            return False

    def get_status(self) -> dict:
        """Get current status information."""
        image_files = self.get_image_files()
        state = self.load_state()

        current_wallpaper = None
        if 0 <= state.get("current_index", -1) < len(image_files):
            current_wallpaper = image_files[state["current_index"]].name

        return {
            "total_images": len(image_files),
            "current_index": state.get("current_index", -1),
            "current_wallpaper": current_wallpaper,
            "order": state.get("order", "sequential"),
            "wallpaper_folder": str(self.wallpaper_folder),
        }


def main():
    """Main function for command-line usage."""
    import argparse

    parser = argparse.ArgumentParser(description="Rotate desktop wallpapers")
    parser.add_argument(
        "--folder",
        "-f",
        default=rf"{DEFAULT_WALLPAPER_DIR}",
        help="Path to wallpaper folder",
    )
    parser.add_argument(
        "--order",
        "-o",
        choices=["sequential", "random"],
        default="sequential",
        help="Rotation order",
    )
    parser.add_argument(
        "--reset", "-r", action="store_true", help="Reset rotation to beginning"
    )
    parser.add_argument(
        "--status", "-s", action="store_true", help="Show current status"
    )
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Install VirtualDesktop PowerShell module for multi-desktop support",
    )
    parser.add_argument(
        "--check-support",
        action="store_true",
        help="Check if multi-desktop support is available",
    )
    parser.add_argument(
        "--quiet", "-q", action="store_true", help="Quiet mode - minimal output"
    )

    args = parser.parse_args()

    try:
        rotator = WallpaperRotator(args.folder)

        if args.check_support:
            if rotator.check_virtualdesktop_support():
                print("✓ Multi-desktop support is available")
            else:
                print("✗ Multi-desktop support not available")
                print("Run with --setup to install the required PowerShell module")
            return

        if args.status:
            status = rotator.get_status()
            print("Wallpaper Rotator Status:")
            print(f"  Folder: {status['wallpaper_folder']}")
            print(f"  Total images: {status['total_images']}")
            print(f"  Current index: {status['current_index']}")
            print(f"  Current wallpaper: {status['current_wallpaper']}")
            print(f"  Order: {status['order']}")

            # Check virtual desktop support
            if rotator.check_virtualdesktop_support():
                print("  Multi-desktop support: ✓ Available")
            else:
                print(
                    "  Multi-desktop support: ✗ Not available (run --setup to install)"
                )
            return

        if args.reset:
            rotator.reset_rotation()
            return

        # Check support and warn user if needed
        if not args.quiet and not rotator.check_virtualdesktop_support():
            print("⚠️  WARNING: Multi-desktop support not available.")
            print("   Wallpaper will only change on the current virtual desktop.")
            print(
                "   Run 'python wallpaper_rotator.py --setup' to enable support for all desktops."
            )
            print()

        # Set order if specified
        rotator.set_order(args.order)

        # Rotate wallpaper
        success = rotator.rotate_wallpaper(args.order, verbose=not args.quiet)

        if not success:
            exit(1)

    except Exception as e:
        print(f"Error: {e}")
        exit(1)


if __name__ == "__main__":
    main()
