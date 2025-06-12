"""
Fixed Screenshot Processor - Better slide selection logic
Save as: src/processors/diagram_screenshot_processor.py
"""

import os
import subprocess
import platform
import shutil
import tempfile
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import logging

logger = logging.getLogger(__name__)


class DiagramScreenshotProcessor:
    """
    Handles LibreOffice-based screenshot generation with improved slide selection.
    """

    def __init__(self):
        self.libreoffice_path = self._detect_libreoffice()

    def _detect_libreoffice(self) -> Optional[str]:
        """Detect LibreOffice installation across platforms."""
        system = platform.system()

        possible_paths = []

        if system == "Darwin":  # macOS
            possible_paths = [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "/opt/homebrew/bin/soffice",
                "/usr/local/bin/soffice"
            ]
        elif system == "Windows":
            possible_paths = [
                "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
                "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
            ]
        else:  # Linux
            possible_paths = [
                "/usr/bin/soffice",
                "/usr/local/bin/soffice"
            ]

        # Check each possible path
        for path in possible_paths:
            if os.path.exists(path):
                logger.info(f"Found LibreOffice at: {path}")
                return path

        # Try to find in PATH
        try:
            result = subprocess.run(
                ["which", "soffice"] if system != "Windows" else ["where", "soffice"],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip().split('\n')[0]
                logger.info(f"Found LibreOffice in PATH: {path}")
                return path
        except:
            pass

        return None

    def is_available(self) -> bool:
        """Check if LibreOffice is available."""
        return self.libreoffice_path is not None

    def _screenshot_specific_slides(self,
                                    pptx_path: str,
                                    slide_numbers: List[int],
                                    output_dir: str,
                                    base_filename: str,
                                    debug_mode: bool = True) -> Dict[int, str]:
        """
        Screenshot specific slides using LibreOffice with better slide selection.

        Args:
            pptx_path (str): Path to PowerPoint file
            slide_numbers (List[int]): Slide numbers to screenshot (1-based)
            output_dir (str): Directory to save screenshots
            base_filename (str): Base filename for screenshots
            debug_mode (bool): Show debug information

        Returns:
            Dict[int, str]: Mapping of slide numbers to screenshot file paths
        """
        if not self.is_available():
            raise RuntimeError("LibreOffice not available")

        os.makedirs(output_dir, exist_ok=True)
        results = {}

        with tempfile.TemporaryDirectory() as temp_dir:
            if debug_mode:
                print(f"🔧 Using temp directory: {temp_dir}")
                print(f"🎯 Requested slides: {slide_numbers}")

            # Export all slides to images using LibreOffice
            cmd = [
                self.libreoffice_path,
                "--headless",
                "--convert-to", "png",
                "--outdir", temp_dir,
                pptx_path
            ]

            if debug_mode:
                print(f"🚀 Running LibreOffice: {' '.join(cmd)}")

            try:
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=120  # 2 minute timeout
                )

                if result.returncode != 0:
                    error_msg = f"LibreOffice export failed: {result.stderr}"
                    if debug_mode:
                        print(f"❌ {error_msg}")
                    raise RuntimeError(error_msg)

                if debug_mode:
                    print("✅ LibreOffice export completed successfully")

                # List all exported files
                exported_files = [f for f in os.listdir(temp_dir) if f.endswith('.png')]
                exported_files.sort()  # Sort for consistent ordering

                if debug_mode:
                    print(f"📁 Found {len(exported_files)} exported files:")
                    for i, f in enumerate(exported_files, 1):
                        print(f"  {i:2d}: {f}")

                if not exported_files:
                    raise RuntimeError("No PNG files were exported by LibreOffice")

                # Map slide numbers to files using multiple strategies
                results = self._map_slides_to_files(
                    slide_numbers, exported_files, temp_dir,
                    pptx_path, output_dir, base_filename, debug_mode
                )

            except subprocess.TimeoutExpired:
                raise RuntimeError("LibreOffice export timed out (>2 minutes)")
            except Exception as e:
                raise RuntimeError(f"Error during LibreOffice export: {str(e)}")

        return results

    def _map_slides_to_files(self, slide_numbers: List[int], exported_files: List[str],
                             temp_dir: str, pptx_path: str, output_dir: str,
                             base_filename: str, debug_mode: bool) -> Dict[int, str]:
        """Map requested slide numbers to exported files using multiple strategies."""
        results = {}
        pptx_stem = Path(pptx_path).stem

        if debug_mode:
            print(f"\n🔍 Mapping slides to files...")
            print(f"📊 PowerPoint stem: '{pptx_stem}'")

        # Strategy 1: Direct filename matching (most common)
        for slide_num in slide_numbers:
            found_file = None

            # Try multiple naming patterns LibreOffice might use
            patterns_to_try = [
                f"{pptx_stem}_{slide_num:02d}.png",  # presentation_10.png
                f"{pptx_stem}-{slide_num:02d}.png",  # presentation-10.png
                f"{pptx_stem}_{slide_num}.png",  # presentation_10.png (no padding)
                f"{pptx_stem}-{slide_num}.png",  # presentation-10.png (no padding)
                f"{pptx_stem} ({slide_num}).png",  # presentation (10).png
                f"{pptx_stem}_{slide_num:03d}.png",  # presentation_010.png (3-digit padding)
                f"slide{slide_num}.png",  # slide10.png
                f"Slide{slide_num}.png",  # Slide10.png
                f"slide_{slide_num}.png",  # slide_10.png
                f"Slide_{slide_num}.png"  # Slide_10.png
            ]

            if debug_mode:
                print(f"\n🎯 Looking for slide {slide_num}:")

            for pattern in patterns_to_try:
                test_path = os.path.join(temp_dir, pattern)
                if os.path.exists(test_path):
                    found_file = test_path
                    if debug_mode:
                        print(f"  ✅ Found: {pattern}")
                    break
                elif debug_mode:
                    print(f"  ❌ Not found: {pattern}")

            # Strategy 2: If direct matching fails, try positional matching
            if not found_file and len(exported_files) >= slide_num:
                # Assume files are in slide order (1st file = slide 1, etc.)
                positional_file = exported_files[slide_num - 1]  # Convert to 0-based
                test_path = os.path.join(temp_dir, positional_file)

                if os.path.exists(test_path):
                    found_file = test_path
                    if debug_mode:
                        print(f"  ✅ Using positional match: {positional_file} (position {slide_num})")

            # Strategy 3: Single file case
            if not found_file and len(exported_files) == 1 and len(slide_numbers) == 1:
                # If only one file exported and only one slide requested, use it
                single_file = exported_files[0]
                test_path = os.path.join(temp_dir, single_file)
                found_file = test_path
                if debug_mode:
                    print(f"  ✅ Single file match: {single_file}")

            # Copy the found file to final location
            if found_file and os.path.exists(found_file):
                final_filename = f"{base_filename}_slide_{slide_num:02d}_diagram.png"
                final_path = os.path.join(output_dir, final_filename)

                shutil.copy2(found_file, final_path)
                results[slide_num] = final_path

                if debug_mode:
                    file_size = os.path.getsize(final_path) / 1024  # KB
                    print(f"  💾 Saved: {final_filename} ({file_size:.1f} KB)")
            else:
                if debug_mode:
                    print(f"  ❌ Could not find file for slide {slide_num}")

        if debug_mode:
            print(f"\n📊 Final results: {len(results)} of {len(slide_numbers)} slides captured")
            for slide_num in slide_numbers:
                status = "✅" if slide_num in results else "❌"
                print(f"  {status} Slide {slide_num}")

        return results


def test_diagram_screenshot_capability() -> Tuple[bool, str]:
    """Test if diagram screenshot capability is available."""
    processor = DiagramScreenshotProcessor()

    if not processor.is_available():
        return False, "LibreOffice not found - required for diagram screenshots"

    # Test LibreOffice version
    try:
        result = subprocess.run(
            [processor.libreoffice_path, "--version"],
            capture_output=True,
            text=True,
            timeout=10
        )

        if result.returncode == 0:
            version_info = result.stdout.strip()
            return True, f"LibreOffice available: {version_info}"
        else:
            return False, f"LibreOffice found but not working: {result.stderr}"

    except subprocess.TimeoutExpired:
        return False, "LibreOffice version check timed out"
    except Exception as e:
        return False, f"Error testing LibreOffice: {str(e)}"