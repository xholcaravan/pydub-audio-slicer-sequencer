#!/usr/bin/env python3
"""
Build script for Audio Slicer executable
"""

import os
import platform
import subprocess
import sys

def install_dependencies():
    """Install required packages for building"""
    print("📦 Installing dependencies...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

def build_executable():
    """Build executable for current platform"""
    system = platform.system().lower()
    print(f"🔨 Building for {system}...")
    
    # PyInstaller configuration
    cmd = [
        "pyinstaller",
        "--name=AudioSlicer",
        "--onefile",
        "--windowed",  # No console window
        "--icon=icon.ico",  # We'll create this next
        "--add-data=README.md;.",  # Include README
        "slicer.py"
    ]
    
    if system == "darwin":  # macOS
        cmd.extend(["--osx-bundle-identifier", "com.audioslicer.app"])
    
    subprocess.check_call(cmd)
    print(f"✅ Build completed for {system}!")

def create_icon():
    """Create a simple icon file if it doesn't exist"""
    if not os.path.exists("icon.ico"):
        print("🎨 Creating default icon...")
        # We'll use a simple placeholder - you can replace this later
        try:
            from PIL import Image, ImageDraw
            # Create a simple 256x256 icon
            img = Image.new('RGB', (256, 256), color='#4A90E2')
            draw = ImageDraw.Draw(img)
            # Draw a simple waveform icon
            draw.rectangle([80, 100, 176, 156], fill='white')
            for i, height in enumerate([40, 60, 80, 60, 40, 70, 90, 70]):
                x = 90 + i * 12
                draw.rectangle([x, 128 - height//2, x + 8, 128 + height//2], fill='white')
            img.save('icon.ico', format='ICO', sizes=[(256, 256)])
            print("✅ Icon created!")
        except ImportError:
            print("⚠️  PIL not available, skipping icon creation")

def main():
    print("🚀 Audio Slicer Build System")
    print("=============================")
    
    # Create icon
    create_icon()
    
    # Install dependencies
    install_dependencies()
    
    # Build executable
    build_executable()
    
    print("\n🎉 Build process completed!")
    print("📁 Executable should be in the 'dist' folder")

if __name__ == "__main__":
    main()