# 🎵 Audio Slicer with Excel Tracking

![Audio Slicer](https://img.shields.io/badge/Version-1.0.0-blue.svg)
![Python](https://img.shields.io/badge/Python-3.7+-green.svg)
![Platform](https://img.shields.io/badge/Platform-Linux%20%7C%20Windows%20%7C%20macOS-lightgrey.svg)

A sophisticated Python audio processing tool that automatically slices audio files based on climax points, applies professional audio effects, and maintains an Excel database for tracking.

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🎯 **Automated Audio Slicing** | Extract customizable audio segments centered on specified climax points |
| 🎚️ **Professional Audio Processing** | Fade in/out effects, audio normalization, format conversion support |
| 📊 **Excel Integration** | Automatic tracking of all slices with comprehensive metadata |
| 🔍 **File Verification** | Cross-check between file system and database for consistency |
| 🔄 **Batch Processing** | Process multiple slices from a single configuration file |
| ⚙️ **Flexible Configuration** | Customizable slice duration and fade effects |
| 🎲 **Random Generation** | Automatic slice generation with balanced music/voice distribution |
| 🎵 **Smart Sequencing** | Create mixed sequences with offset music/voice channels |

## 🚀 Quick Start (Executable Version)

### For Linux:
1. 📥 Download `AudioSlicer_v1.0_Linux_x64.zip`
2. 📂 Extract and run: `./AudioSlicer`
3. 🔧 Ensure FFmpeg is installed: `sudo apt install ffmpeg`

**No Python installation required!** 🎉

## 🏗️ Project Structure

Pydub Audio Slicer/
├── 🐍 slicer.py # Main processing script
├── 📖 README.md # This file
├── 🙈 .gitignore # Git exclusion rules
├── 🎵 raw_audio/ # Input files directory
│ ├── audio.wav # Source audio file
│ └── audio.txt # Slice definitions
└── 📁 blocks/ # Output directory
├── m1.mp3, v1.mp3... # Generated audio slices
└── 📊 blocks_list.xlsx # Auto-generated tracking database
text


## 📥 Installation (Source Version)

### Prerequisites
- 🐍 **Python 3.7+**
- 🎵 **FFmpeg**

### Setup Steps

1. **Clone the repository:**
   ```bash
   git clone https://github.com/xholcaravan/pydub-audio-slicer-sequencer.git
   cd pydub-audio-slicer-sequencer

Set up virtual environment:
bash

python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

Install Python dependencies:
bash

pip install pydub pandas openpyxl colorama

Install system dependencies:
bash

# Ubuntu/Debian
sudo apt install ffmpeg

# macOS (using Homebrew)
brew install ffmpeg

# Windows (using chocolatey)
choco install ffmpeg

🎮 Usage
1. Prepare Input Files

Source Audio: Place your audio file at raw_audio/audio.wav

Slice Configuration: Create raw_audio/audio.txt with the following format:
text

45    45    v voice description here
75    75    m music description here
120   120   v another voice segment

Format Explanation:

    Column 1: 🕒 Climax time (seconds) - center point for the audio slice

    Column 2: ❌ Ignored (maintained for compatibility)

    Column 3: 🎵 Type (v for voice, m for music)

    Column 4: 📝 Description of the audio segment

2. Run the Application
bash

python3 slicer.py

3. Choose Your Workflow

The application provides three main options:
Option 1: Slice Audio Files

    Extract 30-second segments from climax points

    Apply professional audio processing

    Track everything in Excel database

Option 2: Sequence Existing Blocks

    Create mixed sequences from existing audio blocks

    Music channel starts at 0:00, voice at 0:15

    Random block ordering with timeline generation

Option 3: Complete Automated Workflow

    Option 2→3: Manual slice definition with automatic sequencing

    Option 3→2: Generate random slices and sequence automatically

4. Output Results

The script will generate:

    🎵 Audio Slices: In blocks/ folder (named m1.mp3, v1.mp3, m2.mp3, etc.)

    📊 Tracking Database: blocks/blocks_list.xlsx with comprehensive metadata

    🎚️ Processed Audio: With fade effects and normalization applied

    🔍 Verification Report: File system vs database consistency check

    📄 Timeline Files: For sequenced outputs with block timing information

📊 Excel Database Structure
Music Sheet ("m")
m	origin	description
m1	/path/to/audio.wav	music description
Voice Sheet ("v")
v	origin	description
v1	/path/to/audio.wav	voice description
⚙️ Configuration

Customize the slicing behavior by modifying these constants in slicer.py:
python

SLICE_SIZE = 30        # Total duration of each slice in seconds
FADE_DURATION = 15     # Fade in/out duration in seconds

🎲 Advanced Features
Random Slice Generation

    Automatically generate balanced music/voice slices

    Specify total minutes of content needed

    Intelligent spacing to avoid overlaps

    50/50 distribution between music and voice

Smart Sequencing

    Music and voice channels with 15-second offset

    Automatic block randomization

    Timeline generation with descriptions

    Professional audio mixing

📦 Dependencies
Python Packages

    pydub: Audio processing and manipulation

    pandas: Excel file handling and data management

    openpyxl: Excel file format support

    colorama: Cross-platform colored terminal output

System Requirements

    FFmpeg: Audio format conversion and processing

🐛 Troubleshooting
Common Issues

    FFmpeg not found:
    bash

# Ubuntu/Debian
sudo apt install ffmpeg

# Verify installation
ffmpeg -version

    File not found: Verify that raw_audio/audio.wav and raw_audio/audio.txt exist

    Permission errors: Check write permissions for the blocks/ directory

    Python dependencies: Ensure virtual environment is activated and all packages installed

Getting Help

If you encounter issues:

    Check that all dependencies are properly installed

    Verify your input files are in the correct format

    Ensure the virtual environment is activated when running the script

    Check the console output for specific error messages

🔧 Building Executables

To create standalone executables for distribution:
bash

# Install PyInstaller
pip install pyinstaller

# Build for current platform
pyinstaller --onefile --name AudioSlicer slicer.py

# Executable will be in dist/AudioSlicer

📄 License

Private project - All rights reserved.

    Note: This tool is designed for audio processing workflows where precise timing and metadata tracking are essential. The Excel integration provides a searchable, sortable database of all generated audio segments for easy management and retrieval.

🆘 Support

For issues and questions:

    Check this README and the troubleshooting section

    Verify all prerequisites are met

    Ensure proper file formats and directory structure

Audio Slicer with Excel Tracking - Professional audio processing made simple 🎵