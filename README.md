# Audio Slicer with Excel Tracking

A sophisticated Python audio processing tool that automatically slices audio files based on climax points, applies professional audio effects, and maintains an Excel database for tracking.

## Features

- **Automated Audio Slicing**: Extract 30-second audio segments centered on specified climax points
- **Professional Audio Processing**: 
  - Fade in/out (15 seconds each)
  - Audio normalization
  - Format conversion support
- **Excel Integration**: Automatic tracking of all slices with metadata
- **File Verification**: Cross-check between file system and database
- **Batch Processing**: Process multiple slices from a single configuration file

## Project Structure

Pydub audio slicer/
├── slicer.py # Main processing script
├── README.md # This file
├── .gitignore # Git exclusion rules
├── raw audio/ # Input files
│ ├── audio.wav # Source audio file
│ └── audio.txt # Slice definitions
└── blocks/ # Output directory
├── m1.wav, v1.wav... # Generated audio slices
└── blocks_list.xlsx # Auto-generated tracking database

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/xholcaravan/pydub-audio-slicer-sequencer.git
   cd pydub-audio-slicer-sequencer

2. **Set up virtual environment**:
python3 -m venv venv
source venv/bin/activate

3. **Install dependencies**:
pip install pydub pandas openpyxl

4. **Install system dependencies:**:
sudo apt install ffmpeg


Usage
1. Prepare Input Files

Place your source audio file in raw audio/audio.wav

Create raw audio/audio.txt with slice definitions:

45    45    v voice description here
75    75    m music description here
120   120   v another voice segment

Format: climax_time ignore_time type description

    climax_time: Center point for the 30-second slice (seconds)

    ignore_time: Second column is ignored

    type: v for voice, m for music

    description: Free text description

2. Run the Slicer
python3 slicer.py

3. Output

The script will:

    Create audio slices in blocks/ folder (m1.wav, v1.wav, m2.wav, etc.)

    Generate/update blocks/blocks_list.xlsx with tracking information

    Apply fade in/out and normalization

    Verify file system vs database consistency

Excel Database Structure

Sheet "m" (Music):
m	origin	description
m1	/path/to/audio.wav	music description

Sheet "v" (Voice):
v	origin	description
v1	/path/to/audio.wav	voice description

Configuration

Edit SLICE_SIZE and FADE_DURATION in slicer.py to customize:

SLICE_SIZE = 30        # Seconds
FADE_DURATION = 15     # Seconds (half of SLICE_SIZE)


Dependencies

    pydub: Audio processing

    pandas: Excel file handling

    openpyxl: Excel file support

    ffmpeg: Audio format support (system package)

License

Private project - All rights reserved.


