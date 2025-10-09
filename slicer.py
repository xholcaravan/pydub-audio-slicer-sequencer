#!/usr/bin/env python3
"""
Advanced Audio Slicer with Excel Tracking
"""

from pydub import AudioSegment
from pydub.effects import normalize
import pandas as pd
import os
import argparse
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

# Hardcoded parameters
SLICE_SIZE = 30  # seconds
FADE_DURATION = SLICE_SIZE / 2  # seconds

def parse_audio_txt(file_path):
    """Parse the audio.txt file and return list of slices"""
    slices = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                
                parts = line.split('\t')
                if len(parts) < 3:
                    print(f"Warning: Line {line_num} has invalid format: {line}")
                    continue
                
                try:
                    climax_time = float(parts[0])
                    # parts[1] is ignored
                    audio_type = parts[2].split()[0]  # First character: 'v' or 'm'
                    description = ' '.join(parts[2].split()[1:])  # Rest of the text
                    
                    # Calculate slice times
                    slice_begin = climax_time - (SLICE_SIZE / 2)
                    slice_end = climax_time + (SLICE_SIZE / 2)
                    
                    slices.append({
                        'climax_time': climax_time,
                        'type': audio_type,
                        'description': description,
                        'slice_begin': slice_begin,
                        'slice_end': slice_end
                    })
                    
                except (ValueError, IndexError) as e:
                    print(f"Error parsing line {line_num}: {line} - {e}")
                    continue
                    
    except FileNotFoundError:
        print(f"Error: File {file_path} not found")
        return []
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return []
    
    return slices

def get_next_file_numbers(excel_path):
    """Get the next file numbers for m and v types from Excel file"""
    try:
        # Try to read existing Excel file
        m_df = pd.read_excel(excel_path, sheet_name='m')
        v_df = pd.read_excel(excel_path, sheet_name='v')
        
        next_m = len(m_df) + 1 if not m_df.empty else 1
        next_v = len(v_df) + 1 if not v_df.empty else 1
        
    except FileNotFoundError:
        # If file doesn't exist, start from 1
        print("blocks_list.xlsx not found, creating new file...")
        next_m, next_v = 1, 1
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        next_m, next_v = 1, 1
    
    return next_m, next_v

def process_audio_slice(audio, slice_info, output_folder, file_number):
    """Process a single audio slice"""
    try:
        # Convert times to milliseconds
        begin_ms = int(slice_info['slice_begin'] * 1000)
        end_ms = int(slice_info['slice_end'] * 1000)
        
        # Ensure we don't go beyond audio boundaries
        begin_ms = max(0, begin_ms)
        end_ms = min(len(audio), end_ms)
        
        # Extract slice
        slice_audio = audio[begin_ms:end_ms]
        
        # Apply fade in/out (convert seconds to milliseconds)
        fade_duration_ms = int(FADE_DURATION * 1000)
        slice_audio = slice_audio.fade_in(fade_duration_ms).fade_out(fade_duration_ms)
        
        # Normalize audio
        slice_audio = normalize(slice_audio)
        
        # Export file
        filename = f"{slice_info['type']}{file_number}.wav"
        output_path = os.path.join(output_folder, filename)
        slice_audio.export(output_path, format="wav")
        
        print(f"Created: {filename} (from {slice_info['slice_begin']:.1f}s to {slice_info['slice_end']:.1f}s)")
        return output_path
        
    except Exception as e:
        print(f"Error processing slice: {e}")
        return None

def update_excel_file(excel_path, slice_info, file_number, output_path, origin_file):
    """Update the Excel file with new slice information"""
    try:
        # Create DataFrames for new entries with correct column order
        new_data = {
            slice_info['type']: [f"{slice_info['type']}{file_number}"],  # Column A: m1, v1, etc.
            'origin': [origin_file],  # Column B: origin path
            'description': [slice_info['description']]  # Column C: description
        }
        new_df = pd.DataFrame(new_data)
        
        # Try to read existing file or create new one
        try:
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                sheet_name = slice_info['type']  # 'm' or 'v'
                
                # Read existing sheet
                try:
                    existing_df = pd.read_excel(excel_path, sheet_name=sheet_name)
                    # Append new data
                    updated_df = pd.concat([existing_df, new_df], ignore_index=True)
                except:
                    # Sheet doesn't exist, create new
                    updated_df = new_df
                
                # Write back to sheet
                updated_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        except FileNotFoundError:
            # Create new Excel file
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                new_df.to_excel(writer, sheet_name=slice_info['type'], index=False)
                # Create empty sheet for the other type
                other_type = 'v' if slice_info['type'] == 'm' else 'm'
                pd.DataFrame(columns=[other_type, 'origin', 'description']).to_excel(writer, sheet_name=other_type, index=False)
                
        print(f"Updated Excel: {slice_info['type']}{file_number}")
        
    except Exception as e:
        print(f"Error updating Excel file: {e}")

def verify_files_vs_excel(blocks_dir, excel_path):
    """Verify that files in blocks folder match the Excel database"""
    print("\n=== Verifying Files vs Excel Database ===")
    
    try:
        # Read Excel sheets
        m_df = pd.read_excel(excel_path, sheet_name='m')
        v_df = pd.read_excel(excel_path, sheet_name='v')
        
        # Get all files in blocks directory
        all_files = os.listdir(blocks_dir)
        wav_files = [f for f in all_files if f.endswith('.wav')]
        
        # Extract m and v files from folder
        m_files_folder = [f for f in wav_files if f.startswith('m') and f[1:].split('.')[0].isdigit()]
        v_files_folder = [f for f in wav_files if f.startswith('v') and f[1:].split('.')[0].isdigit()]
        
        # Get file names from Excel (remove file extension for comparison)
        m_files_excel = []
        if not m_df.empty and 'm' in m_df.columns:
            m_files_excel = [f"{row['m']}.wav" for _, row in m_df.iterrows() if pd.notna(row['m'])]
        
        v_files_excel = []
        if not v_df.empty and 'v' in v_df.columns:
            v_files_excel = [f"{row['v']}.wav" for _, row in v_df.iterrows() if pd.notna(row['v'])]
        
        # Compare Music files (m)
        print("\n--- Music Files (m) ---")
        m_folder_set = set(m_files_folder)
        m_excel_set = set(m_files_excel)
        
        missing_in_folder = m_excel_set - m_folder_set
        missing_in_excel = m_folder_set - m_excel_set
        
        if not missing_in_folder and not missing_in_excel:
            print("✅ Perfect match! All Excel records have corresponding files")
        else:
            if missing_in_folder:
                print("❌ Files in Excel but missing in folder:")
                for file in sorted(missing_in_folder):
                    print(f"   - {file}")
            if missing_in_excel:
                print("❌ Files in folder but missing in Excel:")
                for file in sorted(missing_in_excel):
                    print(f"   - {file}")
        
        print(f"Total in Excel: {len(m_files_excel)}, Total in folder: {len(m_files_folder)}")
        
        # Compare Voice files (v)
        print("\n--- Voice Files (v) ---")
        v_folder_set = set(v_files_folder)
        v_excel_set = set(v_files_excel)
        
        missing_in_folder = v_excel_set - v_folder_set
        missing_in_excel = v_folder_set - v_excel_set
        
        if not missing_in_folder and not missing_in_excel:
            print("✅ Perfect match! All Excel records have corresponding files")
        else:
            if missing_in_folder:
                print("❌ Files in Excel but missing in folder:")
                for file in sorted(missing_in_folder):
                    print(f"   - {file}")
            if missing_in_excel:
                print("❌ Files in folder but missing in Excel:")
                for file in sorted(missing_in_excel):
                    print(f"   - {file}")
        
        print(f"Total in Excel: {len(v_files_excel)}, Total in folder: {len(v_files_folder)}")
        
        # Summary
        print("\n--- Summary ---")
        total_excel = len(m_files_excel) + len(v_files_excel)
        total_folder = len(m_files_folder) + len(v_files_folder)
        print(f"Total files in Excel: {total_excel}")
        print(f"Total files in folder: {total_folder}")
        
        if total_excel == total_folder:
            print("✅ Overall: Database and folder are synchronized")
        else:
            print("⚠️  Overall: Database and folder are NOT synchronized")
            
    except FileNotFoundError:
        print("❌ Excel file not found - cannot verify")
    except Exception as e:
        print(f"❌ Error during verification: {e}")

def select_audio_file():
    """Let user select audio file and return its path"""
    root = tk.Tk()
    root.withdraw()
    
    audio_file = filedialog.askopenfilename(
        title="Select Audio File",
        filetypes=[
            ("Audio files", "*.wav *.mp3 *.flac *.aiff *.aac *.ogg *.m4a"),
            ("All files", "*.*")
        ]
    )
    root.destroy()
    return audio_file

def get_corresponding_txt_file(audio_file):
    """Get the corresponding txt file path based on audio file name"""
    if not audio_file:
        return None
    
    # Replace audio extension with .txt
    base_name = os.path.splitext(audio_file)[0]
    txt_file = base_name + '.txt'
    
    return txt_file

def verify_files_exist(audio_file, txt_file):
    """Verify both audio and text files exist"""
    if not audio_file:
        print("No audio file selected. Exiting.")
        return False
    
    if not os.path.exists(audio_file):
        print(f"Error: Audio file not found: {audio_file}")
        return False
    
    if not os.path.exists(txt_file):
        print(f"Error: Text file not found: {txt_file}")
        print(f"Please create a text file named: {os.path.basename(txt_file)}")
        print("in the same folder as your audio file.")
        return False
    
    return True

def select_output_folder():
    """Let user select output folder for slices"""
    root = tk.Tk()
    root.withdraw()
    output_folder = filedialog.askdirectory(title="Select Output Folder for Slices")
    root.destroy()
    return output_folder

def select_audio_file():
    """Let user select audio file and return its path"""
    root = tk.Tk()
    root.withdraw()
    
    audio_file = filedialog.askopenfilename(
        title="Select Audio File",
        filetypes=[
            ("Audio files", "*.wav *.mp3 *.flac *.aiff *.aac *.ogg *.m4a"),
            ("All files", "*.*")
        ]
    )
    root.destroy()
    return audio_file

def get_corresponding_txt_file(audio_file):
    """Get the corresponding txt file path based on audio file name"""
    if not audio_file:
        return None
    
    # Replace audio extension with .txt
    base_name = os.path.splitext(audio_file)[0]
    txt_file = base_name + '.txt'
    
    return txt_file

def verify_files_exist(audio_file, txt_file):
    """Verify both audio and text files exist"""
    if not audio_file:
        print("No audio file selected. Exiting.")
        return False
    
    if not os.path.exists(audio_file):
        print(f"Error: Audio file not found: {audio_file}")
        return False
    
    if not os.path.exists(txt_file):
        audio_filename = os.path.basename(audio_file)
        txt_filename = os.path.basename(txt_file)
        
        print(f"Error: Text file not found: {txt_file}")
        print("\nTo create the required text file:")
        print(f"1. Open {audio_filename} in Audacity")
        print("2. Add labels at the climax points you want to slice")
        print("3. Export labels: File → Export → Export Labels...")
        print(f"4. Save as: {txt_filename}")
        print("5. Run this program again")
        return False
    
    return True

def select_output_folder():
    """Let user select output folder for slices"""
    root = tk.Tk()
    root.withdraw()
    output_folder = filedialog.askdirectory(title="Select Output Folder for Slices")
    root.destroy()
    return output_folder

def show_welcome_screen():
    """Display welcome message and program description"""
    welcome_text = """
╔══════════════════════════════════════════════════════════════╗
║                 pydub-audio-slicer-sequencer                 ║
║                     Audio Processing Tool                    ║
╚══════════════════════════════════════════════════════════════╝

Use this tool to:
1 - Slice an audio file into several blocks
2 - Sequence blocks to create an audio file

Option 1 will:
• Extract 30-second audio segments centered on climax points
• Apply fade in/out and normalization  
• Track all slices in an Excel database
• Maintain file organization and verification

Option 2 will:
• Sequence 30-second audio segments centered on climax points
"""
    print(welcome_text)
    
    while True:
        choice = input("Select option (1 or 2): ").strip()
        if choice in ['1', '2']:
            return choice
        else:
            print("Invalid choice. Please enter 1 or 2.")

    # Show welcome screen
    show_welcome_screen()
    
    print("=== Audio Slicer Started ===")
    print(f"Slice size: {SLICE_SIZE} seconds")
    # ... rest of your main function remains the same    print("=== Audio Slicer Started ===")
    print(f"Slice size: {SLICE_SIZE} seconds")
    print(f"Fade duration: {FADE_DURATION} seconds")
    print()
    
    # Select audio file
    print("Please select the audio file to slice...")
    audio_file = select_audio_file()
    
    # Get corresponding txt file
    txt_file = get_corresponding_txt_file(audio_file)
    
    # Verify files exist
    if not verify_files_exist(audio_file, txt_file):
        return
    
    # Select output folder
    print("Please select output folder for slices...")
    blocks_dir = select_output_folder()
    
    if not blocks_dir:
        print("No output folder selected. Exiting.")
        return
    
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    print(f"Audio file: {audio_file}")
    print(f"Text file: {txt_file}")
    print(f"Output directory: {blocks_dir}")
    print()
    
    # Parse audio.txt
    print("Parsing audio.txt...")
    slices = parse_audio_txt(txt_file)
    if not slices:
        print("No valid slices found in audio.txt")
        return
    
    print(f"Found {len(slices)} slices to process")
    for i, slice_info in enumerate(slices, 1):
        print(f"  {i}. {slice_info['type']} at {slice_info['climax_time']}s: {slice_info['description']}")
    print()
    
    # Get next file numbers from Excel
    next_m, next_v = get_next_file_numbers(excel_path)
    print(f"Next file numbers - m: {next_m}, v: {next_v}")
    print()
    
    # Load audio file
    print("Loading audio file...")
    try:
        audio = AudioSegment.from_file(audio_file)
        print(f"Audio loaded: {len(audio)/1000:.2f} seconds")
    except Exception as e:
        print(f"Error loading audio file: {e}")
        return
    
    # Process each slice
    print("\nProcessing slices...")
    for slice_info in slices:
        # Determine file number based on type
        if slice_info['type'] == 'm':
            file_number = next_m
            next_m += 1
        elif slice_info['type'] == 'v':
            file_number = next_v
            next_v += 1
        else:
            print(f"Warning: Unknown type '{slice_info['type']}', skipping")
            continue
        
        # Process the slice
        output_path = process_audio_slice(audio, slice_info, blocks_dir, file_number)
        if output_path:
            # Update Excel file
            update_excel_file(excel_path, slice_info, file_number, output_path, audio_file)
        print()
    
    # Verify files vs Excel database
    verify_files_vs_excel(blocks_dir, excel_path)
    
    print("=== Audio Slicer Completed ===")    # Initialize tkinter (hidden root window)
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    print("=== Audio Slicer Started ===")
    print(f"Slice size: {SLICE_SIZE} seconds")
    print(f"Fade duration: {FADE_DURATION} seconds")
    print()
    
    # Let user select audio file
    print("Please select the audio file to slice...")
    audio_file = filedialog.askopenfilename(
        title="Select Audio File",
        filetypes=[
            ("Audio files", "*.wav *.mp3 *.flac *.aiff *.aac *.ogg *.m4a"),
            ("All files", "*.*")
        ]
    )
    
    if not audio_file:
        print("No audio file selected. Exiting.")
        return
    
    # Let user select text file with slice definitions
    print("Please select the text file with slice definitions...")
    txt_file = filedialog.askopenfilename(
        title="Select Slice Definitions File",
        filetypes=[
            ("Text files", "*.txt"),
            ("All files", "*.*")
        ]
    )
    
    if not txt_file:
        print("No text file selected. Exiting.")
        return
    
    # Let user select output folder
    print("Please select output folder for slices...")
    blocks_dir = filedialog.askdirectory(title="Select Output Folder")
    
    if not blocks_dir:
        print("No output folder selected. Exiting.")
        return
    
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    print(f"Audio file: {audio_file}")
    print(f"Text file: {txt_file}")
    print(f"Output directory: {blocks_dir}")
    print()
    
    # Rest of your main function remains the same...
    # Remove the old path definitions and file existence checks    # Define paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    raw_audio_dir = os.path.join(base_dir, "raw audio")
    blocks_dir = os.path.join(base_dir, "blocks")
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    # Create blocks directory if it doesn't exist
    os.makedirs(blocks_dir, exist_ok=True)
    
    # File paths
    audio_file = os.path.join(raw_audio_dir, "audio.wav")
    txt_file = os.path.join(raw_audio_dir, "audio.txt")
    
    print("=== Audio Slicer Started ===")
    print(f"Slice size: {SLICE_SIZE} seconds")
    print(f"Fade duration: {FADE_DURATION} seconds")
    print(f"Audio file: {audio_file}")
    print(f"Text file: {txt_file}")
    print(f"Output directory: {blocks_dir}")
    print()
    
    # Check if input files exist
    if not os.path.exists(audio_file):
        print(f"Error: Audio file not found: {audio_file}")
        return
    if not os.path.exists(txt_file):
        print(f"Error: Text file not found: {txt_file}")
        return
    
    # Parse audio.txt
    print("Parsing audio.txt...")
    slices = parse_audio_txt(txt_file)
    if not slices:
        print("No valid slices found in audio.txt")
        return
    
    print(f"Found {len(slices)} slices to process")
    for i, slice_info in enumerate(slices, 1):
        print(f"  {i}. {slice_info['type']} at {slice_info['climax_time']}s: {slice_info['description']}")
    print()
    
    # Get next file numbers from Excel
    next_m, next_v = get_next_file_numbers(excel_path)
    print(f"Next file numbers - m: {next_m}, v: {next_v}")
    print()
    
    # Load audio file
    print("Loading audio file...")
    try:
        audio = AudioSegment.from_file(audio_file)
        print(f"Audio loaded: {len(audio)/1000:.2f} seconds")
    except Exception as e:
        print(f"Error loading audio file: {e}")
        return
    
    # Process each slice
    print("\nProcessing slices...")
    for slice_info in slices:
        # Determine file number based on type
        if slice_info['type'] == 'm':
            file_number = next_m
            next_m += 1
        elif slice_info['type'] == 'v':
            file_number = next_v
            next_v += 1
        else:
            print(f"Warning: Unknown type '{slice_info['type']}', skipping")
            continue
        
        # Process the slice
        output_path = process_audio_slice(audio, slice_info, blocks_dir, file_number)
        if output_path:
            # Update Excel file
            update_excel_file(excel_path, slice_info, file_number, output_path, audio_file)
        print()
    
    # Verify files vs Excel database
    verify_files_vs_excel(blocks_dir, excel_path)
    
    print("=== Audio Slicer Completed ===")

def run_audio_slicer():
    """Run the audio slicing functionality"""
    print("=== Audio Slicer Started ===")
    print(f"Slice size: {SLICE_SIZE} seconds")
    print(f"Fade duration: {FADE_DURATION} seconds")
    print()
    
    # Select audio file
    print("Please select the audio file to slice...")
    audio_file = select_audio_file()
    
    # Get corresponding txt file
    txt_file = get_corresponding_txt_file(audio_file)
    
    # Verify files exist
    if not verify_files_exist(audio_file, txt_file):
        return
    
    # Select output folder
    print("Please select output folder for slices...")
    blocks_dir = select_output_folder()
    
    if not blocks_dir:
        print("No output folder selected. Exiting.")
        return
    
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    print(f"Audio file: {audio_file}")
    print(f"Text file: {txt_file}")
    print(f"Output directory: {blocks_dir}")
    print()
    
    # Parse audio.txt
    print("Parsing audio.txt...")
    slices = parse_audio_txt(txt_file)
    if not slices:
        print("No valid slices found in audio.txt")
        return
    
    print(f"Found {len(slices)} slices to process")
    for i, slice_info in enumerate(slices, 1):
        print(f"  {i}. {slice_info['type']} at {slice_info['climax_time']}s: {slice_info['description']}")
    print()
    
    # Get next file numbers from Excel
    next_m, next_v = get_next_file_numbers(excel_path)
    print(f"Next file numbers - m: {next_m}, v: {next_v}")
    print()
    
    # Load audio file
    print("Loading audio file...")
    try:
        audio = AudioSegment.from_file(audio_file)
        print(f"Audio loaded: {len(audio)/1000:.2f} seconds")
    except Exception as e:
        print(f"Error loading audio file: {e}")
        return
    
    # Process each slice
    print("\nProcessing slices...")
    for slice_info in slices:
        # Determine file number based on type
        if slice_info['type'] == 'm':
            file_number = next_m
            next_m += 1
        elif slice_info['type'] == 'v':
            file_number = next_v
            next_v += 1
        else:
            print(f"Warning: Unknown type '{slice_info['type']}', skipping")
            continue
        
        # Process the slice
        output_path = process_audio_slice(audio, slice_info, blocks_dir, file_number)
        if output_path:
            # Update Excel file
            update_excel_file(excel_path, slice_info, file_number, output_path, audio_file)
        print()
    
    # Verify files vs Excel database
    verify_files_vs_excel(blocks_dir, excel_path)
    
    print("=== Audio Slicer Completed ===")

def main():
    choice = show_welcome_screen()
    
    if choice == '1':
        run_audio_slicer()  # This runs the actual slicer
    elif choice == '2':
        print("Sequencing feature - Under construction")
        return

if __name__ == "__main__":
    main()