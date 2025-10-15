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
from colorama import Fore, Back, Style, init
import random
# Initialize colorama (this makes colors work on Windows too)
init()

def get_base_path():
    """
    Get the base path for the application.
    Works both in development and in PyInstaller executable.
    """
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        base_path = os.path.dirname(sys.executable)
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return base_path

# Hardcoded parameters
SLICE_SIZE = 30  # seconds
FADE_DURATION = SLICE_SIZE / 2  # seconds

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

def parse_audio_txt(file_path, audio_duration=None):
    """Parse the audio.txt file and return list of slices that fit within audio boundaries"""
    slices = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                
                parts = line.split('\t')
                if len(parts) < 3:
                    print(f"{Fore.YELLOW}⚠️  Warning: Line {line_num} has invalid format: {line}{Style.RESET_ALL}")
                    continue
                
                try:
                    climax_time = float(parts[0])
                    audio_type = parts[2].split()[0]
                    description = ' '.join(parts[2].split()[1:])
                    
                    slice_begin = climax_time - (SLICE_SIZE / 2)
                    slice_end = climax_time + (SLICE_SIZE / 2)
                    
                    # Check start boundary
                    if slice_begin < 0:
                        print(f"{Fore.YELLOW}⚠️  Skipping slice at {climax_time}s: starts before audio beginning (needs {abs(slice_begin):.1f}s before start){Style.RESET_ALL}")
                        continue
                    
                    # Check end boundary (if audio_duration is provided)
                    if audio_duration and slice_end > audio_duration:
                        print(f"{Fore.YELLOW}⚠️  Skipping slice at {climax_time}s: extends beyond audio end (needs {slice_end - audio_duration:.1f}s after end){Style.RESET_ALL}")
                        continue
                    
                    slices.append({
                        'climax_time': climax_time,
                        'type': audio_type,
                        'description': description,
                        'slice_begin': slice_begin,
                        'slice_end': slice_end
                    })
                    
                except (ValueError, IndexError) as e:
                    print(f"{Fore.RED}❌ Error parsing line {line_num}: {line} - {e}{Style.RESET_ALL}")
                    continue
                    
    except FileNotFoundError:
        print(f"{Fore.RED}❌ Error: File {file_path} not found{Style.RESET_ALL}")
        return []
    except Exception as e:
        print(f"{Fore.RED}❌ Error reading {file_path}: {e}{Style.RESET_ALL}")
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
        print(f"{Fore.YELLOW}⚠️  blocks_list.xlsx not found, creating new file...{Style.RESET_ALL}")
        next_m, next_v = 1, 1
    except Exception as e:
        print(f"{Fore.RED}❌ Error reading Excel file: {e}{Style.RESET_ALL}")
        next_m, next_v = 1, 1
    
    return next_m, next_v

def process_audio_slice(audio, slice_info, output_folder, file_number):
    """Process a single audio slice and export as MP3 192kbps"""
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
        
        # Export file as MP3 192kbps
        filename = f"{slice_info['type']}{file_number}.mp3"
        output_path = os.path.join(output_folder, filename)
        slice_audio.export(output_path, format="mp3", bitrate="192k")
        
        print(f"{Fore.GREEN}✅ Successfully created: {filename}{Style.RESET_ALL}")
        print(f"{Fore.BLUE}   (from {slice_info['slice_begin']:.1f}s to {slice_info['slice_end']:.1f}s){Style.RESET_ALL}")
        return output_path
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error processing slice: {e}{Style.RESET_ALL}")
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
                
        print(f"{Fore.GREEN}✅ Updated Excel: {slice_info['type']}{file_number}{Style.RESET_ALL}")
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error updating Excel file: {e}{Style.RESET_ALL}")

def verify_files_vs_excel(blocks_dir, excel_path):
    """Verify that files in blocks folder match the Excel database"""
    print(f"\n{Fore.CYAN}=== Verifying Files vs Excel Database ==={Style.RESET_ALL}")
    
    try:
        # Read Excel sheets
        m_df = pd.read_excel(excel_path, sheet_name='m')
        v_df = pd.read_excel(excel_path, sheet_name='v')
        
        # Get all files in blocks directory
        all_files = os.listdir(blocks_dir)
        # Look for both WAV and MP3 files
        audio_files = [f for f in all_files if f.endswith('.wav') or f.endswith('.mp3')]
        
        # Extract m and v files from folder
        m_files_folder = [f for f in audio_files if f.startswith('m') and f[1:].split('.')[0].isdigit()]
        v_files_folder = [f for f in audio_files if f.startswith('v') and f[1:].split('.')[0].isdigit()]
        
        # Get file names from Excel (remove file extension for comparison)
        m_files_excel = []
        if not m_df.empty and 'm' in m_df.columns:
            m_files_excel = [f"{row['m']}.mp3" for _, row in m_df.iterrows() if pd.notna(row['m'])]  # Changed to .mp3
        
        v_files_excel = []
        if not v_df.empty and 'v' in v_df.columns:
            v_files_excel = [f"{row['v']}.mp3" for _, row in v_df.iterrows() if pd.notna(row['v'])]  # Changed to .mp3
        
        # Compare Music files (m)
        print(f"\n{Fore.CYAN}--- Music Files (m) ---{Style.RESET_ALL}")
        m_folder_set = set(m_files_folder)
        m_excel_set = set(m_files_excel)
        
        missing_in_folder = m_excel_set - m_folder_set
        missing_in_excel = m_folder_set - m_excel_set
        
        if not missing_in_folder and not missing_in_excel:
            print(f"{Fore.GREEN}✅ Perfect match! All Excel records have corresponding files{Style.RESET_ALL}")
        else:
            if missing_in_folder:
                print(f"{Fore.RED}❌ Files in Excel but missing in folder:{Style.RESET_ALL}")
                for file in sorted(missing_in_folder):
                    print(f"   - {file}")
            if missing_in_excel:
                print(f"{Fore.RED}❌ Files in folder but missing in Excel:{Style.RESET_ALL}")
                for file in sorted(missing_in_excel):
                    print(f"   - {file}")
        
        print(f"Total in Excel: {len(m_files_excel)}, Total in folder: {len(m_files_folder)}")
        
        # Compare Voice files (v)
        print(f"\n{Fore.CYAN}--- Voice Files (v) ---{Style.RESET_ALL}")
        v_folder_set = set(v_files_folder)
        v_excel_set = set(v_files_excel)
        
        missing_in_folder = v_excel_set - v_folder_set
        missing_in_excel = v_folder_set - v_excel_set
        
        if not missing_in_folder and not missing_in_excel:
            print(f"{Fore.GREEN}✅ Perfect match! All Excel records have corresponding files{Style.RESET_ALL}")
        else:
            if missing_in_folder:
                print(f"{Fore.RED}❌ Files in Excel but missing in folder:{Style.RESET_ALL}")
                for file in sorted(missing_in_folder):
                    print(f"   - {file}")
            if missing_in_excel:
                print(f"{Fore.RED}❌ Files in folder but missing in Excel:{Style.RESET_ALL}")
                for file in sorted(missing_in_excel):
                    print(f"   - {file}")
        
        print(f"Total in Excel: {len(v_files_excel)}, Total in folder: {len(v_files_folder)}")
        
        # Summary
        print(f"\n{Fore.CYAN}--- Summary ---{Style.RESET_ALL}")
        total_excel = len(m_files_excel) + len(v_files_excel)
        total_folder = len(m_files_folder) + len(v_files_folder)
        print(f"Total files in Excel: {total_excel}")
        print(f"Total files in folder: {total_folder}")
        
        if total_excel == total_folder:
            print(f"{Fore.GREEN}✅ Overall: Database and folder are synchronized{Style.RESET_ALL}")
        else:
            print(f"{Fore.YELLOW}⚠️  Overall: Database and folder are NOT synchronized{Style.RESET_ALL}")
            
    except FileNotFoundError:
        print(f"{Fore.RED}❌ Excel file not found - cannot verify{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}❌ Error during verification: {e}{Style.RESET_ALL}")

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
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return False
    
    if not os.path.exists(audio_file):
        print(f"{Fore.RED}❌ Error: Audio file not found: {audio_file}{Style.RESET_ALL}")
        return False
    
    if not os.path.exists(txt_file):
        audio_filename = os.path.basename(audio_file)
        txt_filename = os.path.basename(txt_file)
        
        print(f"{Fore.RED}╔══════════════════════════════════════════════════════════════╗{Style.RESET_ALL}")
        print(f"{Fore.RED}║                       FILE NOT FOUND                         ║{Style.RESET_ALL}")
        print(f"{Fore.RED}╚══════════════════════════════════════════════════════════════╗{Style.RESET_ALL}")
        print(f"{Fore.RED}❌ Error: Text file not found:{Style.RESET_ALL}")
        print(f"{Fore.RED}{txt_file}{Style.RESET_ALL}")
        print()
        print(f"{Fore.YELLOW}📝 To create the required text file:{Style.RESET_ALL}")
        print(f"{Fore.WHITE}1. Open {Style.BRIGHT}{audio_filename}{Style.RESET_ALL}{Fore.WHITE} in Audacity{Style.RESET_ALL}")
        print(f"{Fore.WHITE}2. Add labels at the climax points you want to slice{Style.RESET_ALL}")
        print(f"{Fore.WHITE}3. Export labels: File → Export → Export Labels...{Style.RESET_ALL}")
        print(f"{Fore.WHITE}4. Save as: {Style.BRIGHT}{txt_filename}{Style.RESET_ALL}{Fore.WHITE} in the same folder{Style.RESET_ALL}")
        print(f"{Fore.WHITE}5. Run this program again{Style.RESET_ALL}")
        print(f"{Fore.RED}╔══════════════════════════════════════════════════════════════╗{Style.RESET_ALL}")
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
    welcome_text = f"""
{Fore.CYAN}╔══════════════════════════════════════════════════════════════╗{Style.RESET_ALL}
{Fore.CYAN}║                 AUDIO SLICER & SEQUENCER                    ║{Style.RESET_ALL}
{Fore.CYAN}║                     Audio Processing Tool                   ║{Style.RESET_ALL}
{Fore.CYAN}╚══════════════════════════════════════════════════════════════╝{Style.RESET_ALL}

{Fore.WHITE}Use this tool to:{Style.RESET_ALL}
{Fore.GREEN}1 - Slice an audio file into several blocks{Style.RESET_ALL}
{Fore.BLUE}2 - Sequence blocks to create an audio file{Style.RESET_ALL}  
{Fore.MAGENTA}3 - Slice an audio file and produce a sequence{Style.RESET_ALL}
{Fore.YELLOW}4 - Help!{Style.RESET_ALL}
"""
    print(welcome_text)
    
    while True:
        choice = input(f"{Fore.WHITE}Select option (1, 2, 3, or 4): {Style.RESET_ALL}").strip()
        if choice in ['1', '2', '3', '4']:
            return choice
        else:
            print(f"{Fore.RED}❌ Invalid choice. Please enter 1, 2, 3, or 4.{Style.RESET_ALL}")

def show_help():
    """Display help information"""
    help_text = f"""
{Fore.CYAN}╔══════════════════════════════════════════════════════════════╗{Style.RESET_ALL}
{Fore.CYAN}║                         HELP - AUDIO SLICER & SEQUENCER     ║{Style.RESET_ALL}
{Fore.CYAN}╚══════════════════════════════════════════════════════════════╝{Style.RESET_ALL}

{Fore.YELLOW}Option 1 will:{Style.RESET_ALL}
• Extract 30-second audio segments centered on climax points
• Apply fade in/out and normalization  
• Track all slices in an Excel database
• Maintain file organization and verification

{Fore.YELLOW}Option 2 will:{Style.RESET_ALL}
• Sequence existing blocks into a new audio file

{Fore.YELLOW}Option 3 will:{Style.RESET_ALL}
• Extract 30-second audio segments centered on climax points
• Apply fade in/out and normalization
• Automatically sequence the slices into a new audio file
• Track all slices in an Excel database

{Fore.GREEN}Future updates will include a link to comprehensive documentation on GitHub.{Style.RESET_ALL}

{Fore.WHITE}Press Enter to return to the main menu...{Style.RESET_ALL}
"""
    print(help_text)
    input()

def show_slice_and_sequence_menu():
    """Show submenu for slice and sequence options"""
    submenu_text = f"""
{Fore.CYAN}Slice and Sequence Options:{Style.RESET_ALL}
{Fore.GREEN}1 - I have an audio file with labels{Style.RESET_ALL}
{Fore.YELLOW}2 - I am a lazy bastard and just want random results{Style.RESET_ALL}
"""
    print(submenu_text)
    
    while True:
        sub_choice = input(f"{Fore.WHITE}Select option (1 or 2): {Style.RESET_ALL}").strip()
        if sub_choice in ['1', '2']:
            return sub_choice
        else:
            print(f"{Fore.RED}❌ Invalid choice. Please enter 1 or 2.{Style.RESET_ALL}")

def show_slice_options_menu():
    """Show submenu for slicing options"""
    submenu_text = f"""
{Fore.CYAN}Slicing Options:{Style.RESET_ALL}
{Fore.GREEN}1 - I have a label file I obtained from Audacity{Style.RESET_ALL}
{Fore.YELLOW}2 - I don't have a label file{Style.RESET_ALL}
"""
    print(submenu_text)
    
    while True:
        sub_choice = input(f"{Fore.WHITE}Select option (1 or 2): {Style.RESET_ALL}").strip()
        if sub_choice in ['1', '2']:
            return sub_choice
        else:
            print(f"{Fore.RED}❌ Invalid choice. Please enter 1 or 2.{Style.RESET_ALL}")

def show_no_labels_menu():
    """Show options when user doesn't have labels"""
    submenu_text = f"""
{Fore.CYAN}No Label File Options:{Style.RESET_ALL}
{Fore.GREEN}1 - I will label my audio file in Audacity{Style.RESET_ALL}
{Fore.YELLOW}2 - I just want to randomly slice my audio file{Style.RESET_ALL}
"""
    print(submenu_text)
    
    while True:
        choice = input(f"{Fore.WHITE}Select option (1 or 2): {Style.RESET_ALL}").strip()
        if choice in ['1', '2']:
            return choice
        else:
            print(f"{Fore.RED}❌ Invalid choice. Please enter 1 or 2.{Style.RESET_ALL}")

def slice_audio_from_labels(audio_file, blocks_dir):
    """Slice audio file using labels and return the blocks directory"""
    # Get corresponding txt file FIRST - THIS MUST BE AT THE BEGINNING
    txt_file = get_corresponding_txt_file(audio_file)
    
    # Verify files exist
    if not verify_files_exist(audio_file, txt_file):
        return None
    
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    print(f"{Fore.GREEN}Audio file: {audio_file}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Text file: {txt_file}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Output directory: {blocks_dir}{Style.RESET_ALL}")
    print()
    
    # Load audio file first to get duration
    print(f"{Fore.BLUE}Loading audio file...{Style.RESET_ALL}")
    try:
        audio = AudioSegment.from_file(audio_file)
        audio_duration = len(audio) / 1000
        print(f"{Fore.GREEN}✅ Audio loaded: {audio_duration:.2f} seconds{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}❌ Error loading audio file: {e}{Style.RESET_ALL}")
        return None
    
    # Parse audio.txt with duration checking
    print(f"{Fore.BLUE}Parsing audio.txt...{Style.RESET_ALL}")
    slices = parse_audio_txt(txt_file, audio_duration)  # Now txt_file is defined!
    
    if not slices:
        print(f"{Fore.YELLOW}⚠️  No valid slices found in audio.txt{Style.RESET_ALL}")
        return None
    
    print(f"{Fore.GREEN}Found {len(slices)} slices to process{Style.RESET_ALL}")
    for i, slice_info in enumerate(slices, 1):
        print(f"  {i}. {slice_info['type']} at {slice_info['climax_time']}s: {slice_info['description']}")
    print()
    
    # Get next file numbers from Excel
    next_m, next_v = get_next_file_numbers(excel_path)
    print(f"{Fore.BLUE}Next file numbers - m: {next_m}, v: {next_v}{Style.RESET_ALL}")
    print()
    
    # Process each slice
    print(f"\n{Fore.CYAN}Processing slices...{Style.RESET_ALL}")
    for slice_info in slices:
        # Determine file number based on type
        if slice_info['type'] == 'm':
            file_number = next_m
            next_m += 1
        elif slice_info['type'] == 'v':
            file_number = next_v
            next_v += 1
        else:
            print(f"{Fore.YELLOW}⚠️  Warning: Unknown type '{slice_info['type']}', skipping{Style.RESET_ALL}")
            continue
        
        # Process the slice as MP3
        output_path = process_audio_slice(audio, slice_info, blocks_dir, file_number)
        if output_path:
            # Update Excel file
            update_excel_file(excel_path, slice_info, file_number, output_path, audio_file)
        print()
    
    # Verify files vs Excel database
    verify_files_vs_excel(blocks_dir, excel_path)
    
    print(f"{Fore.GREEN}✅ Audio slicing completed!{Style.RESET_ALL}")
    return blocks_dir

def calculate_slice_density(audio_duration_seconds):
    """Calculate number of slices based on audio duration (~1 per 2 minutes)"""
    base_slices = audio_duration_seconds / 120  # 1 slice per 2 minutes
    # Add some variation (80% to 120% of base)
    variation = random.uniform(0.8, 1.2)
    num_slices = max(1, int(base_slices * variation))
    return num_slices

def get_next_file_numbers_from_folder(blocks_dir):
    """Get next file numbers by scanning existing m and v files in folder"""
    try:
        if not os.path.exists(blocks_dir):
            return 1, 1
        
        all_files = os.listdir(blocks_dir)
        m_files = [f for f in all_files if f.startswith('m') and f[1:].split('.')[0].isdigit()]
        v_files = [f for f in all_files if f.startswith('v') and f[1:].split('.')[0].isdigit()]
        
        # Extract numbers and find maximums
        m_numbers = [int(f[1:].split('.')[0]) for f in m_files if f[1:].split('.')[0].isdigit()]
        v_numbers = [int(f[1:].split('.')[0]) for f in v_files if f[1:].split('.')[0].isdigit()]
        
        next_m = max(m_numbers) + 1 if m_numbers else 1
        next_v = max(v_numbers) + 1 if v_numbers else 1
        
        return next_m, next_v
        
    except Exception as e:
        print(f"{Fore.YELLOW}⚠️  Error scanning folder, starting from 1: {e}{Style.RESET_ALL}")
        return 1, 1

def generate_random_labels(audio_file):
    """Generate random slice positions throughout the audio file with proper density"""
    try:
        print(f"{Fore.BLUE}Loading audio to calculate duration...{Style.RESET_ALL}")
        # Load audio to get duration
        audio = AudioSegment.from_file(audio_file)
        duration_seconds = len(audio) / 1000
        print(f"{Fore.GREEN}✅ Audio duration: {duration_seconds:.1f} seconds{Style.RESET_ALL}")
        
        # Calculate appropriate number of slices
        num_slices = calculate_slice_density(duration_seconds)
        print(f"{Fore.BLUE}Calculated {num_slices} slices for {duration_seconds:.1f}s audio{Style.RESET_ALL}")
        
        # Ensure we don't try to create slices beyond audio length
        max_start_time = duration_seconds - SLICE_SIZE
        if max_start_time <= 0:
            print(f"{Fore.RED}❌ Audio file is too short ({duration_seconds:.1f}s) for {SLICE_SIZE}s slices{Style.RESET_ALL}")
            return None
        
        slices = []
        for i in range(num_slices):
            # Generate random start time within valid range with some spacing
            min_spacing = SLICE_SIZE * 1.5  # Ensure slices don't overlap too much
            max_attempts = 100
            attempt = 0
            
            while attempt < max_attempts:
                start_time = random.uniform(0, max_start_time)
                climax_time = start_time + (SLICE_SIZE / 2)
                
                # Check if this slice overlaps significantly with existing slices
                overlap = False
                for existing_slice in slices:
                    if abs(existing_slice['climax_time'] - climax_time) < min_spacing:
                        overlap = True
                        break
                
                if not overlap:
                    break
                attempt += 1
            
            # Randomly choose between 'm' (music) and 'v' (voice)
            audio_type = random.choice(['m', 'v'])
            description = f"random_{audio_type}_{i+1}"
            
            slices.append({
                'climax_time': climax_time,
                'type': audio_type,
                'description': description,
                'slice_begin': start_time,
                'slice_end': start_time + SLICE_SIZE
            })
        
        print(f"{Fore.GREEN}✅ Generated {num_slices} random slices for {duration_seconds:.1f}s audio{Style.RESET_ALL}")
        return slices
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error generating random labels: {e}{Style.RESET_ALL}")
        return None

def process_audio_slice_mp3(audio, slice_info, output_folder, file_number):
    """Process a single audio slice and export as MP3 192kbps"""
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
        
        # Export file as MP3 192kbps
        filename = f"{slice_info['type']}{file_number}.mp3"
        output_path = os.path.join(output_folder, filename)
        slice_audio.export(output_path, format="mp3", bitrate="192k")
        
        print(f"{Fore.GREEN}✅ Successfully created: {filename}{Style.RESET_ALL}")
        print(f"{Fore.BLUE}   (from {slice_info['slice_begin']:.1f}s to {slice_info['slice_end']:.1f}s){Style.RESET_ALL}")
        return output_path
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error processing slice: {e}{Style.RESET_ALL}")
        return None

def run_random_slicer():
    """Run random audio slicing functionality"""
    print(f"{Fore.CYAN}=== Random Audio Slicer Started ==={Style.RESET_ALL}")
    print(f"Slice size: {SLICE_SIZE} seconds")
    print(f"Fade duration: {FADE_DURATION} seconds")
    print(f"Output format: MP3 192kbps{Style.RESET_ALL}")
    print()
    
    # Select audio file
    print(f"{Fore.BLUE}Please select the audio file to slice...{Style.RESET_ALL}")
    audio_file = select_audio_file()
    
    if not audio_file:
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return
    
    # Generate random slices
    print(f"{Fore.BLUE}Generating random slices...{Style.RESET_ALL}")
    slices = generate_random_labels(audio_file)
    if not slices:
        return
    
    # Select output folder
    print(f"{Fore.BLUE}Please select output folder for slices...{Style.RESET_ALL}")
    blocks_dir = select_output_folder()
    
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    # Get next file numbers from existing files in folder
    next_m, next_v = get_next_file_numbers_from_folder(blocks_dir)
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    print(f"{Fore.GREEN}Audio file: {audio_file}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Label source: Randomly generated ({len(slices)} slices){Style.RESET_ALL}")
    print(f"{Fore.GREEN}Output directory: {blocks_dir}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Starting file numbers - m: {next_m}, v: {next_v}{Style.RESET_ALL}")
    print()
    
    print(f"{Fore.GREEN}Found {len(slices)} slices to process{Style.RESET_ALL}")
    for i, slice_info in enumerate(slices, 1):
        print(f"  {i}. {slice_info['type']} at {slice_info['climax_time']:.1f}s: {slice_info['description']}")
    print()
    
    # Load audio file
    print(f"{Fore.BLUE}Loading audio file...{Style.RESET_ALL}")
    try:
        audio = AudioSegment.from_file(audio_file)
        print(f"{Fore.GREEN}✅ Audio loaded: {len(audio)/1000:.2f} seconds{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}❌ Error loading audio file: {e}{Style.RESET_ALL}")
        return
    
    # Process each slice
    print(f"\n{Fore.CYAN}Processing slices...{Style.RESET_ALL}")
    for slice_info in slices:
        # Determine file number based on type
        if slice_info['type'] == 'm':
            file_number = next_m
            next_m += 1
        elif slice_info['type'] == 'v':
            file_number = next_v
            next_v += 1
        else:
            print(f"{Fore.YELLOW}⚠️  Warning: Unknown type '{slice_info['type']}', skipping{Style.RESET_ALL}")
            continue
        
        # Process the slice as MP3
        output_path = process_audio_slice_mp3(audio, slice_info, blocks_dir, file_number)
        if output_path:
            # Update Excel file
            update_excel_file(excel_path, slice_info, file_number, output_path, audio_file)
        print()
    
    # Verify files vs Excel database
    verify_files_vs_excel(blocks_dir, excel_path)
    
    print(f"{Fore.CYAN}=== Random Audio Slicer Completed ==={Style.RESET_ALL}")

def scan_available_blocks(blocks_dir):
    """Scan blocks directory for m and v audio files"""
    if not os.path.exists(blocks_dir):
        return [], []
    
    all_files = os.listdir(blocks_dir)
    # Look for both WAV and MP3 files
    m_blocks = [f for f in all_files if f.startswith('m') and (f.endswith('.mp3') or f.endswith('.wav'))]
    v_blocks = [f for f in all_files if f.startswith('v') and (f.endswith('.mp3') or f.endswith('.wav'))]
    
    # Sort by number for consistent ordering before shuffling
    m_blocks.sort(key=lambda x: int(x[1:].split('.')[0]))
    v_blocks.sort(key=lambda x: int(x[1:].split('.')[0]))
    
    return m_blocks, v_blocks  
  
def validate_sequence_requirements(m_blocks, v_blocks):
    """Validate that we have enough blocks for sequencing"""
    if len(m_blocks) < 3:
        print(f"{Fore.RED}❌ Not enough music blocks: {len(m_blocks)} found (minimum 3 required){Style.RESET_ALL}")
        return False
    if len(v_blocks) < 3:
        print(f"{Fore.RED}❌ Not enough voice blocks: {len(v_blocks)} found (minimum 3 required){Style.RESET_ALL}")
        return False
    
    print(f"{Fore.GREEN}✅ Found {len(m_blocks)} music blocks and {len(v_blocks)} voice blocks{Style.RESET_ALL}")
    return True

def create_random_sequence(m_blocks, v_blocks):
    """Create random sequences for both channels"""
    # Shuffle blocks randomly
    random.shuffle(m_blocks)
    random.shuffle(v_blocks)
    
    # Use the minimum length to determine sequence duration
    sequence_length = min(len(m_blocks), len(v_blocks))
    
    # Take only the blocks we'll actually use
    m_sequence = m_blocks[:sequence_length]
    v_sequence = v_blocks[:sequence_length]
    
    # Show which blocks are used vs skipped
    print(f"{Fore.BLUE}🎵 Sequence Configuration:{Style.RESET_ALL}")
    print(f"{Fore.GREEN}   Using {sequence_length} blocks from each channel{Style.RESET_ALL}")
    
    if len(m_blocks) > sequence_length:
        print(f"{Fore.YELLOW}   Skipping {len(m_blocks) - sequence_length} music blocks: {m_blocks[sequence_length:]}{Style.RESET_ALL}")
    if len(v_blocks) > sequence_length:
        print(f"{Fore.YELLOW}   Skipping {len(v_blocks) - sequence_length} voice blocks: {v_blocks[sequence_length:]}{Style.RESET_ALL}")
    
    return m_sequence, v_sequence

def build_multi_channel_sequence(blocks_dir, m_sequence, v_sequence):
    """Build the final sequence with 15-second voice channel offset"""
    try:
        print(f"{Fore.BLUE}🔊 Building audio sequence...{Style.RESET_ALL}")
        
        # Create 15 seconds of silence for voice channel offset
        silence_15s = AudioSegment.silent(duration=15000)  # 15 seconds in milliseconds
        
        # Initialize channels
        music_channel = AudioSegment.empty()
        voice_channel = silence_15s  # Start voice channel with 15s silence
        
        # Load and concatenate music blocks
        print(f"{Fore.BLUE}   Loading music channel...{Style.RESET_ALL}")
        for i, block in enumerate(m_sequence, 1):
            block_path = os.path.join(blocks_dir, block)
            audio_segment = AudioSegment.from_file(block_path)
            music_channel += audio_segment
            print(f"{Fore.GREEN}     [{i}/{len(m_sequence)}] Added: {block}{Style.RESET_ALL}")
        
        # Load and concatenate voice blocks
        print(f"{Fore.BLUE}   Loading voice channel...{Style.RESET_ALL}")
        for i, block in enumerate(v_sequence, 1):
            block_path = os.path.join(blocks_dir, block)
            audio_segment = AudioSegment.from_file(block_path)
            voice_channel += audio_segment
            print(f"{Fore.GREEN}     [{i}/{len(v_sequence)}] Added: {block}{Style.RESET_ALL}")
        
        # Ensure both channels are the same length (pad with silence if needed)
        if len(music_channel) > len(voice_channel):
            voice_channel += AudioSegment.silent(duration=len(music_channel) - len(voice_channel))
        elif len(voice_channel) > len(music_channel):
            music_channel += AudioSegment.silent(duration=len(voice_channel) - len(music_channel))
        
        # Mix the two stereo channels
        print(f"{Fore.BLUE}   Mixing channels...{Style.RESET_ALL}")
        final_audio = music_channel.overlay(voice_channel)
        
        print(f"{Fore.GREEN}✅ Sequence built: {len(final_audio)/1000:.1f}s total duration{Style.RESET_ALL}")
        return final_audio
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error building sequence: {e}{Style.RESET_ALL}")
        return None

def run_sequencer():
    """Main sequencing workflow"""
    print(f"{Fore.CYAN}=== Audio Sequencer Started ==={Style.RESET_ALL}")
    print(f"{Fore.BLUE}This will create a mixed sequence with:{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Music channel starting at 0:00{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Voice channel starting at 0:15{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Random block order{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Stereo output{Style.RESET_ALL}")
    print()
    
    # Select blocks directory
    print(f"{Fore.BLUE}Please select the blocks folder...{Style.RESET_ALL}")
    blocks_dir = filedialog.askdirectory(title="Select Blocks Folder")
    
    if not blocks_dir:
        print(f"{Fore.RED}❌ No blocks folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    # Scan for available blocks
    print(f"{Fore.BLUE}Scanning for audio blocks...{Style.RESET_ALL}")
    m_blocks, v_blocks = scan_available_blocks(blocks_dir)
    
    if not validate_sequence_requirements(m_blocks, v_blocks):
        return
    
    # Create random sequence
    m_sequence, v_sequence = create_random_sequence(m_blocks, v_blocks)
    
    # Build the audio sequence
    final_audio = build_multi_channel_sequence(blocks_dir, m_sequence, v_sequence)
    if not final_audio:
        return
    
    # Let user choose output location and filename
    print(f"{Fore.BLUE}Please choose where to save the sequence...{Style.RESET_ALL}")
    output_path = filedialog.asksaveasfilename(
        title="Save Sequence As",
        defaultextension=".mp3",
        filetypes=[("MP3 files", "*.mp3"), ("All files", "*.*")]
    )
    
    if not output_path:
        print(f"{Fore.RED}❌ No output file selected. Exiting.{Style.RESET_ALL}")
        return
    
    # Export final sequence
    try:
        print(f"{Fore.BLUE}Exporting sequence...{Style.RESET_ALL}")
        final_audio.export(output_path, format="mp3", bitrate="192k")
        print(f"{Fore.GREEN}✅ Sequence saved: {output_path}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}🎵 Final duration: {len(final_audio)/1000:.1f} seconds{Style.RESET_ALL}")
        
        # Generate timeline file - MOVED TO AFTER final_audio IS DEFINED
        audio_duration = len(final_audio) / 1000
        generate_sequence_timeline(output_path, blocks_dir, m_sequence, v_sequence, audio_duration)
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error exporting sequence: {e}{Style.RESET_ALL}")
        return
    
    print(f"{Fore.CYAN}=== Audio Sequencer Completed ==={Style.RESET_ALL}")

def run_slice_and_sequence_with_labels():
    """Run complete workflow: slice audio then sequence the blocks"""
    print(f"{Fore.CYAN}=== Slice & Sequence with Labels ==={Style.RESET_ALL}")
    
    # 1. Select audio file
    print(f"{Fore.BLUE}Step 1: Select the source audio file...{Style.RESET_ALL}")
    audio_file = select_audio_file()
    
    if not audio_file:
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return
    
    # 2. Select output folder for slices
    print(f"{Fore.BLUE}Step 2: Select output folder for audio slices...{Style.RESET_ALL}")
    blocks_dir = select_output_folder()
    
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    # 3. Slice audio using labels
    print(f"{Fore.BLUE}Step 3: Slicing audio with labels...{Style.RESET_ALL}")
    result_dir = slice_audio_from_labels(audio_file, blocks_dir)
    
    if not result_dir:
        print(f"{Fore.RED}❌ Audio slicing failed. Exiting.{Style.RESET_ALL}")
        return
    
    # 4. Scan for created blocks
    print(f"{Fore.BLUE}Step 4: Scanning for created blocks...{Style.RESET_ALL}")
    m_blocks, v_blocks = scan_available_blocks(blocks_dir)
    
    if not validate_sequence_requirements(m_blocks, v_blocks):
        print(f"{Fore.RED}❌ Not enough blocks for sequencing{Style.RESET_ALL}")
        return
    
    # 5. Create random sequence
    m_sequence, v_sequence = create_random_sequence(m_blocks, v_blocks)
    
    # 6. Build the audio sequence
    final_audio = build_multi_channel_sequence(blocks_dir, m_sequence, v_sequence)
    if not final_audio:
        print(f"{Fore.RED}❌ Error building audio sequence{Style.RESET_ALL}")
        return
    
    # 7. Let user choose output location
    print(f"{Fore.BLUE}Step 5: Save final sequence...{Style.RESET_ALL}")
    output_path = filedialog.asksaveasfilename(
        title="Save Final Sequence As",
        defaultextension=".mp3",
        filetypes=[("MP3 files", "*.mp3"), ("All files", "*.*")]
    )
    
    if not output_path:
        print(f"{Fore.RED}❌ No output file selected. Exiting.{Style.RESET_ALL}")
        return
    
    # 8. Export final sequence
    try:
        print(f"{Fore.BLUE}Exporting final sequence...{Style.RESET_ALL}")
        final_audio.export(output_path, format="mp3", bitrate="192k")
        print(f"{Fore.GREEN}✅ Final sequence saved: {output_path}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}🎵 Final duration: {len(final_audio)/1000:.1f} seconds{Style.RESET_ALL}")
        
        # Generate timeline file
        audio_duration = len(final_audio) / 1000
        generate_sequence_timeline(output_path, blocks_dir, m_sequence, v_sequence, audio_duration)
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error exporting sequence: {e}{Style.RESET_ALL}")
        return
    
    print(f"{Fore.CYAN}=== Slice & Sequence Workflow Completed ==={Style.RESET_ALL}")
    
def generate_sequence_timeline(sequence_path, blocks_dir, m_sequence, v_sequence, audio_duration):
    """Generate a timeline text file for the created sequence"""
    try:
        # Create txt file path (same name, different extension)
        txt_path = os.path.splitext(sequence_path)[0] + '.txt'
        
        # Get current timestamp
        from datetime import datetime
        created_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Calculate total duration in MM:SS format
        minutes = int(audio_duration // 60)
        seconds = int(audio_duration % 60)
        duration_str = f"{minutes:02d}:{seconds:02d}"
        
        # Count blocks by category dynamically
        block_counts = {}
        for block in m_sequence + v_sequence:
            category = block[0]  # First character: 'm', 'v', etc.
            block_counts[category] = block_counts.get(category, 0) + 1
        
        # Create blocks used string
        blocks_used = ", ".join([f"{cat}={count}" for cat, count in sorted(block_counts.items())])
        
        # Read descriptions from Excel
        excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
        descriptions = {}
        try:
            # Try to read all sheets and build description mapping
            for category in block_counts.keys():
                try:
                    df = pd.read_excel(excel_path, sheet_name=category)
                    if not df.empty and category in df.columns:
                        for _, row in df.iterrows():
                            if pd.notna(row[category]):
                                filename = row[category]
                                desc = row.get('description', 'No description')
                                descriptions[filename] = desc
                except:
                    continue
        except Exception as e:
            print(f"{Fore.YELLOW}⚠️  Could not read Excel for descriptions: {e}{Style.RESET_ALL}")
        
        # Build timeline entries with interleaved music and voice
        timeline_entries = []
        
        # Calculate start times for each block
        for i in range(len(m_sequence)):
            # Music block start time
            music_time = i * 30
            music_minutes = music_time // 60
            music_seconds = music_time % 60
            music_time_str = f"{music_minutes:02d}:{music_seconds:02d}"
            
            music_block = m_sequence[i]
            music_name = os.path.splitext(music_block)[0]
            music_desc = descriptions.get(music_name, 'No description')
            
            timeline_entries.append({
                'time': music_time,
                'time_str': music_time_str,
                'block': music_name,
                'description': music_desc
            })
            
            # Voice block start time (15 seconds after corresponding music)
            if i < len(v_sequence):
                voice_time = (i * 30) + 15
                voice_minutes = voice_time // 60
                voice_seconds = voice_time % 60
                voice_time_str = f"{voice_minutes:02d}:{voice_seconds:02d}"
                
                voice_block = v_sequence[i]
                voice_name = os.path.splitext(voice_block)[0]
                voice_desc = descriptions.get(voice_name, 'No description')
                
                timeline_entries.append({
                    'time': voice_time,
                    'time_str': voice_time_str,
                    'block': voice_name,
                    'description': voice_desc
                })
        
        # Sort all entries by time
        timeline_entries.sort(key=lambda x: x['time'])
        
        # Write the timeline file
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(f"Sequence: {os.path.basename(sequence_path)}\n")
            f.write(f"Created: {created_time}\n")
            f.write(f"Blocks source: {blocks_dir}\n")
            f.write(f"Total duration: {duration_str}\n")
            f.write(f"Blocks used: {blocks_used}\n\n")
            
            for entry in timeline_entries:
                f.write(f"{entry['time_str']} / {entry['block']} / {entry['description']}\n")
        
        print(f"{Fore.GREEN}✅ Timeline saved: {txt_path}{Style.RESET_ALL}")
        return True
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error generating timeline: {e}{Style.RESET_ALL}")
        return False

def calculate_max_possible_minutes(audio_duration_seconds, slice_duration=30, min_spacing=60):
    """
    Calculate maximum minutes of content that can be extracted with proper spacing
    
    Args:
        audio_duration_seconds: Total audio duration in seconds
        slice_duration: Duration of each slice in seconds (default: 30)
        min_spacing: Minimum spacing between slice centers in seconds (default: 60)
    
    Returns:
        Maximum minutes that can be extracted
    """
    # Each slice takes (min_spacing) seconds of "space" in the audio
    max_slices = int(audio_duration_seconds / min_spacing)
    # Convert slices to minutes (each slice = 0.5 minutes)
    max_minutes = (max_slices * slice_duration) / 60
    return max_minutes

def generate_balanced_random_slices(audio_duration_seconds, total_minutes, slice_duration=30, min_spacing=60):
    """
    Generate random slices with 50/50 m/v split and proper spacing
    
    Args:
        audio_duration_seconds: Audio duration in seconds
        total_minutes: Total minutes of content requested by user
        slice_duration: Duration of each slice (default: 30s)
        min_spacing: Minimum spacing between slices (default: 60s)
    
    Returns:
        List of slice dictionaries in the same format as parse_audio_txt()
    """
    # Calculate number of slices needed
    num_slices = int(total_minutes * 2)  # 2 slices per minute (each 30s)
    
    # Ensure even number for 50/50 split
    if num_slices % 2 != 0:
        num_slices += 1
    
    num_m = num_v = num_slices // 2
    
    # Calculate safe bounds
    buffer = slice_duration / 2
    safe_start = buffer
    safe_end = audio_duration_seconds - buffer
    
    slices = []
    used_positions = []
    
    # Generate music slices (m)
    for i in range(num_m):
        slice_info = _generate_slice_with_spacing(
            safe_start, safe_end, used_positions, min_spacing, 
            'm', f"Random music segment {i+1}", slice_duration
        )
        if slice_info:
            slices.append(slice_info)
    
    # Generate voice slices (v)
    for i in range(num_v):
        slice_info = _generate_slice_with_spacing(
            safe_start, safe_end, used_positions, min_spacing,
            'v', f"Random voice segment {i+1}", slice_duration
        )
        if slice_info:
            slices.append(slice_info)
    
    # Sort by time
    slices.sort(key=lambda x: x['climax_time'])
    return slices

def _generate_slice_with_spacing(safe_start, safe_end, used_positions, min_spacing, slice_type, description, slice_duration):
    """
    Helper function to generate a single slice with proper spacing
    """
    max_attempts = 100
    
    for attempt in range(max_attempts):
        center = random.uniform(safe_start, safe_end)
        
        # Check if this position respects minimum spacing
        too_close = any(abs(center - pos) < min_spacing for pos in used_positions)
        
        if not too_close:
            used_positions.append(center)
            return {
                'climax_time': center,
                'type': slice_type,
                'description': description,
                'slice_begin': center - (slice_duration / 2),
                'slice_end': center + (slice_duration / 2)
            }
    
    # If we can't find a valid position after max attempts
    print(f"{Fore.YELLOW}⚠️  Could not find valid position for {description} after {max_attempts} attempts{Style.RESET_ALL}")
    return None

def generate_random_slices_and_sequence():
    """
    Main Option 3 → Option 2 workflow: 
    Audio file → Generate random slices → Slice → Sequence
    """
    print(f"{Fore.CYAN}=== Option 3 → Option 2: Random Slice & Sequence ==={Style.RESET_ALL}")
    print(f"{Fore.BLUE}This will:{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Generate random slices from your audio{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Create balanced music/voice content{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Automatically sequence the slices{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Apply professional audio processing{Style.RESET_ALL}")
    print()
    
    # 1. Select audio file
    print(f"{Fore.BLUE}Step 1: Select the source audio file...{Style.RESET_ALL}")
    audio_file = select_audio_file()
    
    if not audio_file:
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return
    
    # 2. Load audio to get duration
    print(f"{Fore.BLUE}Loading audio file...{Style.RESET_ALL}")
    try:
        audio = AudioSegment.from_file(audio_file)
        audio_duration_seconds = len(audio) / 1000
        audio_duration_minutes = audio_duration_seconds / 60
        
        print(f"{Fore.GREEN}✅ Audio loaded: {audio_duration_minutes:.1f} minutes ({audio_duration_seconds:.0f} seconds){Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}❌ Error loading audio file: {e}{Style.RESET_ALL}")
        return
    
    # 3. Calculate and show maximum possible minutes
    max_minutes = calculate_max_possible_minutes(audio_duration_seconds)
    print(f"{Fore.BLUE}Maximum content that can be extracted: {max_minutes:.1f} minutes{Style.RESET_ALL}")
    print()
    
    # 4. Ask user for desired minutes
    while True:
        try:
            user_input = input(f"{Fore.WHITE}How many minutes of sliced content do you want to generate? (max {max_minutes:.1f}): {Style.RESET_ALL}").strip()
            requested_minutes = float(user_input)
            
            if requested_minutes <= 0:
                print(f"{Fore.RED}❌ Please enter a positive number{Style.RESET_ALL}")
                continue
                
            if requested_minutes > max_minutes:
                print(f"{Fore.RED}❌ Cannot generate {requested_minutes:.1f} minutes. Maximum possible is {max_minutes:.1f} minutes{Style.RESET_ALL}")
                continue
                
            break
                
        except ValueError:
            print(f"{Fore.RED}❌ Please enter a valid number{Style.RESET_ALL}")
    
    # 5. Generate random slices
    print(f"{Fore.BLUE}Generating {requested_minutes:.1f} minutes of random slices...{Style.RESET_ALL}")
    slices = generate_balanced_random_slices(audio_duration_seconds, requested_minutes)
    
    if not slices:
        print(f"{Fore.RED}❌ Could not generate valid slices{Style.RESET_ALL}")
        return
    
    num_slices = len(slices)
    num_m = len([s for s in slices if s['type'] == 'm'])
    num_v = len([s for s in slices if s['type'] == 'v'])
    
    print(f"{Fore.GREEN}✅ Generated {num_slices} slices ({num_m} music, {num_v} voice){Style.RESET_ALL}")
    
    # 6. Select blocks directory for slicing
    print(f"{Fore.BLUE}Step 2: Select output folder for audio slices...{Style.RESET_ALL}")
    blocks_dir = select_output_folder()
    
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    # 7. Create temporary audio.txt file with generated slices
    import tempfile
    temp_txt_path = os.path.join(tempfile.gettempdir(), "random_slices_temp.txt")
    
    try:
        with open(temp_txt_path, 'w', encoding='utf-8') as f:
            for slice_info in slices:
                f.write(f"{slice_info['climax_time']:.2f}\t{slice_info['climax_time']:.2f}\t{slice_info['type']}\t{slice_info['description']}\n")
        
        print(f"{Fore.GREEN}✅ Created slice definitions{Style.RESET_ALL}")
        
        # 8. Run the slicing process (reusing your existing function)
        print(f"{Fore.BLUE}Step 3: Slicing audio...{Style.RESET_ALL}")
        
        # We need to temporarily modify the audio_file to use our temp txt file
        # Let's create a wrapper that uses our generated slices instead of file reading
        excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
        next_m, next_v = get_next_file_numbers(excel_path)
        
        print(f"{Fore.GREEN}Audio file: {audio_file}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}Output directory: {blocks_dir}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}Starting file numbers - m: {next_m}, v: {next_v}{Style.RESET_ALL}")
        print()
        
        # Process each slice using your existing logic
        print(f"{Fore.CYAN}Processing slices...{Style.RESET_ALL}")
        for slice_info in slices:
            # Determine file number based on type
            if slice_info['type'] == 'm':
                file_number = next_m
                next_m += 1
            elif slice_info['type'] == 'v':
                file_number = next_v
                next_v += 1
            else:
                continue
            
            # Process the slice as MP3 (reuse your existing function)
            output_path = process_audio_slice_mp3(audio, slice_info, blocks_dir, file_number)
            if output_path:
                # Update Excel file (reuse your existing function)
                update_excel_file(excel_path, slice_info, file_number, output_path, audio_file)
            print()
        
        # Verify files vs Excel database
        verify_files_vs_excel(blocks_dir, excel_path)
        
        print(f"{Fore.GREEN}✅ Audio slicing completed!{Style.RESET_ALL}")
        
        # 9. Automatically run sequencing
        print(f"{Fore.BLUE}Step 4: Sequencing slices...{Style.RESET_ALL}")
        
        # Scan for the blocks we just created
        m_blocks, v_blocks = scan_available_blocks(blocks_dir)
        
        if not validate_sequence_requirements(m_blocks, v_blocks):
            print(f"{Fore.YELLOW}⚠️  Not enough blocks for sequencing, but slicing completed successfully{Style.RESET_ALL}")
            return
        
        # Create random sequence
        m_sequence, v_sequence = create_random_sequence(m_blocks, v_blocks)
        
        # Build the audio sequence
        final_audio = build_multi_channel_sequence(blocks_dir, m_sequence, v_sequence)
        if not final_audio:
            print(f"{Fore.YELLOW}⚠️  Sequencing failed, but slicing completed successfully{Style.RESET_ALL}")
            return
        
        # 10. Let user choose sequence output location
        print(f"{Fore.BLUE}Step 5: Save final sequence...{Style.RESET_ALL}")
        output_path = filedialog.asksaveasfilename(
            title="Save Final Sequence As",
            defaultextension=".mp3",
            filetypes=[("MP3 files", "*.mp3"), ("All files", "*.*")]
        )
        
        if not output_path:
            print(f"{Fore.YELLOW}⚠️  No output file selected, but slicing completed successfully{Style.RESET_ALL}")
            return
        
        # Export final sequence
        final_audio.export(output_path, format="mp3", bitrate="192k")
        print(f"{Fore.GREEN}✅ Final sequence saved: {output_path}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}🎵 Final duration: {len(final_audio)/1000:.1f} seconds{Style.RESET_ALL}")
        
        # Generate timeline file
        audio_duration = len(final_audio) / 1000
        generate_sequence_timeline(output_path, blocks_dir, m_sequence, v_sequence, audio_duration)
        
        print(f"{Fore.CYAN}=== Option 3 → Option 2 Workflow Completed ==={Style.RESET_ALL}")
        print(f"{Fore.GREEN}🎉 Successfully created {requested_minutes:.1f} minutes of sequenced content!{Style.RESET_ALL}")
        
    finally:
        # Clean up temporary file
        if os.path.exists(temp_txt_path):
            os.remove(temp_txt_path)

def run_audio_slicer_with_labels():
    """Run audio slicer with existing label file"""
    print(f"{Fore.CYAN}=== Audio Slicer with Labels ==={Style.RESET_ALL}")
    
    # Select audio file
    print(f"{Fore.BLUE}Please select the audio file...{Style.RESET_ALL}")
    audio_file = select_audio_file()
    
    if not audio_file:
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return
    
    # Select output folder
    print(f"{Fore.BLUE}Please select output folder for slices...{Style.RESET_ALL}")
    blocks_dir = select_output_folder()
    
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    # Run the slicing process
    result = slice_audio_from_labels(audio_file, blocks_dir)
    if result:
        print(f"{Fore.GREEN}✅ Audio slicing completed successfully!{Style.RESET_ALL}")
    else:
        print(f"{Fore.RED}❌ Audio slicing failed.{Style.RESET_ALL}")

def main():
    choice = show_welcome_screen()
    
    if choice == '1':
        slice_choice = show_slice_options_menu()
        if slice_choice == '1':
            run_audio_slicer_with_labels()
        elif slice_choice == '2':
            no_labels_choice = show_no_labels_menu()
            if no_labels_choice == '1':
                print(f"{Fore.GREEN}🎯 Great! Please label your audio file in Audacity...{Style.RESET_ALL}")
                return
            elif no_labels_choice == '2':
                run_random_slicer()
                return
                
    elif choice == '2':
        run_sequencer()
        return
        
    elif choice == '3':
        sub_choice = show_slice_and_sequence_menu()
        if sub_choice == '1':
            run_slice_and_sequence_with_labels()
        elif sub_choice == '2':
            generate_random_slices_and_sequence()  # NEW: Replace the placeholder
        return
        
    elif choice == '4':
        show_help()
        main()

if __name__ == "__main__":
    main()