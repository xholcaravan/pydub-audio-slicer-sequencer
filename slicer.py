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
# Initialize colorama (this makes colors work on Windows too)
init()
import random

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
                    print(f"{Fore.YELLOW}⚠️  Warning: Line {line_num} has invalid format: {line}{Style.RESET_ALL}")
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
{Fore.CYAN}║                 pydub-audio-slicer-sequencer                 ║{Style.RESET_ALL}
{Fore.CYAN}║                     Audio Processing Tool                    ║{Style.RESET_ALL}
{Fore.CYAN}╚══════════════════════════════════════════════════════════════╝{Style.RESET_ALL}

{Fore.WHITE}Use this tool to:{Style.RESET_ALL}
{Fore.GREEN}1 - Slice an audio file into several blocks{Style.RESET_ALL}
{Fore.BLUE}2 - Sequence blocks to create an audio file{Style.RESET_ALL}  
{Fore.MAGENTA}3 - Slice an audio file and produce a sequence{Style.RESET_ALL}

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
"""
    print(welcome_text)
    
    while True:
        choice = input(f"{Fore.WHITE}Select option (1, 2, or 3): {Style.RESET_ALL}").strip()
        if choice in ['1', '2', '3']:
            return choice
        else:
            print(f"{Fore.RED}❌ Invalid choice. Please enter 1, 2, or 3.{Style.RESET_ALL}")

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

def run_audio_slicer_with_labels():
    """Run the audio slicing functionality"""
    print(f"{Fore.CYAN}=== Audio Slicer Started ==={Style.RESET_ALL}")
    print(f"Slice size: {SLICE_SIZE} seconds")
    print(f"Fade duration: {FADE_DURATION} seconds")
    print()
    
    # Select audio file
    print(f"{Fore.BLUE}Please select the audio file to slice...{Style.RESET_ALL}")
    audio_file = select_audio_file()
    
    # Get corresponding txt file
    txt_file = get_corresponding_txt_file(audio_file)
    
    # Verify files exist
    if not verify_files_exist(audio_file, txt_file):
        return
    
    # Select output folder
    print(f"{Fore.BLUE}Please select output folder for slices...{Style.RESET_ALL}")
    blocks_dir = select_output_folder()
    
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    print(f"{Fore.GREEN}Audio file: {audio_file}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Text file: {txt_file}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Output directory: {blocks_dir}{Style.RESET_ALL}")
    print()
    
    # Parse audio.txt
    print(f"{Fore.BLUE}Parsing audio.txt...{Style.RESET_ALL}")
    slices = parse_audio_txt(txt_file)
    if not slices:
        print(f"{Fore.YELLOW}⚠️  No valid slices found in audio.txt{Style.RESET_ALL}")
        return
    
    print(f"{Fore.GREEN}Found {len(slices)} slices to process{Style.RESET_ALL}")
    for i, slice_info in enumerate(slices, 1):
        print(f"  {i}. {slice_info['type']} at {slice_info['climax_time']}s: {slice_info['description']}")
    print()
    
    # Get next file numbers from Excel
    next_m, next_v = get_next_file_numbers(excel_path)
    print(f"{Fore.BLUE}Next file numbers - m: {next_m}, v: {next_v}{Style.RESET_ALL}")
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
        
        # Process the slice
        output_path = process_audio_slice(audio, slice_info, blocks_dir, file_number)
        if output_path:
            # Update Excel file
            update_excel_file(excel_path, slice_info, file_number, output_path, audio_file)
        print()
    
    # Verify files vs Excel database
    verify_files_vs_excel(blocks_dir, excel_path)
    
    print(f"{Fore.CYAN}=== Audio Slicer Completed ==={Style.RESET_ALL}")

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
        
        # Rest of the function...
        
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
   
def main():
    choice = show_welcome_screen()
    
    if choice == '1':
        slice_choice = show_slice_options_menu()
        if slice_choice == '1':
            run_audio_slicer_with_labels()  # Your existing label-based slicer
        elif slice_choice == '2':
            no_labels_choice = show_no_labels_menu()
            if no_labels_choice == '1':
                print(f"{Fore.GREEN}🎯 Great! Please label your audio file in Audacity and then run this program again.{Style.RESET_ALL}")
                print(f"{Fore.GREEN}   Remember to export labels as a .txt file with the same name as your audio file.{Style.RESET_ALL}")
                return
            elif no_labels_choice == '2':
                run_random_slicer()  # New random slicer!
                return
    # ... rest of main function
    #    

if __name__ == "__main__":
    main()