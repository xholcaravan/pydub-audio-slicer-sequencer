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
import eyed3
from eyed3.id3.frames import ImageFrame
import json
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

# ============================================================================
# FOLDER MEMORY SYSTEM
# ============================================================================

def load_settings():
    """Load settings from settings.json in script directory"""
    settings_path = os.path.join(get_base_path(), 'slicer_settings.json')
    
    if os.path.exists(settings_path):
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                last_parent_folder = settings.get('last_parent_folder', '')
                print(f"{Fore.BLUE}📁 Loaded last parent folder: {last_parent_folder}{Style.RESET_ALL}")
                return last_parent_folder
        except Exception as e:
            print(f"{Fore.YELLOW}⚠️ Could not load settings: {e}{Style.RESET_ALL}")
            return ""
    else:
        print(f"{Fore.BLUE}⚙️ First run - no settings file found{Style.RESET_ALL}")
        return ""

def save_settings(last_parent_folder):
    """Save last parent folder to settings.json in script directory"""
    settings_path = os.path.join(get_base_path(), 'slicer_settings.json')
    
    try:
        settings = {
            'last_parent_folder': last_parent_folder
        }
        
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2)
        
        print(f"{Fore.BLUE}💾 Saved parent folder: {last_parent_folder}{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.YELLOW}⚠️ Could not save settings: {e}{Style.RESET_ALL}")

def update_parent_folder(selected_path):
    """Update the parent folder in settings based on selected file/folder"""
    if selected_path:
        parent_folder = os.path.dirname(selected_path)
        save_settings(parent_folder)
        return parent_folder
    return None

def get_initial_directory():
    """Get the initial directory for file/folder dialogs"""
    last_parent = load_settings()
    if last_parent and os.path.exists(last_parent):
        return last_parent
    return os.path.expanduser("~")  # Fallback to home directory

# ============================================================================
# MODIFIED DIALOG FUNCTIONS WITH MEMORY
# ============================================================================

def select_audio_file():
    """Let user select audio file and return its path"""
    root = tk.Tk()
    root.withdraw()
    
    initial_dir = get_initial_directory()
    
    audio_file = filedialog.askopenfilename(
        title="Select Audio File",
        initialdir=initial_dir,
        filetypes=[
            ("Audio files", "*.wav *.mp3 *.flac *.aiff *.aac *.ogg *.m4a"),
            ("All files", "*.*")
        ]
    )
    root.destroy()
    
    if audio_file:
        update_parent_folder(audio_file)
    
    return audio_file

def select_output_folder():
    """Let user select output folder for slices"""
    root = tk.Tk()
    root.withdraw()
    
    initial_dir = get_initial_directory()
    
    output_folder = filedialog.askdirectory(
        title="Select Output Folder for Slices",
        initialdir=initial_dir
    )
    root.destroy()
    
    if output_folder:
        update_parent_folder(output_folder)
    
    return output_folder

def select_blocks_folder():
    """Let user select blocks folder for sequencing"""
    root = tk.Tk()
    root.withdraw()
    
    initial_dir = get_initial_directory()
    
    blocks_folder = filedialog.askdirectory(
        title="Select Blocks Folder",
        initialdir=initial_dir
    )
    root.destroy()
    
    if blocks_folder:
        update_parent_folder(blocks_folder)
    
    return blocks_folder

def ask_save_file(default_ext=".mp3", filetypes=[("MP3 files", "*.mp3"), ("All files", "*.*")]):
    """Ask user to save a file"""
    root = tk.Tk()
    root.withdraw()
    
    initial_dir = get_initial_directory()
    
    file_path = filedialog.asksaveasfilename(
        title="Save As",
        defaultextension=default_ext,
        initialdir=initial_dir,
        filetypes=filetypes
    )
    root.destroy()
    
    if file_path:
        update_parent_folder(file_path)
    
    return file_path

# ============================================================================
# ORIGINAL FUNCTIONS (with updated dialog calls)
# ============================================================================

# Hardcoded parameters
SLICE_SIZE = 30  # seconds
FADE_DURATION = SLICE_SIZE / 2  # seconds

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
                    
                    # Validate audio type
                    if audio_type not in ['m', 'v', 'j']:
                        print(f"{Fore.YELLOW}⚠️  Warning: Line {line_num} has unknown audio type '{audio_type}'. Skipping.{Style.RESET_ALL}")
                        continue
                    
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

def update_excel_file(excel_path, slice_info, timestamp_id, origin_file):
    """Update the Excel file with new slice information"""
    try:
        # Use the provided timestamp ID (same one used for filename)
        unique_filename = f"{slice_info['type']}{timestamp_id}"
        
        # Create DataFrames for new entries with correct column order
        new_data = {
            slice_info['type']: [unique_filename],
            'origin': [origin_file],
            'description': [slice_info['description']]
        }
        new_df = pd.DataFrame(new_data)
        
        # Try to read existing file or create new one
        try:
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                sheet_name = slice_info['type']  # 'm', 'v', or 'j'
                
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
                # Create empty sheets for the other types
                other_types = ['v', 'j'] if slice_info['type'] == 'm' else ['m', 'j'] if slice_info['type'] == 'v' else ['m', 'v']
                for other_type in other_types:
                    pd.DataFrame(columns=[other_type, 'origin', 'description']).to_excel(writer, sheet_name=other_type, index=False)
                
        print(f"{Fore.GREEN}✅ Updated Excel: {unique_filename}{Style.RESET_ALL}")
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error updating Excel file: {e}{Style.RESET_ALL}")
        
def verify_files_vs_excel(blocks_dir, excel_path):
    """Verify that files in blocks folder match the Excel database"""
    print(f"\n{Fore.CYAN}=== Verifying Files vs Excel Database ==={Style.RESET_ALL}")
    
    try:
        # Read Excel sheets
        m_df = pd.read_excel(excel_path, sheet_name='m')
        v_df = pd.read_excel(excel_path, sheet_name='v')
        j_df = pd.read_excel(excel_path, sheet_name='j')
        
        # Get all files in blocks directory
        all_files = os.listdir(blocks_dir)
        audio_files = [f for f in all_files if f.endswith('.wav') or f.endswith('.mp3')]
        
        # Extract m, v, and j files from folder
        m_files_folder = [f for f in audio_files if f.startswith('m')]
        v_files_folder = [f for f in audio_files if f.startswith('v')]
        j_files_folder = [f for f in audio_files if f.startswith('j')]
        
        # Get file names from Excel
        m_files_excel = []
        if not m_df.empty and 'm' in m_df.columns:
            m_files_excel = [f"{row['m']}.mp3" for _, row in m_df.iterrows() if pd.notna(row['m'])]
        
        v_files_excel = []
        if not v_df.empty and 'v' in v_df.columns:
            v_files_excel = [f"{row['v']}.mp3" for _, row in v_df.iterrows() if pd.notna(row['v'])]
        
        j_files_excel = []
        if not j_df.empty and 'j' in j_df.columns:
            j_files_excel = [f"{row['j']}.mp3" for _, row in j_df.iterrows() if pd.notna(row['j'])]
        
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
        
        # Compare Jingles files (j)
        print(f"\n{Fore.CYAN}--- Jingles Files (j) ---{Style.RESET_ALL}")
        j_folder_set = set(j_files_folder)
        j_excel_set = set(j_files_excel)
        
        missing_in_folder = j_excel_set - j_folder_set
        missing_in_excel = j_folder_set - j_excel_set
        
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
        
        print(f"Total in Excel: {len(j_files_excel)}, Total in folder: {len(j_files_folder)}")
        
        # Summary
        print(f"\n{Fore.CYAN}--- Summary ---{Style.RESET_ALL}")
        total_excel = len(m_files_excel) + len(v_files_excel) + len(j_files_excel)
        total_folder = len(m_files_folder) + len(v_files_folder) + len(j_files_folder)
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

def check_for_corrupted_files(blocks_dir, file_list):
    """Check if any files in the list are corrupted and return valid files"""
    valid_files = []
    problematic_files = []
    
    for filename in file_list:
        file_path = os.path.join(blocks_dir, filename)
        if not os.path.exists(file_path):
            problematic_files.append((filename, "File not found"))
            continue
            
        try:
            # Method 1: Try with pydub using specific codec
            try:
                audio = AudioSegment.from_file(file_path, format="mp3")
                if len(audio) == 0:
                    problematic_files.append((filename, "Empty audio file (pydub)"))
                else:
                    valid_files.append(filename)
                    continue
            except Exception as e1:
                # Method 2: Try with different format parameter
                try:
                    audio = AudioSegment.from_file(file_path)
                    if len(audio) == 0:
                        problematic_files.append((filename, "Empty audio file (pydub auto)"))
                    else:
                        valid_files.append(filename)
                        continue
                except Exception as e2:
                    # Method 3: Try using eyed3 to check if it's a valid MP3
                    try:
                        audiofile = eyed3.load(file_path)
                        if audiofile is None:
                            problematic_files.append((filename, "Not a valid MP3 file"))
                        elif audiofile.info is None:
                            problematic_files.append((filename, "MP3 file has no audio info"))
                        else:
                            # File seems valid but pydub can't read it - check for false video detection
                            if _is_false_video_detection(file_path):
                                problematic_files.append((filename, "False video detection by FFmpeg"))
                            else:
                                problematic_files.append((filename, f"Pydub incompatible: {str(e1)}"))
                    except Exception as e3:
                        problematic_files.append((filename, f"All methods failed: pydub1:{e1}, pydub2:{e2}, eyed3:{e3}"))
                        
        except Exception as e:
            problematic_files.append((filename, f"Unexpected error: {e}"))
    
    return valid_files, problematic_files

def _is_false_video_detection(file_path):
    """Check if FFmpeg is falsely detecting audio as video"""
    try:
        import subprocess
        # Use ffprobe to check what FFmpeg thinks the file is
        cmd = ['ffprobe', '-v', 'quiet', '-print_format', 'json', '-show_streams', file_path]
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            import json
            probe_data = json.loads(result.stdout)
            
            if 'streams' in probe_data:
                for stream in probe_data['streams']:
                    if 'codec_type' in stream:
                        # If FFmpeg detects video in an MP3 file, it's a false positive
                        if stream['codec_type'] == 'video':
                            codec_name = stream.get('codec_name', 'unknown')
                            return f"False video detection (codec: {codec_name})"
            
            # Check if there are no audio streams
            audio_streams = [s for s in probe_data.get('streams', []) if s.get('codec_type') == 'audio']
            if not audio_streams:
                return "No audio streams detected"
                
        return None
    except Exception as e:
        return f"Error checking file: {e}"

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
{Fore.CYAN}4 - Advanced options (Excel management, verification){Style.RESET_ALL}
{Fore.YELLOW}5 - Help!{Style.RESET_ALL}
"""
    print(welcome_text)
    
    while True:
        choice = input(f"{Fore.WHITE}Select option (1-5): {Style.RESET_ALL}").strip()
        if choice in ['1', '2', '3', '4', '5']:
            return choice
        else:
            print(f"{Fore.RED}❌ Invalid choice. Please enter 1-5.{Style.RESET_ALL}")

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
    txt_file = get_corresponding_txt_file(audio_file)
    
    if not verify_files_exist(audio_file, txt_file):
        return None
    
    # ADD THIS LINE:
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    print(f"{Fore.GREEN}Audio file: {audio_file}{Style.RESET_ALL}")
    try:
        audio = AudioSegment.from_file(audio_file)
        audio_duration = len(audio) / 1000
        print(f"{Fore.GREEN}✅ Audio loaded: {audio_duration:.2f} seconds{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}❌ Error loading audio file: {e}{Style.RESET_ALL}")
        return None
    
    # Parse audio.txt with duration checking
    print(f"{Fore.BLUE}Parsing audio.txt...{Style.RESET_ALL}")
    slices = parse_audio_txt(txt_file, audio_duration)
    
    if not slices:
        print(f"{Fore.YELLOW}⚠️  No valid slices found in audio.txt{Style.RESET_ALL}")
        return None
    
    print(f"{Fore.GREEN}Found {len(slices)} slices to process{Style.RESET_ALL}")
    for i, slice_info in enumerate(slices, 1):
        print(f"  {i}. {slice_info['type']} at {slice_info['climax_time']}s: {slice_info['description']}")
    print()
    
    # Process each slice
    print(f"\n{Fore.CYAN}Processing slices...{Style.RESET_ALL}")
    for slice_info in slices:       
        # Process the slice as MP3
        output_path = process_audio_slice_mp3(audio, slice_info, blocks_dir, audio_file)
        print()
    
    # Verify files vs Excel database
    verify_files_vs_excel(blocks_dir, excel_path)
    
    print(f"{Fore.GREEN}✅ Audio slicing completed!{Style.RESET_ALL}")
    return blocks_dir

def diagnose_problematic_file(file_path):
    """Diagnose why a file can't be loaded and suggest fixes"""
    print(f"{Fore.CYAN}🔍 Diagnosing problematic file: {os.path.basename(file_path)}{Style.RESET_ALL}")
    
    if not os.path.exists(file_path):
        print(f"{Fore.RED}❌ File does not exist{Style.RESET_ALL}")
        return False
    
    file_size = os.path.getsize(file_path)
    print(f"{Fore.BLUE}   File size: {file_size} bytes{Style.RESET_ALL}")
    
    if file_size == 0:
        print(f"{Fore.RED}❌ File is empty (0 bytes){Style.RESET_ALL}")
        return False
    
    # Try to get basic file info
    try:
        import subprocess
        result = subprocess.run(['file', file_path], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"{Fore.BLUE}   File type: {result.stdout.strip()}{Style.RESET_ALL}")
    except:
        pass
    
    # Check if it's actually an MP3
    try:
        audiofile = eyed3.load(file_path)
        if audiofile:
            print(f"{Fore.GREEN}   ✅ File is recognized as MP3 by eyed3{Style.RESET_ALL}")
            if audiofile.tag:
                print(f"{Fore.BLUE}   📝 Has ID3 tags{Style.RESET_ALL}")
            if audiofile.info:
                print(f"{Fore.BLUE}   🎵 Duration: {audiofile.info.time_secs:.2f}s, Bitrate: {audiofile.info.bit_rate[1]} kbps{Style.RESET_ALL}")
            else:
                print(f"{Fore.YELLOW}   ⚠️  No audio info available{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}   ❌ Not recognized as MP3 by eyed3{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}   ❌ Error with eyed3: {e}{Style.RESET_ALL}")
    
    # Try alternative loading methods
    print(f"{Fore.BLUE}   Testing alternative loading methods...{Style.RESET_ALL}")
    
    # Method 1: Try with explicit codec
    try:
        audio = AudioSegment.from_file(file_path, format="mp3")
        print(f"{Fore.GREEN}   ✅ Loads with format='mp3'{Style.RESET_ALL}")
        return True
    except Exception as e:
        print(f"{Fore.YELLOW}   ⚠️  Fails with format='mp3': {e}{Style.RESET_ALL}")
    
    # Method 2: Try without format
    try:
        audio = AudioSegment.from_file(file_path)
        print(f"{Fore.GREEN}   ✅ Loads without format parameter{Style.RESET_ALL}")
        return True
    except Exception as e:
        print(f"{Fore.YELLOW}   ⚠️  Fails without format: {e}{Style.RESET_ALL}")
    
    # Method 3: Try with ffmpeg directly
    try:
        import subprocess
        # Test if ffmpeg can read the file
        cmd = ['ffmpeg', '-i', file_path, '-f', 'null', '-']
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"{Fore.GREEN}   ✅ FFmpeg can read the file{Style.RESET_ALL}")
            return True
        else:
            print(f"{Fore.RED}   ❌ FFmpeg cannot read the file{Style.RESET_ALL}")
            print(f"{Fore.RED}   FFmpeg error: {result.stderr}{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.YELLOW}   ⚠️  FFmpeg test failed: {e}{Style.RESET_ALL}")
    
    print(f"{Fore.YELLOW}💡 Suggestion: Try re-encoding the file with Audacity or another audio editor{Style.RESET_ALL}")
    return False

def fix_problematic_file(file_path):
    """Attempt to fix a problematic MP3 file by re-encoding it"""
    try:
        print(f"{Fore.BLUE}🛠️  Attempting to fix: {os.path.basename(file_path)}{Style.RESET_ALL}")
        
        # Create a temporary file
        import tempfile
        temp_dir = tempfile.gettempdir()
        temp_output = os.path.join(temp_dir, f"fixed_{os.path.basename(file_path)}")
        
        # Use FFmpeg to re-encode the file
        import subprocess
        cmd = [
            'ffmpeg', '-y', '-i', file_path,
            '-c:a', 'libmp3lame', '-b:a', '192k',
            '-map_metadata', '0',  # Copy metadata
            temp_output
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0 and os.path.exists(temp_output):
            # Replace the original file with the fixed one
            import shutil
            shutil.move(temp_output, file_path)
            print(f"{Fore.GREEN}✅ Successfully fixed: {os.path.basename(file_path)}{Style.RESET_ALL}")
            return True
        else:
            print(f"{Fore.RED}❌ Failed to fix: {os.path.basename(file_path)}{Style.RESET_ALL}")
            if os.path.exists(temp_output):
                os.remove(temp_output)
            return False
            
    except Exception as e:
        print(f"{Fore.RED}❌ Error fixing file: {e}{Style.RESET_ALL}")
        return False

def calculate_slice_density(audio_duration_seconds):
    """Calculate number of slices based on audio duration (~1 per 2 minutes)"""
    base_slices = audio_duration_seconds / 120
    variation = random.uniform(0.8, 1.2)
    num_slices = max(1, int(base_slices * variation))
    return num_slices

def generate_random_labels(audio_file):
    """Generate random slice positions throughout the audio file with proper density"""
    try:
        print(f"{Fore.BLUE}Loading audio to calculate duration...{Style.RESET_ALL}")
        audio = AudioSegment.from_file(audio_file)
        duration_seconds = len(audio) / 1000
        print(f"{Fore.GREEN}✅ Audio duration: {duration_seconds:.1f} seconds{Style.RESET_ALL}")
        
        num_slices = calculate_slice_density(duration_seconds)
        print(f"{Fore.BLUE}Calculated {num_slices} slices for {duration_seconds:.1f}s audio{Style.RESET_ALL}")
        
        max_start_time = duration_seconds - SLICE_SIZE
        if max_start_time <= 0:
            print(f"{Fore.RED}❌ Audio file is too short ({duration_seconds:.1f}s) for {SLICE_SIZE}s slices{Style.RESET_ALL}")
            return None
        
        slices = []
        for i in range(num_slices):
            min_spacing = SLICE_SIZE * 1.5
            max_attempts = 100
            attempt = 0
            
            while attempt < max_attempts:
                start_time = random.uniform(0, max_start_time)
                climax_time = start_time + (SLICE_SIZE / 2)
                
                overlap = False
                for existing_slice in slices:
                    if abs(existing_slice['climax_time'] - climax_time) < min_spacing:
                        overlap = True
                        break
                
                if not overlap:
                    break
                attempt += 1
            
            audio_type = random.choice(['m', 'v', 'j'])
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

def process_audio_slice_mp3(audio, slice_info, output_folder, origin_file):
    """Process a single audio slice and export as MP3 192kbps with metadata"""
    try:
        begin_ms = int(slice_info['slice_begin'] * 1000)
        end_ms = int(slice_info['slice_end'] * 1000)
        
        begin_ms = max(0, begin_ms)
        end_ms = min(len(audio), end_ms)
        
        slice_audio = audio[begin_ms:end_ms]
        
        fade_duration_ms = int(FADE_DURATION * 1000)
        slice_audio = slice_audio.fade_in(fade_duration_ms).fade_out(fade_duration_ms)
        
        slice_audio = normalize(slice_audio)
        
        timestamp_id = generate_timestamp_id()
        filename = f"{slice_info['type']}{timestamp_id}.mp3"
        output_path = os.path.join(output_folder, filename)
        
        slice_audio.export(output_path, format="mp3", bitrate="192k")
        
        metadata_success = write_audio_metadata(
            output_path, 
            origin_file, 
            slice_info['description'], 
            slice_info['type'], 
            slice_info['climax_time']
        )
        
        if metadata_success:
            print(f"{Fore.GREEN}✅ Successfully created: {filename} (with metadata){Style.RESET_ALL}")
        else:
            print(f"{Fore.GREEN}✅ Successfully created: {filename} (metadata failed){Style.RESET_ALL}")
        
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
    
    audio_file = select_audio_file()
    if not audio_file:
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return
    
    slices = generate_random_labels(audio_file)
    if not slices:
        return
    
    blocks_dir = select_output_folder()
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
    
    print(f"{Fore.GREEN}Audio file: {audio_file}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}Label source: Randomly generated ({len(slices)} slices){Style.RESET_ALL}")
    print(f"{Fore.GREEN}Output directory: {blocks_dir}{Style.RESET_ALL}")
    print()
    
    print(f"{Fore.GREEN}Found {len(slices)} slices to process{Style.RESET_ALL}")
    for i, slice_info in enumerate(slices, 1):
        print(f"  {i}. {slice_info['type']} at {slice_info['climax_time']:.1f}s: {slice_info['description']}")
    print()
    
    print(f"{Fore.BLUE}Loading audio file...{Style.RESET_ALL}")
    try:
        audio = AudioSegment.from_file(audio_file)
        print(f"{Fore.GREEN}✅ Audio loaded: {len(audio)/1000:.2f} seconds{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}❌ Error loading audio file: {e}{Style.RESET_ALL}")
        return
    
    print(f"\n{Fore.CYAN}Processing slices...{Style.RESET_ALL}")
    for slice_info in slices:
        
        output_path = process_audio_slice_mp3(audio, slice_info, blocks_dir, audio_file)
        print()
    
    verify_files_vs_excel(blocks_dir, excel_path)
    print(f"{Fore.CYAN}=== Random Audio Slicer Completed ==={Style.RESET_ALL}")

def scan_available_blocks(blocks_dir):
    """Scan blocks directory for m, v, and j audio files"""
    if not os.path.exists(blocks_dir):
        return [], [], []
    
    all_files = os.listdir(blocks_dir)
    m_blocks = [f for f in all_files if f.startswith('m') and (f.endswith('.mp3') or f.endswith('.wav'))]
    v_blocks = [f for f in all_files if f.startswith('v') and (f.endswith('.mp3') or f.endswith('.wav'))]
    j_blocks = [f for f in all_files if f.startswith('j') and (f.endswith('.mp3') or f.endswith('.wav'))]
    
    # Sort by number for consistent ordering before shuffling
    m_blocks.sort(key=lambda x: int(x[1:].split('.')[0]) if x[1:].split('.')[0].isdigit() else 0)
    v_blocks.sort(key=lambda x: int(x[1:].split('.')[0]) if x[1:].split('.')[0].isdigit() else 0)
    j_blocks.sort(key=lambda x: int(x[1:].split('.')[0]) if x[1:].split('.')[0].isdigit() else 0)
    
    return m_blocks, v_blocks, j_blocks

def validate_sequence_requirements(m_blocks, v_blocks, j_blocks):
    """Validate that we have enough blocks for sequencing"""
    total_voice_jingle = len(v_blocks) + len(j_blocks)
    
    if len(m_blocks) < 3:
        print(f"{Fore.RED}❌ Not enough music blocks: {len(m_blocks)} found (minimum 3 required){Style.RESET_ALL}")
        return False
    if total_voice_jingle < 3:
        print(f"{Fore.RED}❌ Not enough voice+jingle blocks: {total_voice_jingle} found (minimum 3 required){Style.RESET_ALL}")
        return False
    
    print(f"{Fore.GREEN}✅ Found {len(m_blocks)} music blocks and {total_voice_jingle} voice+jingle blocks{Style.RESET_ALL}")
    return True

def create_voice_sequence(v_blocks, j_blocks):
    """Create a voice sequence that mixes v and j blocks, starting with a jingle if available"""
    all_voice_blocks = v_blocks + j_blocks
    
    if not all_voice_blocks:
        return []
    
    # Separate jingles for potential first position
    jingles = [b for b in all_voice_blocks if b.startswith('j')]
    voice_only = [b for b in all_voice_blocks if b.startswith('v')]
    
    # Start with a jingle if available
    if jingles:
        first_block = random.choice(jingles)
        jingles.remove(first_block)
        remaining_blocks = voice_only + jingles
        random.shuffle(remaining_blocks)
        voice_sequence = [first_block] + remaining_blocks
    else:
        # No jingles, just shuffle all voice blocks
        voice_sequence = voice_only.copy()
        random.shuffle(voice_sequence)
    
    return voice_sequence

def create_random_sequence(m_blocks, v_blocks, j_blocks):
    """Create random sequences for music and mixed voice channels"""
    # Shuffle music blocks randomly
    random.shuffle(m_blocks)
    
    # Create mixed voice sequence
    voice_sequence = create_voice_sequence(v_blocks, j_blocks)
    
    # Use the minimum length to determine sequence duration
    sequence_length = min(len(m_blocks), len(voice_sequence))
    
    # Take only the blocks we'll actually use
    m_sequence = m_blocks[:sequence_length]
    voice_sequence_trimmed = voice_sequence[:sequence_length]  # Trim to match length
    
    # Show which blocks are used vs skipped
    print(f"{Fore.BLUE}🎵 Sequence Configuration:{Style.RESET_ALL}")
    print(f"{Fore.GREEN}   Using {sequence_length} blocks from each channel{Style.RESET_ALL}")
    
    if len(m_blocks) > sequence_length:
        print(f"{Fore.YELLOW}   Skipping {len(m_blocks) - sequence_length} music blocks{Style.RESET_ALL}")
    if len(v_blocks) + len(j_blocks) > sequence_length:
        print(f"{Fore.YELLOW}   Skipping {len(v_blocks) + len(j_blocks) - sequence_length} voice+jingle blocks{Style.RESET_ALL}")
    
    return m_sequence, voice_sequence_trimmed

def build_multi_channel_sequence(blocks_dir, m_sequence, voice_sequence):
    """Build the final sequence with 15-second music channel offset"""
    try:
        print(f"{Fore.BLUE}🔊 Building audio sequence...{Style.RESET_ALL}")
        
        # Validate sequences have the same length
        if len(m_sequence) != len(voice_sequence):
            print(f"{Fore.RED}❌ Error: Music sequence ({len(m_sequence)}) and voice sequence ({len(voice_sequence)}) have different lengths{Style.RESET_ALL}")
            return None
        
        # Create 15 seconds of silence for music channel offset
        silence_15s = AudioSegment.silent(duration=15000)
        
        # Initialize channels - voice starts immediately, music starts after 15s
        music_channel = silence_15s
        voice_channel = AudioSegment.empty()
        
        # Load and concatenate music blocks
        print(f"{Fore.BLUE}   Loading music channel...{Style.RESET_ALL}")
        for i, block in enumerate(m_sequence, 1):
            block_path = os.path.join(blocks_dir, block)
            if not os.path.exists(block_path):
                print(f"{Fore.RED}❌ Music block not found: {block}{Style.RESET_ALL}")
                return None
            audio_segment = AudioSegment.from_file(block_path)
            music_channel += audio_segment
            print(f"{Fore.GREEN}     [{i}/{len(m_sequence)}] Added: {block}{Style.RESET_ALL}")
        
        # Load and concatenate voice blocks (mixed v and j)
        print(f"{Fore.BLUE}   Loading voice channel...{Style.RESET_ALL}")
        for i, block in enumerate(voice_sequence, 1):
            block_path = os.path.join(blocks_dir, block)
            if not os.path.exists(block_path):
                print(f"{Fore.RED}❌ Voice block not found: {block}{Style.RESET_ALL}")
                return None
            audio_segment = AudioSegment.from_file(block_path)
            voice_channel += audio_segment
            block_type = "JINGLE" if block.startswith('j') else "VOICE"
            print(f"{Fore.GREEN}     [{i}/{len(voice_sequence)}] Added: {block} ({block_type}){Style.RESET_ALL}")
        
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
        import traceback
        traceback.print_exc()
        return None

def run_sequencer():
    """Main sequencing workflow - Option 2"""
    print(f"{Fore.CYAN}=== Audio Sequencer Started ==={Style.RESET_ALL}")
    print(f"{Fore.BLUE}This will create a mixed sequence with:{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Music channel starting at 0:15{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Voice channel starting at 0:00 (mixed voice and jingles){Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Random block order{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Stereo output{Style.RESET_ALL}")
    print()
    
    blocks_dir = select_blocks_folder()
    if not blocks_dir:
        print(f"{Fore.RED}❌ No blocks folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    # Scan available blocks and calculate maximum minutes
    print(f"{Fore.BLUE}Scanning for audio blocks...{Style.RESET_ALL}")
    m_blocks, v_blocks, j_blocks = scan_available_blocks(blocks_dir)
    
    if not validate_sequence_requirements(m_blocks, v_blocks, j_blocks):
        return
    
    max_blocks = min(len(m_blocks), len(v_blocks) + len(j_blocks))
    max_minutes = (max_blocks * 30) / 60
    
    print(f"{Fore.GREEN}✅ Found {len(m_blocks)} music blocks and {len(v_blocks) + len(j_blocks)} voice+jingle blocks{Style.RESET_ALL}")
    print(f"{Fore.GREEN}📊 Maximum sequence: {max_minutes:.1f} minutes{Style.RESET_ALL}")
    print()
    
    while True:
        try:
            user_input = input(f"{Fore.WHITE}How many minutes? (Enter for max {max_minutes:.1f}): {Style.RESET_ALL}").strip()
            if not user_input:
                desired_minutes = None
                break
            desired_minutes = float(user_input)
            if desired_minutes <= 0:
                print(f"{Fore.RED}❌ Please enter a positive number{Style.RESET_ALL}")
                continue
            if desired_minutes > max_minutes:
                print(f"{Fore.RED}❌ Cannot create {desired_minutes:.1f} minutes. Maximum possible is {max_minutes:.1f} minutes{Style.RESET_ALL}")
                continue
            break
        except ValueError:
            print(f"{Fore.RED}❌ Please enter a valid number{Style.RESET_ALL}")
    
    success, final_audio, blocks_info = create_sequence_from_blocks(blocks_dir, desired_minutes)
    if not success:
        print(f"{Fore.RED}❌ Sequencing failed{Style.RESET_ALL}")
        return
    
    output_path = ask_save_file()
    
    if not output_path:
        print(f"{Fore.RED}❌ No output file selected. Exiting.{Style.RESET_ALL}")
        return
    
    try:
        print(f"{Fore.BLUE}Exporting sequence...{Style.RESET_ALL}")
        final_audio.export(output_path, format="mp3", bitrate="192k")
        print(f"{Fore.GREEN}✅ Sequence saved: {output_path}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}🎵 Final duration: {blocks_info['total_duration']:.1f} seconds{Style.RESET_ALL}")
        
        generate_sequence_timeline(output_path, blocks_dir, 
                                 blocks_info['m_sequence'], blocks_info['voice_sequence'], 
                                 blocks_info['total_duration'])
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error exporting sequence: {e}{Style.RESET_ALL}")
        return
    
    print(f"{Fore.CYAN}=== Audio Sequencer Completed ==={Style.RESET_ALL}")

def run_slice_and_sequence_with_labels():
    """Run complete workflow: slice audio then sequence the blocks"""
    print(f"{Fore.CYAN}=== Slice & Sequence with Labels ==={Style.RESET_ALL}")
    
    audio_file = select_audio_file()
    if not audio_file:
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return
    
    blocks_dir = select_output_folder()
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    print(f"{Fore.BLUE}Step 3: Slicing audio with labels...{Style.RESET_ALL}")
    result_dir = slice_audio_from_labels(audio_file, blocks_dir)
    
    if not result_dir:
        print(f"{Fore.RED}❌ Audio slicing failed. Exiting.{Style.RESET_ALL}")
        return
    
    print(f"{Fore.BLUE}Step 4: Sequence configuration...{Style.RESET_ALL}")
    m_blocks, v_blocks, j_blocks = scan_available_blocks(blocks_dir)
    
    if not validate_sequence_requirements(m_blocks, v_blocks, j_blocks):
        print(f"{Fore.RED}❌ Not enough blocks for sequencing{Style.RESET_ALL}")
        return

    max_blocks = min(len(m_blocks), len(v_blocks) + len(j_blocks))
    max_minutes = (max_blocks * 30) / 60

    print(f"{Fore.GREEN}✅ Available: {len(m_blocks)} music + {len(v_blocks)} voice + {len(j_blocks)} jingle blocks{Style.RESET_ALL}")
    print(f"{Fore.GREEN}📊 Maximum sequence: {max_minutes:.1f} minutes{Style.RESET_ALL}")

    while True:
        try:
            user_input = input(f"{Fore.WHITE}How many minutes? (Enter for max {max_minutes:.1f}): {Style.RESET_ALL}").strip()
            if not user_input:
                desired_minutes = None
                break
            desired_minutes = float(user_input)
            if desired_minutes <= 0:
                print(f"{Fore.RED}❌ Please enter a positive number{Style.RESET_ALL}")
                continue
            if desired_minutes > max_minutes:
                print(f"{Fore.RED}❌ Cannot create {desired_minutes:.1f} minutes. Maximum possible is {max_minutes:.1f} minutes{Style.RESET_ALL}")
                continue
            break
        except ValueError:
            print(f"{Fore.RED}❌ Please enter a valid number{Style.RESET_ALL}")

    success, final_audio, blocks_info = create_sequence_from_blocks(blocks_dir, desired_minutes)
    if not success:
        print(f"{Fore.RED}❌ Sequencing failed{Style.RESET_ALL}")
        return

    output_path = ask_save_file()
    
    if not output_path:
        print(f"{Fore.RED}❌ No output file selected. Exiting.{Style.RESET_ALL}")
        return
    
    try:
        print(f"{Fore.BLUE}Exporting final sequence...{Style.RESET_ALL}")
        final_audio.export(output_path, format="mp3", bitrate="192k")
        print(f"{Fore.GREEN}✅ Final sequence saved: {output_path}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}🎵 Final duration: {blocks_info['total_duration']:.1f} seconds{Style.RESET_ALL}")
        
        generate_sequence_timeline(output_path, blocks_dir, 
                                 blocks_info['m_sequence'], blocks_info['voice_sequence'],
                                 blocks_info['total_duration'])
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error exporting sequence: {e}{Style.RESET_ALL}")
        return
    
    print(f"{Fore.CYAN}=== Slice & Sequence Workflow Completed ==={Style.RESET_ALL}")

def generate_sequence_timeline(sequence_path, blocks_dir, m_sequence, voice_sequence, audio_duration):
    """Generate a timeline text file for the created sequence"""
    try:
        txt_path = os.path.splitext(sequence_path)[0] + '.txt'
        
        from datetime import datetime
        created_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        minutes = int(audio_duration // 60)
        seconds = int(audio_duration % 60)
        duration_str = f"{minutes:02d}:{seconds:02d}"
        
        # Count blocks by category
        block_counts = {}
        for block in m_sequence + voice_sequence:
            category = block[0]
            block_counts[category] = block_counts.get(category, 0) + 1
        
        blocks_used = ", ".join([f"{cat}={count}" for cat, count in sorted(block_counts.items())])
        
        # Initialize metadata dictionaries
        descriptions = {}
        origins = {}
        
        excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
        
        # Try to read from Excel first
        try:
            for category in ['m', 'v', 'j']:
                try:
                    df = pd.read_excel(excel_path, sheet_name=category)
                    if not df.empty and category in df.columns:
                        for _, row in df.iterrows():
                            if pd.notna(row[category]):
                                filename = row[category]
                                desc = row.get('description', 'No description')
                                origin = row.get('origin', 'Unknown origin')
                                descriptions[filename] = desc
                                origins[filename] = origin
                except:
                    continue
        except Exception as e:
            print(f"{Fore.YELLOW}⚠️  Could not read Excel for metadata: {e}{Style.RESET_ALL}")
        
        # Fall back to MP3 metadata
        for block in m_sequence + voice_sequence:
            block_name = os.path.splitext(block)[0]
            
            if (block_name not in descriptions or 
                descriptions.get(block_name) == 'No description' or
                block_name not in origins or 
                origins.get(block_name) == 'Unknown origin'):
                
                mp3_path = os.path.join(blocks_dir, block)
                if os.path.exists(mp3_path):
                    metadata = read_audio_metadata(mp3_path)
                    if metadata:
                        if metadata.get('description'):
                            descriptions[block_name] = metadata['description']
                        if metadata.get('origin'):
                            origins[block_name] = metadata['origin']
        
        # Build timeline entries
        timeline_entries = []
        
        for i in range(len(m_sequence)):
            # Music block start time (delayed by 15 seconds)
            music_time = (i * 30) + 15
            music_minutes = music_time // 60
            music_seconds = music_time % 60
            music_time_str = f"{music_minutes:02d}:{music_seconds:02d}"
            
            music_block = m_sequence[i]
            music_name = os.path.splitext(music_block)[0]
            music_desc = descriptions.get(music_name, 'No description')
            music_origin = origins.get(music_name, 'Unknown origin')
            
            timeline_entries.append({
                'time': music_time,
                'time_str': music_time_str,
                'block': music_name,
                'description': music_desc,
                'origin': music_origin,
                'type': 'music'
            })
            
            # Voice/jingle block start time (starts immediately)
            if i < len(voice_sequence):
                voice_time = i * 30
                voice_minutes = voice_time // 60
                voice_seconds = voice_time % 60
                voice_time_str = f"{voice_minutes:02d}:{voice_seconds:02d}"
                
                voice_block = voice_sequence[i]
                voice_name = os.path.splitext(voice_block)[0]
                voice_desc = descriptions.get(voice_name, 'No description')
                voice_origin = origins.get(voice_name, 'Unknown origin')
                voice_type = "jingle" if voice_block.startswith('j') else "voice"
                
                timeline_entries.append({
                    'time': voice_time,
                    'time_str': voice_time_str,
                    'block': voice_name,
                    'description': voice_desc,
                    'origin': voice_origin,
                    'type': voice_type
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
                type_indicator = "[J]" if entry['type'] == 'jingle' else "[V]" if entry['type'] == 'voice' else "[M]"
                f.write(f"{entry['time_str']} {type_indicator} {entry['block']} - {entry['description']} (from: {entry['origin']})\n")
        
        print(f"{Fore.GREEN}✅ Timeline saved: {txt_path}{Style.RESET_ALL}")
        return True
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error generating timeline: {e}{Style.RESET_ALL}")
        return False

def calculate_max_possible_minutes(audio_duration_seconds, slice_duration=30, min_spacing=60):
    """Calculate maximum minutes of content that can be extracted with proper spacing"""
    max_slices = int(audio_duration_seconds / min_spacing)
    max_minutes = (max_slices * slice_duration) / 60
    return max_minutes

def generate_balanced_random_slices(audio_duration_seconds, total_minutes, slice_duration=30, min_spacing=60):
    """Generate random slices with balanced m/v/j split and proper spacing"""
    num_slices = int(total_minutes * 2)
    
    # Ensure divisible by 3 for balanced distribution
    if num_slices % 3 != 0:
        num_slices = ((num_slices // 3) + 1) * 3
    
    num_m = num_v = num_j = num_slices // 3
    
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
    
    # Generate jingle slices (j)
    for i in range(num_j):
        slice_info = _generate_slice_with_spacing(
            safe_start, safe_end, used_positions, min_spacing,
            'j', f"Random jingle segment {i+1}", slice_duration
        )
        if slice_info:
            slices.append(slice_info)
    
    slices.sort(key=lambda x: x['climax_time'])
    return slices

def _generate_slice_with_spacing(safe_start, safe_end, used_positions, min_spacing, slice_type, description, slice_duration):
    """Helper function to generate a single slice with proper spacing"""
    max_attempts = 100
    
    for attempt in range(max_attempts):
        center = random.uniform(safe_start, safe_end)
        
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
    
    print(f"{Fore.YELLOW}⚠️  Could not find valid position for {description} after {max_attempts} attempts{Style.RESET_ALL}")
    return None

def generate_random_slices_and_sequence():
    """Option 3 → Option 2 workflow: Audio file → Generate random slices → Slice → Sequence"""
    print(f"{Fore.CYAN}=== Option 3 → Option 2: Random Slice & Sequence ==={Style.RESET_ALL}")
    print(f"{Fore.BLUE}This will:{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Generate random slices from your audio{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Create balanced music/voice/jingle content{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Automatically sequence the slices{Style.RESET_ALL}")
    print(f"{Fore.BLUE}  • Apply professional audio processing{Style.RESET_ALL}")
    print()
    
    audio_file = select_audio_file()
    if not audio_file:
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return
    
    print(f"{Fore.BLUE}Loading audio file...{Style.RESET_ALL}")
    try:
        audio = AudioSegment.from_file(audio_file)
        audio_duration_seconds = len(audio) / 1000
        audio_duration_minutes = audio_duration_seconds / 60
        
        print(f"{Fore.GREEN}✅ Audio loaded: {audio_duration_minutes:.1f} minutes ({audio_duration_seconds:.0f} seconds){Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}❌ Error loading audio file: {e}{Style.RESET_ALL}")
        return
    
    max_minutes = calculate_max_possible_minutes(audio_duration_seconds)
    print(f"{Fore.BLUE}Maximum content that can be extracted: {max_minutes:.1f} minutes{Style.RESET_ALL}")
    print()
    
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
    
    print(f"{Fore.BLUE}Generating {requested_minutes:.1f} minutes of random slices...{Style.RESET_ALL}")
    slices = generate_balanced_random_slices(audio_duration_seconds, requested_minutes)
    
    if not slices:
        print(f"{Fore.RED}❌ Could not generate valid slices{Style.RESET_ALL}")
        return
    
    num_slices = len(slices)
    num_m = len([s for s in slices if s['type'] == 'm'])
    num_v = len([s for s in slices if s['type'] == 'v'])
    num_j = len([s for s in slices if s['type'] == 'j'])
    
    print(f"{Fore.GREEN}✅ Generated {num_slices} slices ({num_m} music, {num_v} voice, {num_j} jingles){Style.RESET_ALL}")
    
    blocks_dir = select_output_folder()
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")

    import tempfile
    temp_txt_path = os.path.join(tempfile.gettempdir(), "random_slices_temp.txt")
    
    try:
        with open(temp_txt_path, 'w', encoding='utf-8') as f:
            for slice_info in slices:
                f.write(f"{slice_info['climax_time']:.2f}\t{slice_info['climax_time']:.2f}\t{slice_info['type']}\t{slice_info['description']}\n")
        
        print(f"{Fore.GREEN}✅ Created slice definitions{Style.RESET_ALL}")
        
        print(f"{Fore.BLUE}Step 3: Slicing audio...{Style.RESET_ALL}")
                
        print(f"{Fore.GREEN}Audio file: {audio_file}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}Output directory: {blocks_dir}{Style.RESET_ALL}")
        print()
        
        print(f"{Fore.CYAN}Processing slices...{Style.RESET_ALL}")
        for slice_info in slices:

            
            output_path = process_audio_slice_mp3(audio, slice_info, blocks_dir, audio_file)
            print()
        
        verify_files_vs_excel(blocks_dir, excel_path)
        print(f"{Fore.GREEN}✅ Audio slicing completed!{Style.RESET_ALL}")
        
        print(f"{Fore.BLUE}Step 4: Sequencing slices...{Style.RESET_ALL}")
        success, final_audio, blocks_info = create_sequence_from_blocks(blocks_dir, requested_minutes)
        if not success:
            print(f"{Fore.YELLOW}⚠️  Sequencing failed, but slicing completed successfully{Style.RESET_ALL}")
            return

        output_path = ask_save_file()
        
        if not output_path:
            print(f"{Fore.YELLOW}⚠️  No output file selected, but slicing completed successfully{Style.RESET_ALL}")
            return
        
        final_audio.export(output_path, format="mp3", bitrate="192k")
        print(f"{Fore.GREEN}✅ Final sequence saved: {output_path}{Style.RESET_ALL}")
        print(f"{Fore.GREEN}🎵 Final duration: {len(final_audio)/1000:.1f} seconds{Style.RESET_ALL}")
        
        generate_sequence_timeline(output_path, blocks_dir, blocks_info['m_sequence'], blocks_info['voice_sequence'], blocks_info['total_duration'])
        
        print(f"{Fore.CYAN}=== Option 3 → Option 2 Workflow Completed ==={Style.RESET_ALL}")
        actual_minutes = blocks_info['total_duration'] / 60
        print(f"{Fore.GREEN}🎉 Successfully created {actual_minutes:.1f} minutes of sequenced content!{Style.RESET_ALL}")

    finally:
        if os.path.exists(temp_txt_path):
            os.remove(temp_txt_path)

def run_audio_slicer_with_labels():
    """Run audio slicer with existing label file"""
    print(f"{Fore.CYAN}=== Audio Slicer with Labels ==={Style.RESET_ALL}")
    
    audio_file = select_audio_file()
    if not audio_file:
        print(f"{Fore.RED}❌ No audio file selected. Exiting.{Style.RESET_ALL}")
        return
    
    blocks_dir = select_output_folder()
    if not blocks_dir:
        print(f"{Fore.RED}❌ No output folder selected. Exiting.{Style.RESET_ALL}")
        return
    
    result = slice_audio_from_labels(audio_file, blocks_dir)
    if result:
        print(f"{Fore.GREEN}✅ Audio slicing completed successfully!{Style.RESET_ALL}")
    else:
        print(f"{Fore.RED}❌ Audio slicing failed.{Style.RESET_ALL}")

def generate_timestamp_id():
    """Generate a unique timestamp ID in format YYYYMMDDHHMMSSCC"""
    from datetime import datetime
    now = datetime.now()
    return now.strftime("%Y%m%d%H%M%S") + f"{now.microsecond // 10000:02d}"

def create_sequence_from_blocks(blocks_dir, desired_minutes=None):
    """
    Core sequencing function used by both Option 2 and Option 3.2
    If desired_minutes is None, use all available blocks
    Returns: success (bool), final_audio (AudioSegment), selected_blocks_info (dict)
    """
    print(f"{Fore.CYAN}=== Creating Audio Sequence ==={Style.RESET_ALL}")
    
    print(f"{Fore.BLUE}Scanning for audio blocks...{Style.RESET_ALL}")
    m_blocks, v_blocks, j_blocks = scan_available_blocks(blocks_dir)
    
    if not validate_sequence_requirements(m_blocks, v_blocks, j_blocks):
        return False, None, None
    
    # Check for problematic files
    print(f"{Fore.BLUE}Checking audio files...{Style.RESET_ALL}")
    m_blocks_valid, m_problematic = check_for_corrupted_files(blocks_dir, m_blocks)
    v_blocks_valid, v_problematic = check_for_corrupted_files(blocks_dir, v_blocks)
    j_blocks_valid, j_problematic = check_for_corrupted_files(blocks_dir, j_blocks)
    
    # Report problematic files and offer to fix them
    all_problematic = m_problematic + v_problematic + j_problematic
    if all_problematic:
        print(f"{Fore.YELLOW}⚠️  Found {len(all_problematic)} files with loading issues:{Style.RESET_ALL}")
        for filename, error in all_problematic:
            print(f"   - {filename}: {error}")
        
        # Offer to fix problematic files
        if all_problematic:
            response = input(f"{Fore.WHITE}Would you like to attempt to fix these files? (y/N): {Style.RESET_ALL}").strip().lower()
            if response in ['y', 'yes']:
                fixed_count = 0
                for filename, error in all_problematic:
                    file_path = os.path.join(blocks_dir, filename)
                    if fix_problematic_file(file_path):
                        fixed_count += 1
                        # Re-check if the file is now valid
                        try:
                            audio = AudioSegment.from_file(file_path)
                            if len(audio) > 0:
                                # Add to valid lists based on file prefix
                                if filename.startswith('m'):
                                    m_blocks_valid.append(filename)
                                elif filename.startswith('v'):
                                    v_blocks_valid.append(filename)
                                elif filename.startswith('j'):
                                    j_blocks_valid.append(filename)
                        except:
                            pass
                
                print(f"{Fore.GREEN}✅ Fixed {fixed_count} files{Style.RESET_ALL}")
                
                # Remove fixed files from problematic list
                all_problematic = [p for p in all_problematic if not any(p[0] in valid_list for valid_list in [m_blocks_valid, v_blocks_valid, j_blocks_valid])]
    
    # Use only valid files
    m_blocks = m_blocks_valid
    v_blocks = v_blocks_valid
    j_blocks = j_blocks_valid
    
    # Check if we still have enough files
    if len(m_blocks) < 3 or (len(v_blocks) + len(j_blocks)) < 3:
        print(f"{Fore.RED}❌ Not enough valid files after filtering. Need at least 3 music and 3 voice+jingle blocks.{Style.RESET_ALL}")
        print(f"{Fore.RED}   Valid music: {len(m_blocks)}, Valid voice+jingle: {len(v_blocks) + len(j_blocks)}{Style.RESET_ALL}")
        return False, None, None
    
    if desired_minutes is not None:
        blocks_needed = int((desired_minutes * 60) / 30)
        blocks_to_use = min(blocks_needed, len(m_blocks), len(v_blocks) + len(j_blocks))
        print(f"{Fore.BLUE}Using {blocks_to_use} blocks from each channel for {desired_minutes} minute sequence{Style.RESET_ALL}")
    else:
        blocks_to_use = min(len(m_blocks), len(v_blocks) + len(j_blocks))
        print(f"{Fore.BLUE}Using all available blocks: {blocks_to_use} from each channel{Style.RESET_ALL}")
    
    m_sequence, voice_sequence = create_random_sequence(m_blocks, v_blocks, j_blocks)
    
    if desired_minutes is not None:
        m_sequence = m_sequence[:blocks_to_use]
        voice_sequence = voice_sequence[:blocks_to_use]
        print(f"{Fore.GREEN}Selected {blocks_to_use} blocks from each channel{Style.RESET_ALL}")
    
    final_audio = build_multi_channel_sequence(blocks_dir, m_sequence, voice_sequence)
    if not final_audio:
        return False, None, None
    
    selected_blocks_info = {
        'm_sequence': m_sequence,
        'voice_sequence': voice_sequence,
        'blocks_dir': blocks_dir,
        'total_duration': len(final_audio) / 1000
    }
    
    return True, final_audio, selected_blocks_info

def write_audio_metadata(file_path, origin, description, audio_type, climax_time):
    """Write metadata to MP3 file including origin and description"""
    try:
        audiofile = eyed3.load(file_path)
        if audiofile.tag is None:
            audiofile.initTag()
        
        audiofile.tag.artist = f"Audio Slicer - {audio_type}"
        audiofile.tag.album = "Audio Blocks"
        audiofile.tag.title = f"{audio_type} block - {description[:50]}"
        
        audiofile.tag.comments.set(f"Origin: {origin} | Description: {description} | Climax: {climax_time}s | Type: {audio_type}")

        audiofile.tag.user_text_frames.set("ORIGIN_FILE", origin)
        audiofile.tag.user_text_frames.set("DESCRIPTION", description)
        audiofile.tag.user_text_frames.set("AUDIO_TYPE", audio_type)
        audiofile.tag.user_text_frames.set("CLIMAX_TIME", str(climax_time))
        audiofile.tag.user_text_frames.set("SLICE_SIZE", str(SLICE_SIZE))
        
        audiofile.tag.save()
        print(f"{Fore.BLUE}   📝 Metadata written: origin, description, type{Style.RESET_ALL}")
        return True
    except Exception as e:
        print(f"{Fore.YELLOW}⚠️  Could not write metadata to {file_path}: {e}{Style.RESET_ALL}")
        return False

def read_audio_metadata(file_path):
    """Read metadata from MP3 file, including parsing the Origin comment field"""
    try:
        audiofile = eyed3.load(file_path)
        if audiofile.tag is None:
            return None
        
        metadata = {
            'origin': None,
            'description': None, 
            'audio_type': None,
            'climax_time': None
        }
        
        for frame in audiofile.tag.user_text_frames:
            if frame.description == "ORIGIN_FILE":
                metadata['origin'] = frame.text
            elif frame.description == "DESCRIPTION":
                metadata['description'] = frame.text
            elif frame.description == "AUDIO_TYPE":
                metadata['audio_type'] = frame.text
            elif frame.description == "CLIMAX_TIME":
                metadata['climax_time'] = frame.text
        
        if (not metadata['origin'] or not metadata['description']) and audiofile.tag.comments:
            for comment in audiofile.tag.comments:
                comment_text = comment.text
                if "Origin: " in comment_text and "Description: " in comment_text:
                    try:
                        import re
                        if not metadata['origin']:
                            origin_match = re.search(r'Origin: ([^|]+)', comment_text)
                            if origin_match:
                                metadata['origin'] = origin_match.group(1).strip()
                        
                        if not metadata['description']:
                            desc_match = re.search(r'Description: ([^|]+?)(?:\s*\||$)', comment_text)
                            if desc_match:
                                desc = desc_match.group(1).strip()
                                metadata['description'] = desc
                    except Exception as e:
                        print(f"{Fore.YELLOW}⚠️  Error parsing comment: {e}{Style.RESET_ALL}")
        
        return metadata
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error reading metadata from {file_path}: {e}{Style.RESET_ALL}")
        return None

def verify_audio_metadata(blocks_dir):
    """Verify that all audio files have proper metadata"""
    print(f"\n{Fore.CYAN}=== Verifying Audio File Metadata ==={Style.RESET_ALL}")
    
    try:
        all_files = os.listdir(blocks_dir)
        audio_files = [f for f in all_files if f.endswith('.mp3')]
        
        if not audio_files:
            print(f"{Fore.YELLOW}⚠️  No MP3 files found in {blocks_dir}{Style.RESET_ALL}")
            return
        
        metadata_count = 0
        missing_metadata = []
        
        for audio_file in audio_files:
            file_path = os.path.join(blocks_dir, audio_file)
            metadata = read_audio_metadata(file_path)
            
            if metadata and metadata.get('origin') and metadata.get('description'):
                metadata_count += 1
                print(f"{Fore.GREEN}   ✅ {audio_file}: has metadata{Style.RESET_ALL}")
            else:
                missing_metadata.append(audio_file)
                print(f"{Fore.YELLOW}   ⚠️  {audio_file}: missing metadata{Style.RESET_ALL}")
        
        print(f"\n{Fore.CYAN}--- Metadata Summary ---{Style.RESET_ALL}")
        print(f"{Fore.GREEN}✅ Files with complete metadata: {metadata_count}/{len(audio_files)}{Style.RESET_ALL}")
        
        if missing_metadata:
            print(f"{Fore.YELLOW}⚠️  Files missing metadata: {len(missing_metadata)}{Style.RESET_ALL}")
            for file in missing_metadata:
                print(f"   - {file}")
                
        return metadata_count == len(audio_files)
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error verifying metadata: {e}{Style.RESET_ALL}")
        return False

# Add this section to the function after scanning files
def update_excel_from_folder(blocks_dir, excel_path):
    # ... existing code ...
    
    # Remove entries for files that don't exist
    cleanup_count = 0
    for sheet_name in ['m', 'v', 'j']:
        existing_df = existing_data[sheet_name]
        column_name = sheet_name
        
        if not existing_df.empty and column_name in existing_df.columns:
            # Keep only entries where the file actually exists
            valid_entries = []
            for _, row in existing_df.iterrows():
                filename = f"{row[column_name]}.mp3"
                if filename in all_blocks:
                    valid_entries.append(row)
                else:
                    cleanup_count += 1
                    print(f"{Fore.YELLOW}   🗑️  Removing orphaned entry: {filename}{Style.RESET_ALL}")
            
            existing_data[sheet_name] = pd.DataFrame(valid_entries)
    
    if cleanup_count > 0:
        print(f"{Fore.GREEN}✅ Cleaned up {cleanup_count} orphaned Excel entries{Style.RESET_ALL}")

def show_advanced_menu():
    """Show advanced options menu"""
    advanced_text = f"""
{Fore.CYAN}Advanced Options:{Style.RESET_ALL}
{Fore.GREEN}1 - Update Excel from existing blocks folder{Style.RESET_ALL}
{Fore.BLUE}2 - Verify files vs Excel database{Style.RESET_ALL}
{Fore.MAGENTA}3 - Verify audio file metadata{Style.RESET_ALL}
{Fore.YELLOW}4 - Back to main menu{Style.RESET_ALL}
"""
    print(advanced_text)
    
    while True:
        choice = input(f"{Fore.WHITE}Select option (1-4): {Style.RESET_ALL}").strip()
        if choice in ['1', '2', '3', '4']:
            return choice
        else:
            print(f"{Fore.RED}❌ Invalid choice. Please enter 1-4.{Style.RESET_ALL}")

def run_advanced_options():
    """Run advanced options"""
    while True:
        choice = show_advanced_menu()
        
        if choice == '1':
            # Update Excel from folder
            blocks_dir = select_blocks_folder()
            if blocks_dir:
                excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
                update_excel_from_folder(blocks_dir, excel_path)
            else:
                print(f"{Fore.RED}❌ No folder selected{Style.RESET_ALL}")
                
        elif choice == '2':
            # Verify files vs Excel (existing function)
            blocks_dir = select_blocks_folder()
            if blocks_dir:
                excel_path = os.path.join(blocks_dir, "blocks_list.xlsx")
                verify_files_vs_excel(blocks_dir, excel_path)
            else:
                print(f"{Fore.RED}❌ No folder selected{Style.RESET_ALL}")
                
        elif choice == '3':
            # Verify metadata (existing function)
            blocks_dir = select_blocks_folder()
            if blocks_dir:
                verify_audio_metadata(blocks_dir)
            else:
                print(f"{Fore.RED}❌ No folder selected{Style.RESET_ALL}")
                
        elif choice == '4':
            break
        
        input(f"\n{Fore.WHITE}Press Enter to continue...{Style.RESET_ALL}")

def update_excel_from_folder(blocks_dir, excel_path):
    """Scan blocks folder and update Excel with all files found, removing orphaned entries"""
    print(f"{Fore.CYAN}=== Updating Excel from Folder Scan ==={Style.RESET_ALL}")
    
    if not os.path.exists(blocks_dir):
        print(f"{Fore.RED}❌ Blocks directory not found: {blocks_dir}{Style.RESET_ALL}")
        return False
    
    # Scan for audio files
    m_blocks, v_blocks, j_blocks = scan_available_blocks(blocks_dir)
    all_blocks = m_blocks + v_blocks + j_blocks
    
    if not all_blocks:
        print(f"{Fore.YELLOW}⚠️  No audio blocks found in {blocks_dir}{Style.RESET_ALL}")
        return False
    
    print(f"{Fore.GREEN}Found {len(all_blocks)} audio files to process{Style.RESET_ALL}")
    
    # Create or load Excel file
    try:
        # Try to read existing Excel file
        try:
            existing_data = {}
            for sheet_name in ['m', 'v', 'j']:
                try:
                    df = pd.read_excel(excel_path, sheet_name=sheet_name)
                    existing_data[sheet_name] = df
                except:
                    existing_data[sheet_name] = pd.DataFrame(columns=[sheet_name, 'origin', 'description'])
        except FileNotFoundError:
            # Create new Excel structure
            existing_data = {
                'm': pd.DataFrame(columns=['m', 'origin', 'description']),
                'v': pd.DataFrame(columns=['v', 'origin', 'description']),
                'j': pd.DataFrame(columns=['j', 'origin', 'description'])
            }
        
        # CLEANUP PHASE: Remove orphaned entries
        cleanup_count = 0
        for sheet_name in ['m', 'v', 'j']:
            existing_df = existing_data[sheet_name]
            column_name = sheet_name
            
            if not existing_df.empty and column_name in existing_df.columns:
                # Keep only entries where the file actually exists
                valid_entries = []
                for _, row in existing_df.iterrows():
                    filename = f"{row[column_name]}.mp3"
                    if filename in all_blocks:
                        valid_entries.append(row)
                    else:
                        cleanup_count += 1
                        print(f"{Fore.YELLOW}   🗑️  Removing orphaned entry: {filename}{Style.RESET_ALL}")
                
                existing_data[sheet_name] = pd.DataFrame(valid_entries)
        
        if cleanup_count > 0:
            print(f"{Fore.GREEN}✅ Cleaned up {cleanup_count} orphaned Excel entries{Style.RESET_ALL}")
        
        # ADDITION PHASE: Add new files
        new_entries = {'m': [], 'v': [], 'j': []}
        updated_count = 0
        skipped_count = 0
        
        for block_file in all_blocks:
            block_path = os.path.join(blocks_dir, block_file)
            block_name = os.path.splitext(block_file)[0]  # Remove extension
            
            # Determine type from filename prefix
            block_type = block_file[0]  # 'm', 'v', or 'j'
            
            if block_type not in ['m', 'v', 'j']:
                print(f"{Fore.YELLOW}⚠️  Skipping {block_file}: unknown type prefix{Style.RESET_ALL}")
                skipped_count += 1
                continue
            
            # Check if already in Excel (after cleanup)
            existing_df = existing_data[block_type]
            column_name = block_type  # 'm', 'v', or 'j'
            
            if not existing_df.empty and column_name in existing_df.columns:
                if block_name in existing_df[column_name].values:
                    print(f"{Fore.BLUE}   📋 Already in Excel: {block_file}{Style.RESET_ALL}")
                    skipped_count += 1
                    continue
            
            # Read metadata from audio file
            metadata = read_audio_metadata(block_path)
            
            if metadata:
                origin = metadata.get('origin', 'Unknown origin')
                description = metadata.get('description', 'No description')
                
                new_entry = {
                    column_name: block_name,
                    'origin': origin,
                    'description': description
                }
                new_entries[block_type].append(new_entry)
                print(f"{Fore.GREEN}   ✅ Will add: {block_file}{Style.RESET_ALL}")
                updated_count += 1
            else:
                # Create entry with basic info if no metadata
                new_entry = {
                    column_name: block_name,
                    'origin': 'Unknown origin',
                    'description': 'Imported from folder scan'
                }
                new_entries[block_type].append(new_entry)
                print(f"{Fore.YELLOW}   ⚠️  Adding without metadata: {block_file}{Style.RESET_ALL}")
                updated_count += 1
        
        # Update Excel file
        if updated_count > 0 or cleanup_count > 0:
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                for sheet_name in ['m', 'v', 'j']:
                    # Combine existing and new data
                    existing_df = existing_data[sheet_name]
                    new_df = pd.DataFrame(new_entries[sheet_name])
                    
                    if not new_df.empty:
                        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                    else:
                        combined_df = existing_df
                    
                    # Write to Excel
                    combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            if updated_count > 0:
                print(f"{Fore.GREEN}✅ Excel updated: {updated_count} new entries added{Style.RESET_ALL}")
            if cleanup_count > 0:
                print(f"{Fore.GREEN}🗑️  Excel cleaned: {cleanup_count} orphaned entries removed{Style.RESET_ALL}")
        else:
            print(f"{Fore.BLUE}📋 No changes needed - Excel is already synchronized with folder{Style.RESET_ALL}")
        
        return True
        
    except Exception as e:
        print(f"{Fore.RED}❌ Error updating Excel: {e}{Style.RESET_ALL}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Main program entry point"""
    try:
        while True:
            choice = show_welcome_screen()
            
            if choice == '1':
                slice_choice = show_slice_options_menu()
                if slice_choice == '1':
                    run_audio_slicer_with_labels()
                elif slice_choice == '2':
                    no_labels_choice = show_no_labels_menu()
                    if no_labels_choice == '1':
                        print(f"{Fore.GREEN}🎯 Great! Please label your audio file in Audacity...{Style.RESET_ALL}")
                    elif no_labels_choice == '2':
                        run_random_slicer()
                
            elif choice == '2':
                run_sequencer()
                
            elif choice == '3':
                sub_choice = show_slice_and_sequence_menu()
                if sub_choice == '1':
                    run_slice_and_sequence_with_labels()
                elif sub_choice == '2':
                    generate_random_slices_and_sequence()
                
            elif choice == '4':  # Advanced options
                run_advanced_options()
                
            elif choice == '5':  # Help
                show_help()
            
            # Ask if user wants to continue or exit
            print(f"\n{Fore.CYAN}═══════════════════════════════════════════════════{Style.RESET_ALL}")
            continue_choice = input(f"{Fore.WHITE}Return to main menu? (y/N): {Style.RESET_ALL}").strip().lower()
            if continue_choice not in ['y', 'yes']:
                print(f"{Fore.YELLOW}👋 Goodbye!{Style.RESET_ALL}")
                break
                
    except KeyboardInterrupt:
        print(f"\n{Fore.YELLOW}👋 Program interrupted by user. Goodbye!{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}❌ Unexpected error: {e}{Style.RESET_ALL}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()