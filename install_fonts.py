import os
import shutil
import zipfile
import platform
import ctypes
from pathlib import Path

def extract_fonts_from_zip(zip_path, extract_to):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file in zip_ref.namelist():
            if file.lower().endswith(('.ttf', '.otf')):
                zip_ref.extract(file, extract_to)
                print(f"Extracted {file} from {zip_path}")

def install_font_windows(font_path):
    # Register the font with Windows
    try:
        print(f"Installing {font_path}")
        if ctypes.windll.gdi32.AddFontResourceW(str(font_path)):
            print(f"Successfully registered: {font_path}")
            ctypes.windll.user32.SendMessageW(0xFFFF, 0x1D, 0, 0)  # Refresh system fonts
        else:
            print(f"Failed to register: {font_path}")
    except Exception as e:
        print(f"Error installing {font_path}: {e}")

def install_fonts_from_directory(fonts_dir):
    if not os.path.exists(fonts_dir):
        print(f"Directory '{fonts_dir}' does not exist.")
        return

    temp_extract_dir = os.path.join(fonts_dir, 'extracted_fonts')
    if not os.path.exists(temp_extract_dir):
        os.makedirs(temp_extract_dir)

    zip_files = [f for f in os.listdir(fonts_dir) if f.lower().endswith('.zip')]
    for zip_file in zip_files:
        zip_path = os.path.join(fonts_dir, zip_file)
        extract_fonts_from_zip(zip_path, temp_extract_dir)
    
    font_extensions = ('.ttf', '.otf')
    fonts = [f for f in os.listdir(temp_extract_dir) if f.lower().endswith(font_extensions)]
    
    if not fonts:
        print(f"No font files found in '{temp_extract_dir}'.")
        return

    system = platform.system()
    print(f"Detected OS: {system}")

    if system == 'Windows':
        font_dir = os.path.join(os.environ['WINDIR'], 'Fonts')

        for font in fonts:
            src_font_path = os.path.join(temp_extract_dir, font)
            dest_font_path = os.path.join(font_dir, font)

            # Check if the font is already installed
            if os.path.exists(dest_font_path):
                print(f"Font '{font}' is already installed. Skipping...")
                continue

            try:
                shutil.copy(src_font_path, dest_font_path)
                print(f"Copied {font} to {font_dir}")
                # Register the font with the system
                install_font_windows(dest_font_path)
            except Exception as e:
                print(f"Failed to install {font}: {e}")

    shutil.rmtree(temp_extract_dir)
    print("Font installation completed and temporary files cleaned up.")

if __name__ == '__main__':
    fonts_directory = './fonts'
    install_fonts_from_directory(fonts_directory)
