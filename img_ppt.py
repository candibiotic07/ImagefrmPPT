import os
import zipfile
import re
import xml.etree.ElementTree as ET
import io # Needed for handling image data in memory

# --- Add Pillow Import ---
try:
    from PIL import Image
    PIL_INSTALLED = True
except ImportError:
    PIL_INSTALLED = False
    print("Warning: Pillow library not found (pip install Pillow). Image conversion disabled.")
    print("         Only original image formats will be extracted.")
# ------------------------


def extract_images_from_ppt(ppt_path, output_folder):
    """
    Extracts images from a PPTX file, converts non-GIF images to PNG,
    and saves them with slide numbers in the filename. GIFs are saved as is.

    Requires the Pillow library (`pip install Pillow`) for conversion.

    Args:
        ppt_path (str): The path to the .pptx file.
        output_folder (str): The folder where extracted images will be saved.

    Raises:
        ValueError: If the provided file is not a .pptx file.
        FileNotFoundError: If the ppt_path does not exist.
        ImportError: If Pillow is required but not installed.
    """
    if not PIL_INSTALLED:
         raise ImportError("Pillow library is required for image conversion. Please install it: pip install Pillow")

    # Ensure the file provided is a PPTX file
    if not ppt_path.lower().endswith(".pptx"):
        raise ValueError("Only .pptx files are supported.")

    if not os.path.exists(ppt_path):
         raise FileNotFoundError(f"PPTX file not found at: {ppt_path}")

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # --- Step 1: Build a map from image path to slide number ---
    image_to_slide_map = {}
    slide_num_pattern = re.compile(r'slide(\d+)\.xml\.rels') # Regex to find slide number

    try:
        with zipfile.ZipFile(ppt_path, 'r') as ppt_zip:
            # Find all slide relationship files
            rel_files = [f for f in ppt_zip.namelist() if f.startswith("ppt/slides/_rels/slide") and f.endswith(".xml.rels")]

            for rel_file in rel_files:
                # Extract slide number from the filename
                match = slide_num_pattern.search(os.path.basename(rel_file))
                if not match:
                    continue
                slide_number = int(match.group(1))

                # Read and parse the relationship XML file
                try:
                    rel_xml_content = ppt_zip.read(rel_file)
                    root = ET.fromstring(rel_xml_content)
                    namespaces = {
                        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
                    }

                    for relationship in root.findall('rel:Relationship', namespaces):
                        rel_type = relationship.get('Type')
                        if rel_type and 'image' in rel_type:
                            target = relationship.get('Target')
                            if target:
                                media_filename = os.path.basename(target)
                                full_media_path = f"ppt/media/{media_filename}"
                                if full_media_path not in image_to_slide_map:
                                    image_to_slide_map[full_media_path] = slide_number

                except ET.ParseError:
                    print(f"Warning: Could not parse XML in {rel_file}")
                except Exception as e:
                    print(f"Warning: Error processing {rel_file}: {e}")


            # --- Step 2: Extract images, convert (if needed), and save ---
            extracted_count = 0
            skipped_count = 0
            for file in ppt_zip.namelist():
                # Check if the file is in the ppt/media directory
                if file.startswith("ppt/media/"):
                    original_filename = os.path.basename(file)
                    base_name, original_ext = os.path.splitext(original_filename)
                    original_ext_lower = original_ext.lower()

                    # Determine slide prefix
                    slide_num = image_to_slide_map.get(file)
                    slide_prefix = f"slide{slide_num}_" if slide_num is not None else "slideUNK_"

                    # Only process files with common image extensions or GIF
                    supported_input_formats = ['.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp']
                    if original_ext_lower == '.gif' or original_ext_lower in supported_input_formats:
                        try:
                            # Read image data from zip
                            with ppt_zip.open(file) as source:
                                image_data = source.read()

                            if not image_data: # Skip empty files if they somehow exist
                                print(f"Skipped (Empty File): {original_filename}")
                                skipped_count += 1
                                continue

                            # Handle GIF: Save directly without conversion
                            if original_ext_lower == '.gif':
                                output_filename = f"{slide_prefix}{original_filename}"
                                output_path = os.path.join(output_folder, output_filename)
                                with open(output_path, 'wb') as target:
                                    target.write(image_data)
                                print(f"Extracted (GIF): {output_filename}")
                                extracted_count += 1

                            # Handle other formats: Convert to PNG using Pillow
                            else:
                                output_filename = f"{slide_prefix}{base_name}.png" # Force .png extension
                                output_path = os.path.join(output_folder, output_filename)
                                try:
                                    img = Image.open(io.BytesIO(image_data))

                                    # Handle potential mode issues for PNG saving
                                    # Convert Palette/CMYK to RGBA, other modes to RGB if needed
                                    if img.mode == 'P' or img.mode == 'CMYK':
                                        img = img.convert('RGBA')
                                    elif img.mode not in ['RGB', 'RGBA', 'L']: # L=Grayscale is ok for PNG
                                        img = img.convert('RGB')

                                    img.save(output_path, format='PNG')
                                    print(f"Extracted (Converted to PNG): {output_filename}")
                                    extracted_count += 1
                                except Exception as img_err:
                                    print(f"Warning: Failed to convert '{original_filename}' to PNG. Error: {img_err}. Skipping.")
                                    skipped_count += 1

                        except Exception as e:
                            print(f"Error processing file {file} ('{original_filename}'): {e}. Skipping.")
                            skipped_count += 1
                    else:
                        # Skip files that are not recognized image types or GIF
                        print(f"Skipped (Unsupported format {original_ext}): {original_filename}")
                        skipped_count += 1


            if extracted_count == 0 and skipped_count == 0:
                 print("No media files found in ppt/media/ directory or no image relationships detected.")
            elif extracted_count == 0 and skipped_count > 0:
                 print(f"No images successfully extracted. Skipped {skipped_count} files.")
            else:
                 print(f"\nExtraction complete.")
                 print(f" - Successfully extracted: {extracted_count} images (saved as PNG or GIF)")
                 if skipped_count > 0:
                     print(f" - Skipped: {skipped_count} files (unsupported format, empty, or conversion error)")
                 print(f"Output folder: {output_folder}")


    except zipfile.BadZipFile:
        print(f"Error: The file '{ppt_path}' is not a valid zip file or is corrupted.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


# --- Configuration ---
# Use raw strings (r"...") for Windows paths
ppt_path = r"D:\internships\myedu\pp_timageext\03_Biology\Chapter_6_Life Processes\Lecture_1\Module_1.pptx"  # Replace with your PPTX file path
output_folder = r"D:\internships\myedu\pp_timageext\outputs"     # Replace with your desired output folder

# Call the function
try:
    extract_images_from_ppt(ppt_path, output_folder)
except (ValueError, FileNotFoundError, ImportError) as e:
     print(f"Error: {e}")