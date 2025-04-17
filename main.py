import os
import shutil
from img_ppt import extract_images_from_ppt
from descpgen import extract_images_from_ppt as generate_descriptions

def find_pptx_files(root_dir):
    """
    Recursively finds all .pptx files in the given directory and its subdirectories.
    Returns a list of tuples containing (full_path, relative_path)
    """
    pptx_files = []
    for root, _, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith('.pptx'):
                full_path = os.path.join(root, file)
                # Get relative path from the input directory
                relative_path = os.path.relpath(full_path, root_dir)
                pptx_files.append((full_path, relative_path))
    return pptx_files

def setup_output_directories(base_output_path):
    """
    Creates the necessary directory structure for outputs.
    Returns the paths to the images and gifs directories.
    """
    # Create main output directory if it doesn't exist
    os.makedirs(base_output_path, exist_ok=True)
    
    # Create subdirectories for images and gifs
    images_dir = os.path.join(base_output_path, "images")
    gifs_dir = os.path.join(base_output_path, "gifs")
    
    os.makedirs(images_dir, exist_ok=True)
    os.makedirs(gifs_dir, exist_ok=True)
    
    return images_dir, gifs_dir

def organize_files(temp_output_dir, images_dir, gifs_dir, ppt_relative_path):
    """
    Moves files from the temporary output directory to their respective
    final directories based on their extension.
    Also updates the CSV file to include the full PPT path.
    """
    # Move files to appropriate directories
    for filename in os.listdir(temp_output_dir):
        if filename.endswith('.csv'):
            # Read and update the CSV file
            csv_path = os.path.join(temp_output_dir, filename)
            updated_rows = []
            with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                import csv
                reader = csv.reader(csvfile)
                header = next(reader)
                updated_rows.append(header)
                for row in reader:
                    # Update the ppt_name column to include the full path
                    row[2] = ppt_relative_path
                    updated_rows.append(row)
            
            # Write updated CSV to both directories
            for target_dir in [images_dir, gifs_dir]:
                output_csv = os.path.join(target_dir, filename)
                with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerows(updated_rows)
                    
        elif filename.endswith('.gif'):
            shutil.move(
                os.path.join(temp_output_dir, filename),
                os.path.join(gifs_dir, filename)
            )
        else:
            # Move all other image files to images directory
            shutil.move(
                os.path.join(temp_output_dir, filename),
                os.path.join(images_dir, filename)
            )

def process_ppt(ppt_path, ppt_relative_path, base_output_path):
    """
    Processes a single PPT file:
    1. Sets up output directories
    2. Extracts images and gifs
    3. Generates descriptions
    4. Organizes the final output
    """
    try:
        # Setup directories
        images_dir, gifs_dir = setup_output_directories(base_output_path)
        
        # Create a temporary directory for initial processing
        temp_dir = os.path.join(base_output_path, "temp")
        os.makedirs(temp_dir, exist_ok=True)
        
        try:
            # Step 1: Extract images and gifs from PPT
            print(f"\nProcessing PPT: {ppt_relative_path}")
            print("Step 1: Extracting images and gifs from PPT...")
            extract_images_from_ppt(ppt_path, temp_dir)
            
            # Step 2: Generate descriptions
            print("Step 2: Generating descriptions...")
            generate_descriptions(ppt_path, temp_dir)
            
            # Step 3: Organize files into final structure
            print("Step 3: Organizing files...")
            organize_files(temp_dir, images_dir, gifs_dir, ppt_relative_path)
            
        finally:
            # Clean up temporary directory
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
                
    except Exception as e:
        print(f"Error processing {ppt_relative_path}: {str(e)}")
        raise

def process_directory(input_dir, base_output_path):
    """
    Main function that processes all PPT files in the input directory and its subdirectories.
    """
    try:
        # Find all PPTX files
        pptx_files = find_pptx_files(input_dir)
        
        if not pptx_files:
            print(f"No PPTX files found in {input_dir}")
            return
            
        print(f"Found {len(pptx_files)} PPTX files to process")
        
        # Process each PPT file
        for ppt_path, ppt_relative_path in pptx_files:
            process_ppt(ppt_path, ppt_relative_path, base_output_path)
            
        print("\nAll processing complete!")
        print(f"Output directory: {base_output_path}")
        
    except Exception as e:
        print(f"Error during processing: {str(e)}")
        raise

# --- Configuration ---
# Use raw strings (r"...") for Windows paths
INPUT_DIR = r"D:\internships\myedu\pp_timageext"  # Directory containing PPT files
BASE_OUTPUT_PATH = r"D:\internships\myedu\pp_timageext\outputs"

if __name__ == "__main__":
    process_directory(INPUT_DIR, BASE_OUTPUT_PATH) 