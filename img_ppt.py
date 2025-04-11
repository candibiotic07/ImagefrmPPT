import os
import zipfile

def extract_images_from_ppt(ppt_path, output_folder):
    # Ensure the file provided is a PPTX file
    if not ppt_path.endswith(".pptx"):
        raise ValueError("Only .pptx files are supported.")

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Open the PPTX file as a ZIP archive
    with zipfile.ZipFile(ppt_path, 'r') as ppt_zip:
        # Iterate over each file in the archive
        for file in ppt_zip.namelist():
            # Check if the file is in the ppt/media directory,
            # which is where images in a pptx are stored
            if file.startswith("ppt/media/"):
                # Construct the output path for the image
                filename = os.path.basename(file)
                output_path = os.path.join(output_folder, filename)
                
                # Extract the image data and write it to the output file
                with ppt_zip.open(file) as source, open(output_path, 'wb') as target:
                    target.write(source.read())
                
                print(f"Extracted: {filename}")

# Set these variables directly in the code
ppt_path = r"D:\internships\myedu\pp_timageext\03_Biology\Chapter_6_Life Processes\Lecture_1\Module_1.pptx"  # Replace with your PPTX file path
output_folder = r"D:\internships\myedu\pp_timageext\outputs"     # Replace with your desired output folder

# Call the function using the hard-coded values
extract_images_from_ppt(ppt_path, output_folder)
