import os
import shutil
import logging
import csv
import boto3
import urllib.parse  # Keep this
from botocore.exceptions import ClientError
from img_ppt import extract_images_from_ppt
from descpgen import extract_images_from_ppt as generate_descriptions
from dotenv import load_dotenv
from pymongo import MongoClient

# ——— Setup logging ———
logging.basicConfig(
    format="%(asctime)s %(levelname)s %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)
S3_BUCKET = "server-ai-bucket" # Ensure this is correct
def sanitize_path(path: str) -> str:
    """Replace spaces with underscores in directory and file names."""
    # Split the path into components
    parts = path.split(os.sep)
    # Replace spaces with underscores in each component
    sanitized_parts = [part.replace(" ", "_") for part in parts]
    # Join the parts back together
    return os.sep.join(sanitized_parts)

def sanitize_directory_structure(root_dir: str) -> None:
    """Recursively sanitize all directory and file names in the given directory."""
    logger.info(f"Starting directory sanitization for: {root_dir}")
    
    # Walk through the directory tree from bottom to top
    for root, dirs, files in os.walk(root_dir, topdown=False):
        # Process files first
        for name in files:
            old_path = os.path.join(root, name)
            new_name = name.replace(" ", "_")
            if new_name != name:
                new_path = os.path.join(root, new_name)
                try:
                    os.rename(old_path, new_path)
                    logger.info(f"Renamed file: {old_path} -> {new_path}")
                except Exception as e:
                    logger.error(f"Error renaming file {old_path}: {e}")
        
        # Then process directories
        for name in dirs:
            old_path = os.path.join(root, name)
            new_name = name.replace(" ", "_")
            if new_name != name:
                new_path = os.path.join(root, new_name)
                try:
                    os.rename(old_path, new_path)
                    logger.info(f"Renamed directory: {old_path} -> {new_path}")
                except Exception as e:
                    logger.error(f"Error renaming directory {old_path}: {e}")

# ——— Load env & init S3 ———
load_dotenv()
s3 = boto3.client(
    's3',
    aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'),
    region_name='ap-south-1'  # Set the region explicitly
)

# ——— Setup MongoDB ———
# Ensure MongoDB is running or adjust the connection string
try:
    client = MongoClient("mongodb://localhost:27017/", serverSelectionTimeoutMS=5000)
    # The ismaster command is cheap and does not require auth.
    client.admin.command('ismaster')
    db = client["PPT"]
    images_col = db["Images"]
    gifs_col = db["Gifs"]
    logger.info("MongoDB connection successful.")
except Exception as e:
    logger.error(f"MongoDB connection failed: {e}. Please ensure MongoDB is running.")
    # Depending on requirements, you might want to exit here
    # exit(1)
    # Or provide default dummy collections if feasible for testing
    db = None
    images_col = None
    gifs_col = None


# --- *** MODIFY THIS FUNCTION *** ---
def get_s3_url(bucket_name: str, region: str, s3_key: str) -> str:
    """Generate a properly formatted S3 URL.
       Ensures forward slashes and encodes components as needed (e.g., %20 for space).
    """
    # The s3_key passed here *should* now have only forward slashes '/'.
    # Use urllib.parse.quote, keeping '/' safe.
    # Correctly handles spaces (as %20) etc. in path components.
    encoded_key = urllib.parse.quote(s3_key, safe='~/') # << THIS LINE CHANGED
    url = f"https://{bucket_name}.s3.{region}.amazonaws.com/{encoded_key}"
    return url
# --- *** MODIFY THIS FUNCTION *** ---
def upload_file_to_s3(file_path: str, bucket: str, key: str) -> str | None:
    """Uploads a file and returns its S3 URL."""
    # Key received from make_s3_key should already have ONLY forward slashes.
    s3_compatible_key = key # << SIMPLIFIED - No replace needed here

    try:
        s3.upload_file(Filename=file_path, Bucket=bucket, Key=s3_compatible_key)
        # Generate the URL using the already-correct forward-slashed key
        url = get_s3_url(bucket, 'ap-south-1', s3_compatible_key) # << Pass the correct key
        logger.info(f"Uploaded '{os.path.basename(file_path)}' to S3 key: {s3_compatible_key}")
        logger.info(f"Generated S3 URL: {url}") # Log the correctly generated URL
        return url
    except ClientError as e:
        logger.error(f"S3 upload failed for {file_path} (key: {s3_compatible_key}): {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error uploading {file_path} (key: {s3_compatible_key}): {e}")
        return None
# --- *** MODIFY THIS FUNCTION *** ---
def make_s3_key(ppt_relative_path: str, filename: str, media_type: str) -> str:
    """Create an S3 key string using ONLY forward slashes."""
    # 1. Normalize ppt_relative_path to use forward slashes FIRST << NEW LINE
    normalized_ppt_rel_path = ppt_relative_path.replace("\\", "/") # << NEW LINE

    # 2. Extract components using the normalized path (os.path works with '/')
    ppt_dir = os.path.dirname(normalized_ppt_rel_path) # << USES NORMALIZED PATH
    ppt_name_ext = os.path.basename(normalized_ppt_rel_path) # << USES NORMALIZED PATH
    ppt_name = os.path.splitext(ppt_name_ext)[0]

    # 3. Combine using only forward slashes
    clean_filename = os.path.basename(filename) # Good practice

    if ppt_dir:
        # Construct the key string consistently with forward slashes
        key = f"PPT/{ppt_dir}/{ppt_name}/{media_type}/{clean_filename}" # << SAME f-string logic
    else:
        key = f"PPT/{ppt_name}/{media_type}/{clean_filename}" # << SAME f-string logic

    # Replace potential double slashes just in case ppt_dir was empty
    key = key.replace('//', '/') # << MINOR cleanup added

    # Returns a key like: "PPT/Chapter_6_Life_Processes/Lecture_1/Module_1/images/slide5_image25.png"
    return key
def find_pptx_files(root_dir: str) -> list[tuple[str, str]]:
    out = []
    for root, _, files in os.walk(root_dir):
        for f in files:
            if f.lower().endswith('.pptx'):
                full = os.path.join(root, f)
                # Use relpath to get the path relative to the input directory
                # This will naturally use os.sep ('\' on Windows)
                rel = os.path.relpath(full, root_dir)
                out.append((full, rel))
    return out

def setup_output_directories(base_output_path: str, ppt_relative_path: str) -> tuple[str, str]:
    """Create output directories that match the input directory structure"""
    # Preserve the full directory structure from the relative path
    ppt_dir_component = os.path.dirname(ppt_relative_path) # e.g., "Folder Name\SubFolder"
    ppt_name = os.path.splitext(os.path.basename(ppt_relative_path))[0]

    # Create the base output path mirroring the input structure
    # os.path.join handles joining paths correctly regardless of OS
    output_base_for_ppt = os.path.join(base_output_path, ppt_dir_component, ppt_name)

    images = os.path.join(output_base_for_ppt, "images")
    gifs = os.path.join(output_base_for_ppt, "gifs")

    os.makedirs(images, exist_ok=True)
    os.makedirs(gifs, exist_ok=True)
    logger.debug(f"Created output directories: {images} and {gifs}")
    return images, gifs

def update_csv_with_s3_url(csv_path: str, filename: str, s3_url: str, input_path: str) -> bool:
    """Update a specific CSV file with S3 URL information for a given filename"""
    if not os.path.exists(csv_path):
        logger.warning(f"CSV file not found: {csv_path}")
        return False

    rows = []
    fieldnames = []
    updated = False

    try:
        with open(csv_path, 'r', newline='', encoding='utf-8') as infile:
            reader = csv.DictReader(infile)
            # Handle potential empty CSV or missing headers
            fieldnames = list(reader.fieldnames or [])
            if not fieldnames:
                 logger.warning(f"CSV file is empty or has no headers: {csv_path}")
                 # Decide how to handle: skip, create headers, etc.
                 # For now, let's assume we need headers and add them if missing
                 # Or simply return if processing requires existing headers.
                 # Let's try adding them if the file was completely empty.
                 if os.path.getsize(csv_path) == 0:
                      fieldnames = ['image_filename', 'description', 's3_url', 'input_path'] # Example headers
                      logger.info(f"Added default headers to empty CSV: {csv_path}")
                 else: # File has content but no headers? More problematic.
                      logger.error(f"CSV file has content but no headers detected: {csv_path}")
                      return False # Cannot proceed without known headers

            # Ensure required columns exist
            if 's3_url' not in fieldnames:
                fieldnames.append('s3_url')
            if 'input_path' not in fieldnames:
                fieldnames.append('input_path')

            # Determine the filename column (handle variations)
            filename_col = None
            if 'image_filename' in fieldnames:
                 filename_col = 'image_filename'
            elif 'filename' in fieldnames:
                 filename_col = 'filename'
            elif 'gif_filename' in fieldnames: # Add check for gif filename column if necessary
                 filename_col = 'gif_filename'

            if not filename_col:
                logger.error(f"Could not find a suitable filename column in {csv_path}. Headers: {fieldnames}")
                return False # Cannot proceed without knowing which column holds the filename

            # Process rows
            for row in reader:
                # Ensure row is not empty and filename column exists
                if row and row.get(filename_col) == filename:
                    row['s3_url'] = s3_url
                    row['input_path'] = input_path # Store the logical path
                    updated = True
                    logger.info(f"Updated {filename} with S3 URL in {os.path.basename(csv_path)}")
                rows.append(row) # Append row regardless of update status

        if not updated:
            # This might be expected if the CSV is generated by a different process first
            logger.warning(f"No matching filename '{filename}' found in {csv_path} to update.")

    except Exception as e:
        logger.error(f"Error reading CSV {csv_path}: {e}")
        return False

    # Write updated content back
    try:
        with open(csv_path, 'w', newline='', encoding='utf-8') as outfile:
            writer = csv.DictWriter(outfile, fieldnames=fieldnames, quoting=csv.QUOTE_MINIMAL) # Use minimal quoting
            writer.writeheader()
            writer.writerows(rows)
        # Return True if the file was written, even if no specific row was updated
        # (e.g., headers might have been added to an empty file)
        return True
    except Exception as e:
        logger.error(f"Error writing to CSV {csv_path}: {e}")
        return False

def import_csv_dir_to_collection(dir_path: str, collection):
    # Check if MongoDB connection is available
    if collection is None:
        logger.error(f"MongoDB collection not available. Skipping import for {dir_path}")
        return

    total_inserted = 0
    if not os.path.isdir(dir_path):
        logger.warning(f"Directory not found for import: {dir_path}")
        return

    for fname in os.listdir(dir_path):
        if not fname.lower().endswith(".csv"):
            continue

        csv_path = os.path.join(dir_path, fname)
        docs = []

        try:
            with open(csv_path, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                docs = list(reader) # Read all docs into memory

            if not docs:
                logger.info(f"  • No data rows found in {fname}")
                continue

            # Verify S3 URLs are present before import (optional, for logging)
            urls_present = sum(1 for doc in docs if doc.get('s3_url'))
            logger.info(f"  • Found {urls_present}/{len(docs)} docs with S3 URLs in {fname}")

            # Add the relative directory path (if needed)
            # The input_path should already contain the necessary path info from the S3 key
            # base_output_path = os.path.dirname(os.path.dirname(os.path.dirname(dir_path))) # Find base output path
            # rel_dir = os.path.relpath(dir_path, base_output_path)
            # for doc in docs:
            #     if 'directory_path' not in doc:
            #         doc['directory_path'] = rel_dir.replace("\\", "/") # Standardize to forward slashes

            # Perform the insert operation
            result = collection.insert_many(docs)
            count = len(result.inserted_ids)
            total_inserted += count
            logger.info(f"  • Inserted {count} docs from {fname} into {collection.full_name}")

        except FileNotFoundError:
             logger.error(f"CSV file not found during import: {csv_path}")
        except Exception as e:
            logger.error(f"Error importing {csv_path} to {collection.full_name}: {e}")

    logger.info(f"Finished loading {total_inserted} docs into '{collection.full_name}' from directory '{os.path.basename(dir_path)}'.")


def find_filename_column(headers: list[str]) -> str | None:
    """Tries to find the most likely filename column."""
    possible_cols = ['image_filename', 'gif_filename', 'filename']
    for col in possible_cols:
        if col in headers:
            return col
    logger.warning(f"Could not determine filename column from headers: {headers}. Falling back to 'filename'.")
    # Fallback or return None if strict checking is needed
    return 'filename' # Or return None if you want to fail explicitly


def organize_files(temp_output_dir: str, images_dir: str, gifs_dir: str,
                   ppt_relative_path: str, input_dir: str): # Pass input_dir if needed

    # Step 1: Identify and categorize files from temp dir
    image_files = []
    gif_files = []
    temp_csv_files = [] # CSVs generated by descpgen in the temp dir

    if not os.path.exists(temp_output_dir):
        logger.error(f"Temporary output directory does not exist: {temp_output_dir}")
        return

    for fname in os.listdir(temp_output_dir):
        src = os.path.join(temp_output_dir, fname)
        if not os.path.isfile(src): # Skip directories if any
            continue

        low_fname = fname.lower()
        if low_fname.endswith('.gif'):
            gif_files.append(fname)
        elif low_fname.endswith('.csv'):
            temp_csv_files.append(fname) # Keep track of original CSVs
        elif low_fname.endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
            image_files.append(fname)
        else:
            logger.warning(f"Uncategorized file in temp dir: {fname}. Skipping.")

    logger.info(f"Found {len(image_files)} images, {len(gif_files)} GIFs, and {len(temp_csv_files)} CSVs in temp dir.")

    if not temp_csv_files:
        logger.warning(f"No CSV files found in temp directory {temp_output_dir}. Cannot update S3 URLs or import.")
        # Move files without URL info? Or stop? Depends on requirements.
        # For now, let's log a warning and continue moving files, but DB import won't work.
        # Basic move without S3/CSV update:
        # for fname in image_files:
        #     shutil.move(os.path.join(temp_output_dir, fname), os.path.join(images_dir, fname))
        # for fname in gif_files:
        #     shutil.move(os.path.join(temp_output_dir, fname), os.path.join(gifs_dir, fname))
        # return # Exit organize_files if no CSV to work with

    # Assuming one primary CSV generated by descpgen for simplicity. Handle multiple if necessary.
    if len(temp_csv_files) > 1:
        logger.warning(f"Multiple CSVs found in temp dir: {temp_csv_files}. Processing only the first one: {temp_csv_files[0]}")
    elif not temp_csv_files:
         logger.error(f"No CSV file found in {temp_output_dir} to update.")
         # Handle moving files without CSV processing if needed, then return
         return

    original_csv_path = os.path.join(temp_output_dir, temp_csv_files[0])

    # Step 2: Process image files: move, upload, update ORIGINAL CSV
    logger.info("Processing image files...")
    for fname in image_files:
        src = os.path.join(temp_output_dir, fname)
        dst = os.path.join(images_dir, fname)
        try:
            shutil.move(src, dst) # Move image to final image dir
            key = make_s3_key(ppt_relative_path, fname, 'images')
            url = upload_file_to_s3(dst, S3_BUCKET, key)
            input_path = key.replace("PPT/", "", 1) # Consistent relative path for DB

            if url:
                # Update the ORIGINAL CSV in the temp directory
                updated = update_csv_with_s3_url(original_csv_path, fname, url, input_path)
                if not updated:
                    logger.warning(f"Failed to find or update entry for image '{fname}' in {os.path.basename(original_csv_path)}")
            else:
                logger.error(f"Skipping CSV update for image {fname} due to S3 upload failure.")
        except Exception as e:
            logger.error(f"Error processing image file {fname}: {e}")

    # Step 3: Process GIF files: move, upload, update ORIGINAL CSV
    logger.info("Processing GIF files...")
    for fname in gif_files:
        src = os.path.join(temp_output_dir, fname)
        dst = os.path.join(gifs_dir, fname)
        try:
            shutil.move(src, dst) # Move gif to final gif dir
            key = make_s3_key(ppt_relative_path, fname, 'gifs')
            url = upload_file_to_s3(dst, S3_BUCKET, key)
            input_path = key.replace("PPT/", "", 1) # Consistent relative path for DB

            if url:
                 # Update the ORIGINAL CSV in the temp directory
                updated = update_csv_with_s3_url(original_csv_path, fname, url, input_path)
                if not updated:
                    logger.warning(f"Failed to find or update entry for GIF '{fname}' in {os.path.basename(original_csv_path)}")
            else:
                logger.error(f"Skipping CSV update for GIF {fname} due to S3 upload failure.")
        except Exception as e:
            logger.error(f"Error processing GIF file {fname}: {e}")

    # Step 4: Split the updated original CSV into separate image and GIF CSVs
    logger.info(f"Splitting updated CSV '{os.path.basename(original_csv_path)}' into final image/GIF CSVs.")
    image_rows = []
    gif_rows = []
    fieldnames = []
    filename_col = None

    try:
        with open(original_csv_path, 'r', newline='', encoding='utf-8') as infile:
            reader = csv.DictReader(infile)
            fieldnames = list(reader.fieldnames or [])
            if not fieldnames:
                logger.error(f"Original CSV {original_csv_path} is empty or has no headers after updates. Cannot split.")
                return # Cannot proceed

            # Ensure essential columns are present for splitting and import
            if 's3_url' not in fieldnames: fieldnames.append('s3_url')
            if 'input_path' not in fieldnames: fieldnames.append('input_path')

            filename_col = find_filename_column(fieldnames)
            if not filename_col:
                 logger.error(f"Cannot split CSV: Failed to identify filename column in {original_csv_path}. Headers: {fieldnames}")
                 return # Cannot proceed without knowing the filename column

            for row in reader:
                # Check if row is potentially empty or malformed
                if not row or not row.get(filename_col):
                     logger.warning(f"Skipping empty or invalid row in {original_csv_path}: {row}")
                     continue

                # Determine type based on filename extension in the identified column
                filename_in_csv = row.get(filename_col, "")
                if filename_in_csv.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
                    image_rows.append(row)
                elif filename_in_csv.lower().endswith('.gif'):
                    gif_rows.append(row)
                else:
                    logger.warning(f"Row in {original_csv_path} has unhandled file type '{filename_in_csv}'. Skipping.")

    except Exception as e:
        logger.error(f"Error reading original CSV {original_csv_path} for splitting: {e}")
        return # Stop if reading fails

    # Define final CSV paths
    # Use the base name of the original CSV to name the final ones
    base_csv_name = os.path.splitext(os.path.basename(original_csv_path))[0]
    final_image_csv_path = os.path.join(images_dir, f"{base_csv_name}_images.csv")
    final_gif_csv_path = os.path.join(gifs_dir, f"{base_csv_name}_gifs.csv")

    # Write the final image CSV
    try:
        logger.info(f"Writing {len(image_rows)} image rows to {os.path.basename(final_image_csv_path)}")
        with open(final_image_csv_path, 'w', newline='', encoding='utf-8') as outfile:
            # Use fieldnames identified earlier, ensuring consistency
            writer = csv.DictWriter(outfile, fieldnames=fieldnames, quoting=csv.QUOTE_MINIMAL)
            writer.writeheader()
            writer.writerows(image_rows)
    except Exception as e:
        logger.error(f"Error writing final image CSV {final_image_csv_path}: {e}")

    # Write the final GIF CSV
    try:
        logger.info(f"Writing {len(gif_rows)} GIF rows to {os.path.basename(final_gif_csv_path)}")
        with open(final_gif_csv_path, 'w', newline='', encoding='utf-8') as outfile:
             # Use fieldnames identified earlier
            writer = csv.DictWriter(outfile, fieldnames=fieldnames, quoting=csv.QUOTE_MINIMAL)
            writer.writeheader()
            writer.writerows(gif_rows)
    except Exception as e:
        logger.error(f"Error writing final GIF CSV {final_gif_csv_path}: {e}")


    # Step 5: Import the *final, separated* CSV data to MongoDB
    global images_col, gifs_col # Make sure these are accessible
    if images_col is not None:
        logger.info(f"Loading final image data from '{images_dir}' into {images_col.full_name}...")
        import_csv_dir_to_collection(images_dir, images_col)
    else:
        logger.warning("Skipping MongoDB import for images: collection not available.")

    if gifs_col is not None:
        logger.info(f"Loading final GIF data from '{gifs_dir}' into {gifs_col.full_name}...")
        import_csv_dir_to_collection(gifs_dir, gifs_col)
    else:
        logger.warning("Skipping MongoDB import for gifs: collection not available.")

    logger.info(f"Finished organizing files for {ppt_relative_path}")

def process_ppt(ppt_path: str, ppt_relative_path: str, base_output_path: str, input_dir: str):
    ppt_name = os.path.splitext(os.path.basename(ppt_relative_path))[0]
    logger.info(f"Starting processing for: {ppt_relative_path}")

    # Create checkpoint file path based on the *output* structure
    ppt_output_dir_component = os.path.dirname(ppt_relative_path)
    checkpoint_dir = os.path.join(base_output_path, ppt_output_dir_component)
    # Ensure the directory for the checkpoint exists
    os.makedirs(checkpoint_dir, exist_ok=True)
    checkpoint_file = os.path.join(checkpoint_dir, f".{ppt_name}.done")

    if os.path.exists(checkpoint_file):
        logger.info(f"Skipping '{ppt_relative_path}': Checkpoint file found at '{checkpoint_file}'.")
        return

    # Use a unique temp directory for each PPT to avoid conflicts if run in parallel later
    # Place it inside the specific PPT's output base to keep things organized
    ppt_output_base = os.path.join(base_output_path, os.path.dirname(ppt_relative_path), ppt_name)
    temp_dir = os.path.join(ppt_output_base, "temp_processing")
    # Ensure clean state for temp dir
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir, exist_ok=True)
    logger.debug(f"Using temp directory: {temp_dir}")

    images_dir, gifs_dir = setup_output_directories(base_output_path, ppt_relative_path)

    try:
        logger.info(f"Extracting media from: {ppt_relative_path}")
        extract_images_from_ppt(ppt_path, temp_dir) # Extract to temp dir
        logger.info(f"Generating descriptions for: {ppt_relative_path}")
        generate_descriptions(ppt_path, temp_dir) # Generate descriptions in temp dir
        logger.info(f"Organizing files, uploading, and updating DB for: {ppt_relative_path}")
        organize_files(temp_dir, images_dir, gifs_dir, ppt_relative_path, input_dir) # Organize from temp to final dirs
        logger.info(f"Successfully processed: {ppt_relative_path}")

        # Create checkpoint file only on full success
        with open(checkpoint_file, 'w') as f:
            f.write("done") # Write something small to the file
        logger.info(f"Created checkpoint file: {checkpoint_file}")

    except Exception as e:
        logger.exception(f"Error processing {ppt_relative_path}: {e}")
        # Optional: re-raise if you want the script to stop on error
        # raise
    finally:
        # Clean up the unique temp directory
        if os.path.isdir(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                logger.info(f"Cleaned up temp dir: {temp_dir}")
            except Exception as e:
                logger.error(f"Failed to cleanup temp dir {temp_dir}: {e}")


def process_directory(input_dir: str, base_output_path: str):
    if not os.path.isdir(input_dir):
        logger.error(f"Input directory not found: {input_dir}")
        return
    if not os.path.isdir(base_output_path):
        logger.info(f"Base output directory not found: {base_output_path}. Creating it.")
        os.makedirs(base_output_path, exist_ok=True)

    # Sanitize directory structure before processing
    logger.info("Sanitizing directory structure...")
    sanitize_directory_structure(input_dir)
    logger.info("Directory sanitization complete.")

    pptx_files = find_pptx_files(input_dir)
    if not pptx_files:
        logger.warning(f"No PPTX files found in '{input_dir}' or its subdirectories.")
        return

    logger.info(f"Found {len(pptx_files)} PPTX files to process.")
    for ppt_path, ppt_rel in pptx_files:
        process_ppt(ppt_path, ppt_rel, base_output_path, input_dir)

    logger.info("All PPTX file processing finished.")

if __name__ == "__main__":
    # Use raw strings for Windows paths
    INPUT_DIR = r"D:\internships\myedu\pp_timageext\testinputDB"
    BASE_OUTPUT_PATH = r"D:\internships\myedu\pp_timageext\output3"

    # Basic check for input directory existence
    if not os.path.isdir(INPUT_DIR):
        print(f"ERROR: Input directory does not exist: {INPUT_DIR}")
    else:
        process_directory(INPUT_DIR, BASE_OUTPUT_PATH)