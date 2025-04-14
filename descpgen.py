import os
import zipfile
import re
import xml.etree.ElementTree as ET
import io  # For handling image data in memory
import csv  # For CSV output
import time
import base64  # For encoding image data for the new OpenAI API

# --- Add Pillow Import ---
try:
    from PIL import Image
    PIL_INSTALLED = True
except ImportError:
    PIL_INSTALLED = False
    print("Warning: Pillow library not found (pip install Pillow). Image conversion disabled.")
    print("         Only original image formats will be extracted.")
# ------------------------

# --- Load Environment Variables ---
from dotenv import load_dotenv
load_dotenv()  # Loads .env file from the current working directory

# For debugging: verify that the keys are loaded (remove before production!)
print("Google API Key:", os.environ.get("GOOGLE_API_KEY"))
print("OpenAI API Key:", os.environ.get("OPENAI_API_KEY"))

# Ensure that required API keys exist
my_api_key = os.environ.get("GOOGLE_API_KEY")
openai_api_key = os.environ.get("OPENAI_API_KEY")
if not my_api_key:
    raise ValueError("GOOGLE_API_KEY not found in environment variables.")
if not openai_api_key:
    raise ValueError("OPENAI_API_KEY not found in environment variables.")

# --- Gemini API Setup ---
from google import genai
from google.genai import types
client = genai.Client(api_key=my_api_key)

# For GPT-4o, we now use OpenAI's new API, so the legacy GPT4O_MODEL constant is no longer needed.
# Instead, we will use the new model "gpt-4o-mini" in our function below.

# Generation configurations for Gemini (used for non-GIF images)
generation_config = {
    "temperature": 0.2,
    "top_p": 1,
    "top_k": 32,
    "max_output_tokens": 4096,
}
short_output_config = {
    "temperature": 0.1,
    "top_p": 1,
    "top_k": 1,
    "max_output_tokens": 128,
}

def safe_llm_call(model, prompt_parts, is_vision_model=True, config=generation_config):
    """
    Wrapper for LLM calls using the Google Generative AI API.
    Handles both text and image inputs in the prompt_parts list.
    """
    max_retries = 2
    delay = 2  # seconds between retries

    contents = []
    text_parts = ""
    # Separate image objects from text parts
    for part in prompt_parts:
        if isinstance(part, Image.Image):
            if text_parts:
                contents.append(text_parts)
                text_parts = ""
            contents.append(part)
        else:
            text_parts += str(part)
    if text_parts:
        contents.append(text_parts)

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=model,
                contents=contents,
                config=types.GenerateContentConfig(temperature=config.get("temperature", 0))
            )
            if not response.text or not response.text.strip():
                error_msg = "Error: Received empty/invalid response."
                print(f"LLM Call Warning: {error_msg}")
                if attempt < max_retries - 1:
                    print(f"Retrying in {delay} seconds...")
                    time.sleep(delay)
                    continue
                else:
                    return None, error_msg
            return response.text.strip(), None
        except Exception as e:
            error_msg = f"Error during LLM API call: {e}"
            print(f"LLM Call Error: {error_msg}")
            if attempt < max_retries - 1:
                print(f"Retrying in {delay} seconds...")
                time.sleep(delay)
            else:
                return None, error_msg
    return None, "Error: Max retries reached for LLM call."

def generate_description_gpt4o(image_data, hint):
    """
    Generates a description for a GIF image using the new OpenAI API.
    The image_data is provided as raw bytes.
    This function:
      - Encodes the image to base64,
      - Prepares a text prompt (incorporating the hint),
      - Sends both as inputs via the OpenAI client,
      - Returns the resulting description.
    """
    try:
        # Prepare the prompt text: you can modify this as needed.
        prompt_text = f"What is in this image? Hint: {hint}"
        
        # Convert the raw image data to a base64 string
        b64_image = base64.b64encode(image_data).decode("utf-8")
        image_url = f"data:image/png;base64,{b64_image}"  # assuming PNG format; adjust if needed

        # Initialize the OpenAI client using the provided API key.
        from openai import OpenAI
        openai_client = OpenAI(api_key=openai_api_key)

        # Call the new OpenAI API endpoint.
        print("Sending request to OpenAI API for image description...")
        try:
            response = openai_client.chat.completions.create(
                model="gpt-4-turbo",  # Updated to use the current model name
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt_text},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": image_url
                                }
                            }
                        ]
                    }
                ],
                max_tokens=300
            )
            
            # Extract the description from the response
            if response and response.choices and len(response.choices) > 0:
                description = response.choices[0].message.content
                if description:
                    print("Description received from OpenAI API.")
                    return description
                else:
                    print("Warning: Received empty description from OpenAI API")
                    return "No description available"
            else:
                print("Warning: Unexpected response format from OpenAI API")
                return "Error: Unexpected response format"
                
        except Exception as api_error:
            print(f"Error during OpenAI API call: {str(api_error)}")
            return f"Error during API call: {str(api_error)}"
            
    except Exception as e:
        print(f"Error during GPT-4o processing: {str(e)}")
        return f"Error during processing: {str(e)}"

def generate_description_gemini(image_path, hint, context_text="No additional context available."):
    """
    Generates a description for a PNG image using the Gemini API.
    """
    img = None
    try:
        print(f"Loading image for Gemini description: {os.path.basename(image_path)}...")
        try:
            img = Image.open(image_path)
            img.verify()  # Verify integrity
            img.close()   # Close and reopen to reset file pointer
            img = Image.open(image_path)
        except Exception as img_err:
            return None, f"Error: Cannot load or invalid image file: {img_err}"

        prompt_text = (
            f"Analyze the provided image {hint} and its accompanying text context. "
            "Provide a concise and accurate description of what the image illustrates based on its relation to the text. "
            "Focus ONLY on describing the image content. Avoid questions or instructions."
        )
        if not context_text.strip():
            context_text = "No text context available."
        else:
            max_context_length = 8000
            if len(context_text) > max_context_length:
                context_text = context_text[:max_context_length] + "\n... [Context Truncated]"
            context_text = f"Text context:\n---\n{context_text}\n---"
        prompt_parts = [prompt_text, img, context_text]

        print("Sending request to Gemini API for description...")
        description, error = safe_llm_call("gemini-2.0-flash", prompt_parts, is_vision_model=True, config=generation_config)
        img.close()
        if error:
            print(f"Failed to generate description via Gemini API: {error}")
            return None, error
        else:
            print("Description received from Gemini API.")
            return description, None
    except Exception as e:
        if img:
            try:
                img.close()
            except Exception:
                pass
        return None, f"Error during Gemini image processing: {e}"

def extract_images_from_ppt(ppt_path, output_folder):
    """
    Extracts images from a PPTX file, converts non-GIF images to PNG,
    and generates a description for each image using the appropriate API based on image type.
    The final output is saved as a CSV file in the format:
       Page_of_PPT, image_filename, ppt_name, description
    """
    if not PIL_INSTALLED:
        raise ImportError("Pillow library is required for image conversion. Please install it: pip install Pillow")
    if not ppt_path.lower().endswith(".pptx"):
        raise ValueError("Only .pptx files are supported.")
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PPTX file not found at: {ppt_path}")

    os.makedirs(output_folder, exist_ok=True)
    ppt_name = os.path.basename(ppt_path)
    image_to_slide_map = {}
    slide_num_pattern = re.compile(r'slide(\d+)\.xml\.rels')

    try:
        with zipfile.ZipFile(ppt_path, 'r') as ppt_zip:
            # Map media images to slide numbers
            rel_files = [f for f in ppt_zip.namelist() if f.startswith("ppt/slides/_rels/slide") and f.endswith(".xml.rels")]
            for rel_file in rel_files:
                match = slide_num_pattern.search(os.path.basename(rel_file))
                if not match:
                    continue
                slide_number = int(match.group(1))
                try:
                    rel_xml_content = ppt_zip.read(rel_file)
                    root = ET.fromstring(rel_xml_content)
                    namespaces = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
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

            # Prepare CSV rows; header: Page_of_PPT, image_filename, ppt_name, description.
            csv_rows = [["Page_of_PPT", "image_filename", "ppt_name", "description"]]
            extracted_count = 0
            skipped_count = 0

            # Extract images from ppt/media.
            for file in ppt_zip.namelist():
                if file.startswith("ppt/media/"):
                    original_filename = os.path.basename(file)
                    base_name, original_ext = os.path.splitext(original_filename)
                    original_ext_lower = original_ext.lower()
                    slide_num = image_to_slide_map.get(file)
                    slide_prefix = f"slide{slide_num}_" if slide_num is not None else "slideUNK_"

                    supported_input_formats = ['.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp']
                    if original_ext_lower == '.gif' or original_ext_lower in supported_input_formats:
                        try:
                            with ppt_zip.open(file) as source:
                                image_data = source.read()
                            if not image_data:
                                print(f"Skipped (Empty File): {original_filename}")
                                skipped_count += 1
                                continue

                            # Process GIF images with the new OpenAI API.
                            if original_ext_lower == '.gif':
                                output_filename = f"{slide_prefix}{original_filename}"
                                output_path = os.path.join(output_folder, output_filename)
                                with open(output_path, 'wb') as target:
                                    target.write(image_data)
                                print(f"Extracted (GIF): {output_filename}")
                                hint = f"from slide {slide_num}" if slide_num is not None else "with unknown slide context"
                                description = generate_description_gpt4o(image_data, hint)
                            else:
                                # For non-GIF images, convert to PNG and use Gemini.
                                output_filename = f"{slide_prefix}{base_name}.png"
                                output_path = os.path.join(output_folder, output_filename)
                                try:
                                    img = Image.open(io.BytesIO(image_data))
                                    if img.mode in ['P', 'CMYK']:
                                        img = img.convert('RGBA')
                                    elif img.mode not in ['RGB', 'RGBA', 'L']:
                                        img = img.convert('RGB')
                                    img.save(output_path, format='PNG')
                                    print(f"Extracted (Converted to PNG): {output_filename}")
                                    img.close()
                                except Exception as img_err:
                                    print(f"Warning: Failed to convert '{original_filename}' to PNG. Error: {img_err}. Skipping.")
                                    skipped_count += 1
                                    continue
                                hint = f"from slide {slide_num}" if slide_num is not None else "with unknown slide context"
                                context_text = "No additional text context available."
                                desc_result, err = generate_description_gemini(output_path, hint, context_text)
                                if err:
                                    description = f"Error: {err}"
                                else:
                                    description = desc_result

                            extracted_count += 1
                            page_of_ppt = slide_num if slide_num is not None else "UNK"
                            csv_rows.append([page_of_ppt, output_filename, ppt_name, description])
                        except Exception as e:
                            print(f"Error processing file {file} ('{original_filename}'): {e}. Skipping.")
                            skipped_count += 1
                    else:
                        print(f"Skipped (Unsupported format {original_ext}): {original_filename}")
                        skipped_count += 1

            if extracted_count == 0 and skipped_count == 0:
                print("No media files found in ppt/media/ directory or no image relationships detected.")
            elif extracted_count == 0 and skipped_count > 0:
                print(f"No images successfully extracted. Skipped {skipped_count} files.")
            else:
                print("\nExtraction complete.")
                print(f" - Successfully extracted: {extracted_count} images (saved as PNG or GIF)")
                if skipped_count > 0:
                    print(f" - Skipped: {skipped_count} files (unsupported format, empty, or conversion error)")
                print(f"Output folder: {output_folder}")

            # Write CSV output.
            csv_filename = os.path.join(output_folder, "descriptions.csv")
            try:
                with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerows(csv_rows)
                print(f"CSV file saved: {csv_filename}")
            except Exception as csv_err:
                print(f"Error writing CSV file: {csv_err}")
    except zipfile.BadZipFile:
        print(f"Error: The file '{ppt_path}' is not a valid zip file or is corrupted.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# --- Configuration ---
ppt_path = r"D:\internships\myedu\pp_timageext\03_Biology\Chapter_6_Life Processes\Lecture_1\Module_1.pptx"  # Replace with your PPTX file path
output_folder = r"D:\internships\myedu\pp_timageext\outputs"     # Replace with your desired output folder

try:
    extract_images_from_ppt(ppt_path, output_folder)
except (ValueError, FileNotFoundError, ImportError) as e:
    print(f"Error: {e}")
