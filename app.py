from fastapi import FastAPI, Request
from email_utils import send_email_with_ppt
from fastapi.responses import JSONResponse
from fastapi.responses import JSONResponse
from pptx import Presentation
import os, re, json, requests
from io import BytesIO
from docx import Document
import fitz  # PyMuPDF
from zipfile import ZipFile
from pptx.util import Inches

app = FastAPI()

TEMPLATE_PATH ="template3.pptx"
BTEMPLATE_PATH="btemplate.pptx"
OUTPUT_PATH = "output.pptx"
BOUTPUT_PATH = "Boutput.pptx"
MONDAY_API_KEY = os.getenv("MONDAY_API_KEY", "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjU0NjI5MjM1NywiYWFpIjoxMSwidWlkIjo3NDc3Njk5NywiaWFkIjoiMjAyNS0wOC0wNFQwOTo0MzowNS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTIxNDMyMDQsInJnbiI6InVzZTEifQ.yYeelRXHOZlaxwYHBAvi6eXRzD2fNn1H-jX-Pd8Ukcw")
MONDAY_API_URL = "https://api.monday.com/v2"
STYLE_FOLDER = "styleGuide"



def get_file_download_url(asset_id: int) -> str:
    """
    Get the actual downloadable S3 URL for a Monday.com file using the API
    """
    query = """
    query($asset_id: [ID!]) {
      assets(ids: $asset_id) {
        id
        name
        public_url
        file_extension
      }
    }
    """
    headers = {
        "Authorization": MONDAY_API_KEY, 
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.post(
            MONDAY_API_URL, 
            json={
                "query": query, 
                "variables": {"asset_id": str(asset_id)}
            }, 
            headers=headers, 
            timeout=15
        )
        response.raise_for_status()
        data = response.json()
        
        assets = data.get("data", {}).get("assets", [])
        if assets and len(assets) > 0:
            public_url = assets[0].get("public_url")
            if public_url:
                print(f"✅ Got S3 URL for asset {asset_id}: {public_url}")
                return public_url
        
        print(f"⚠️ No public URL found for asset {asset_id}")
        return f"https://api.monday.com/v2/file/{asset_id}"  # Fallback
        
    except Exception as e:
        print(f"⚠️ Failed to get URL for asset {asset_id}: {e}")
        return f"https://api.monday.com/v2/file/{asset_id}"  

# ---- PDF / DOCX / ZIP Image Extraction ----

def extract_images_from_pdf(pdf_bytes_or_path):
    images = []
    doc = fitz.open(stream=pdf_bytes_or_path, filetype="pdf") if isinstance(pdf_bytes_or_path, bytes) else fitz.open(pdf_bytes_or_path)
    for page in doc:
        for img in page.get_images(full=True):
            xref = img[0]
            base_image = doc.extract_image(xref)
            images.append(f"https://api.monday.com/v2/file/{base_image['xref']}")  # Placeholder URL
    return images

def extract_images_from_docx(docx_bytes_or_path):
    images = []
    doc = Document(BytesIO(docx_bytes_or_path)) if isinstance(docx_bytes_or_path, bytes) else Document(docx_bytes_or_path)
    for i, rel in enumerate(doc.part._rels.values()):
        if "image" in rel.target_ref:
            images.append(f"https://api.monday.com/v2/file/extracted_{i}")  # Placeholder URL
    return images

def extract_images_from_zip(zip_bytes_or_path):
    images = []
    zip_file = ZipFile(BytesIO(zip_bytes_or_path)) if isinstance(zip_bytes_or_path, bytes) else ZipFile(zip_bytes_or_path)
    for file_name in zip_file.namelist():
        if file_name.lower().endswith((".jpg", ".jpeg", ".png")):
            images.append(f"https://api.monday.com/v2/file/zip_{file_name}")
        elif file_name.lower().endswith(".pdf"):
            images.extend(extract_images_from_pdf(BytesIO(zip_file.read(file_name))))
        elif file_name.lower().endswith(".docx"):
            images.extend(extract_images_from_docx(BytesIO(zip_file.read(file_name))))
    return images

def normalize_style_name(style_name: str) -> str:
    """Normalize style names to safe filenames"""
    clean = style_name.strip().lower()
    clean = re.sub(r"[ \-/]+", "_", clean)      # spaces, slashes, dashes -> underscore
    clean = re.sub(r"[^a-z0-9_]", "", clean)    # remove any other weird chars
    return clean

def get_style_images(selected_styles):
    style_images = {}
    if not selected_styles:
        return style_images

    for i, style in enumerate(selected_styles, start=1):
        normalized = normalize_style_name(style)
        for ext in [".jpg", ".jpeg", ".png"]:
            filename = f"{normalized}{ext}"       
            path = os.path.join(STYLE_FOLDER, filename)
            if os.path.exists(path):
                style_images[f"{{{{style{i}}}}}"] = path
                print(f"🎨 Style matched: {style} -> {path}")
                break
        else:
            print(f"⚠️ Missing image for style: {style} (expected something like {normalized}.png)")

    return style_images

def map_webhook_to_form(event: dict):
    col = event.get("columnValues", {})

    # Extract image URLs from Monday "files" type column 
    image_urls = []
    if "files" in col and col["files"].get("value"):
        try:
            file_data = json.loads(col["files"]["value"])  # Parse JSON from Monday
            image_urls = [asset["url"] for asset in file_data]  # Collect URLs
        except Exception:
            image_urls = []

    return {
    # --- Project Information ---
    "Project Name": event.get("pulseName"),
    "**Project Name**": event.get("pulseName"),
    "project name": event.get("pulseName"),

    "What is the nature of your project?": col.get("status1", {}).get("label", {}).get("text"),
    "what is the nature of your project?": col.get("status1", {}).get("label", {}).get("text"),

    "What is the property type?": col.get("dropdown76", {}).get("chosenValues", [{}])[0].get("name"),
    "what is the property type?": col.get("dropdown76", {}).get("chosenValues", [{}])[0].get("name"),

    "City": col.get("text8", {}).get("value"),
    "city": col.get("text8", {}).get("value"),

    "Country": col.get("country6", {}).get("countryName"),
    "country": col.get("country6", {}).get("countryName"),

    "Space(S) to be designed": ", ".join([v.get("name", "") for v in col.get("dropdown0", {}).get("chosenValues", [])]),
    "space(s) to be designed": ", ".join([v.get("name", "") for v in col.get("dropdown0", {}).get("chosenValues", [])]),

    "What is the area size?": col.get("short_text8fr4spel", {}).get("value"),
    "what is the area size?": col.get("short_text8fr4spel", {}).get("value"),

    # --- Client Information ---
    "How old are you?": col.get("status", {}).get("label", {}).get("text"),
    "how old are you?": col.get("status", {}).get("label", {}).get("text"),
    "{{How old are you?}}": col.get("status", {}).get("label", {}).get("text"),
    "{{how old are you?}}": col.get("status", {}).get("label", {}).get("text"),

    "How many people will live in the space?": col.get("text1", {}).get("value"),
    "how many people will live in the space?": col.get("text1", {}).get("value"),
    "{{How many people will live in the space?}}": col.get("text1", {}).get("value"),
    "{{how many people will live in the space?}}": col.get("text1", {}).get("value"),

    "What best describes your situation?": col.get("single_selecti4d0sw1", {}).get("label", {}).get("text"),
    "what best describes your situation?": col.get("single_selecti4d0sw1", {}).get("label", {}).get("text"),
    "{{What best describes your situation?}}": col.get("single_selecti4d0sw1", {}).get("label", {}).get("text"),
    "{{what best describes your situation?}}": col.get("single_selecti4d0sw1", {}).get("label", {}).get("text"),

    "Do you have any pets?": col.get("text_1", {}).get("value"),
    "do you have any pets?": col.get("text_1", {}).get("value"),
    "{{Do you have any pets?}}": col.get("text_1", {}).get("value"),
    "{{do you have any pets?}}": col.get("text_1", {}).get("value"),

    "Please describe the scope of work": col.get("text37", {}).get("value"),
    "please describe the scope of work": col.get("text37", {}).get("value"),
    "{{Please describe the scope of work}}": col.get("text37", {}).get("value"),
    "{{please describe the scope of work}}": col.get("text37", {}).get("value"),

    # --- Other Information ---
    "Is there any other information you'd like us to know?": col.get("long_text3", {}).get("text"),
    "is there any other information you'd like us to know?": col.get("long_text3", {}).get("text"),
    "is there any other information you'd like us to know": col.get("long_text3", {}).get("text"),
    "{{Is there any other information you'd like us to know?}}": col.get("long_text3", {}).get("text"),
    "{{is there any other information you'd like us to know?}}": col.get("long_text3", {}).get("text"),
    "{{is there any other information you'd like us to know}}": col.get("long_text3", {}).get("text"),
    "Is there any other information you'd like us to know": col.get("long_text3", {}).get("text"),
    "other information you'd like us to know": col.get("long_text3", {}).get("text"),
    "other information": col.get("long_text3", {}).get("text"),

    # --- Words describe ---
    "Words Describe the feel for space?": col.get("short_text5fonuzuu", {}).get("value"),
    "words describe the feel for space?": col.get("short_text5fonuzuu", {}).get("value"),
    "words describe the feel for space": col.get("short_text5fonuzuu", {}).get("value"),
    "**Words Describe the feel for space?**": col.get("short_text5fonuzuu", {}).get("value"),
    "**words describe the feel for space?**": col.get("short_text5fonuzuu", {}).get("value"),
    "**{{ Words Describe the feel for space? }}**": col.get("short_text5fonuzuu", {}).get("value"),
    "{{ Words Describe the feel for space? }}": col.get("short_text5fonuzuu", {}).get("value"),
    "{{Words Describe the feel for space?}}": col.get("short_text5fonuzuu", {}).get("value"),
    "{{words describe the feel for space?}}": col.get("short_text5fonuzuu", {}).get("value"),
    "{{words describe the feel for space}}": col.get("short_text5fonuzuu", {}).get("value"),
    "Words Describe the feel for space": col.get("short_text5fonuzuu", {}).get("value"),
    "words describe the feel for space": col.get("short_text5fonuzuu", {}).get("value"),
}

def fetch_user_details(email: str):
    try:
        url = f"https://tldp.omrsolutions.com/get_user_details.php?email={email}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print("Error fetching user details:", e)
    return {}

def create_short_ppt(template_path: str, output_path: str, form_data: dict):
    """
    Generate a short PPT with only limited fields.
    """
    # ✅ Build only required text mapping (removed image mappings)
    text_map = {
        "Project Name": form_data.get("Project Name"),
        "Project Type": form_data.get("What is the nature of your project?"),
        "space To be Designed": form_data.get("Space(S) to be designed"),
        "Room size": form_data.get("What is the area size?"),
        "style(s) selected": form_data.get("Which style(s) do you like?"),
        "Location": f"{form_data.get('City', '')}, {form_data.get('Country', '')}",
    }

    # ✅ Generate PPT (only text replacement)
    replace_text_in_ppt(template_path, output_path, text_map)

    print(f"✅ Short PPT generated: {output_path}")
    return output_path

def get_item_files(item_id: int):
    """
    Get all files for an item with their public URLs
    Fixed GraphQL query syntax
    """
    query = """
    query($item_ids: [ID!]!) {
      items(ids: $item_ids) {
        id
        name
        assets {
          id
          name
          public_url
          file_extension
          url
        }
      }
    }
    """
    headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json"}
    try:
        response = requests.post(
            MONDAY_API_URL, 
            json={
                "query": query, 
                "variables": {"item_ids": [str(item_id)]}  # Changed to array format
            }, 
            headers=headers, 
            timeout=15
        )
        response.raise_for_status()
        data = response.json()
        
        if "errors" in data:
            print(f"⚠️ GraphQL errors for item {item_id}: {data['errors']}")
            return []
            
    except Exception as e:
        print(f"⚠️ Request/JSON error: {e}")
        return []
        
    items = data.get("data", {}).get("items", [])
    if not items:
        return []
        
    assets = items[0].get("assets", [])
    result = []
    
    for asset in assets:
        public_url = asset.get("public_url") or asset.get("url")
        if public_url and public_url != "null":
            result.append({
                "id": asset["id"], 
                "name": asset["name"], 
                "url": public_url, 
                "ext": asset["file_extension"]
            })
    
    return result
def replace_text_in_ppt(template_path: str, output_path: str, text_map: dict):
    """
    Replace only text placeholders in PPT (no image insertion)
    """
    prs = Presentation(template_path)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                for run in p.runs:
                    for placeholder, value in text_map.items():
                        if not value:
                            continue
                        if placeholder.lower() in run.text.lower():
                            run.text = run.text.replace(placeholder, str(value))
    
    prs.save(output_path)
    return output_path

def categorize_and_collect_images(event: dict) -> dict:
    """
    Categorize images from Monday.com files based on column source.
    - First tries to fetch a public_url (S3).
    - If no public_url, downloads via Monday API and extracts images locally.
    """
    categorized_images = {
        "floor_plans": [],
        "elevation_drawings": [],
        "existing_pictures": [],
        "inspiration_images": []
    }
    
    col_vals = event.get("columnValues", {})
    print(f"🔍 Found {len(col_vals)} column values to process")
    
    # Map file columns to categories
    column_category_map = {
        "files": "floor_plans",
        "fileb3p8t108": "elevation_drawings", 
        "fileh7us51cr": "existing_pictures",
        "files3": "inspiration_images"
    }
    
    for column_name, category in column_category_map.items():
        print(f"🔍 Checking column: {column_name} -> {category}")
        
        if column_name not in col_vals:
            print(f"⚠️ Column {column_name} not found in webhook data")
            continue

        file_data = col_vals[column_name]
        if not (isinstance(file_data, dict) and "files" in file_data):
            print(f"⚠️ Unexpected structure in {column_name}: {file_data}")
            continue

        files_list = file_data["files"]
        print(f"📁 Found {len(files_list)} files in {column_name}")

        for file_info in files_list:
            if not isinstance(file_info, dict):
                print(f"⚠️ File info is not a dict: {file_info}")
                continue

            asset_id = file_info.get("assetId") or file_info.get("id")
            filename = file_info.get("name", "")
            file_ext = file_info.get("extension", "").lower()

            if not asset_id or not filename:
                print(f"⚠️ Missing asset_id or filename: {file_info}")
                continue

            print(f"🔍 Processing file: {filename} (ID: {asset_id}, Ext: {file_ext})")

            try:
                # 1. Try to get a direct public URL
                public_url = get_file_download_url(asset_id)

                if public_url and ("amazonaws.com" in public_url or "s3" in public_url):
                    categorized_images[category].append(public_url)
                    print(f"✅ Public URL added to {category}: {public_url}")
                    continue

                # 2. If no public URL, download via Monday API
                print(f"⬇️ Downloading file {filename} from Monday API (no public_url)")
                headers = {"Authorization": MONDAY_API_KEY}
                download_url = f"{MONDAY_API_URL}/file/{asset_id}"
                response = requests.get(download_url, headers=headers, timeout=20)
                response.raise_for_status()
                file_bytes = response.content

                # 3. Extract images based on file type
                extracted = []
                if file_ext in ["jpg", "jpeg", "png"]:
                    # Save raw image as a temporary file path
                    temp_path = f"temp_{asset_id}.{file_ext}"
                    with open(temp_path, "wb") as f:
                        f.write(file_bytes)
                    extracted.append(temp_path)

                elif file_ext == "pdf":
                    extracted = extract_images_from_pdf(file_bytes)

                elif file_ext == "docx":
                    extracted = extract_images_from_docx(file_bytes)

                elif file_ext == "zip":
                    extracted = extract_images_from_zip(file_bytes)

                else:
                    print(f"⚠️ Unsupported file type: {filename}")
                
                if extracted:
                    categorized_images[category].extend(extracted)
                    print(f"✅ Extracted {len(extracted)} images from {filename}")

            except Exception as e:
                print(f"⚠️ Failed to process {filename}: {e}")

    # Remove empty categories
    categorized_images = {k: v for k, v in categorized_images.items() if v}

    print("📁 Image categorization complete:")
    for category, urls in categorized_images.items():
        print(f"  {category}: {len(urls)} images")
        for i, url in enumerate(urls, 1):
            print(f"    {i}. {url}")

    return categorized_images


# Also, let's improve the get_file_download_url function for better error handling
def get_file_download_url(asset_id: int) -> str:
    """
    Get the actual downloadable S3 URL for a Monday.com file using the API
    """
    query = """
    query($asset_ids: [ID!]!) {
      assets(ids: $asset_ids) {
        id
        name
        public_url
        file_extension
        url
      }
    }
    """
    headers = {
        "Authorization": MONDAY_API_KEY, 
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.post(
            MONDAY_API_URL, 
            json={
                "query": query, 
                "variables": {"asset_ids": [str(asset_id)]}  # ✅ Correct: list format
            }, 
            headers=headers, 
            timeout=15
        )
        response.raise_for_status()
        data = response.json()
        
        if "errors" in data:
            print(f"⚠️ GraphQL errors for asset {asset_id}: {data['errors']}")
            return None
            
        assets = data.get("data", {}).get("assets", [])
        if assets:
            asset = assets[0]
            public_url = asset.get("public_url") or asset.get("url")
            if public_url and public_url != "null":
                return public_url
        
        return None
        
    except Exception as e:
        print(f"⚠️ Failed to get URL for asset {asset_id}: {e}")
        return None










def replace_placeholders_with_images(pptx_path, output_path, categorized_images):
    """
    Replace placeholders like {{Image1}}, {{Layout2}}, {{Elevation1}}, {{Inspiration1}}, {{Existing1}}
    with categorized images. 
    If no image is found for any placeholder and no images remain in the slide, delete the slide.
    """
    prs = Presentation(pptx_path)

    # Build category mapping
    category_map = {
        "Image": sum(categorized_images.values(), []),  # All images flattened
        "Layout": categorized_images.get("floor_plans", []),
        "Elevation": categorized_images.get("elevation_drawings", []),
        "Inspiration": categorized_images.get("inspiration_images", []),
        "Existing": categorized_images.get("existing_pictures", []),
    }

    slides_to_delete = []

    for slide_idx, slide in enumerate(prs.slides):
        replaced_any = False
        found_placeholder = False

        for shape in list(slide.shapes):
            if not shape.has_text_frame:
                continue

            text = shape.text.strip()

            # Match placeholders like {{CategoryN}}
            match = re.match(r"\{\{(\w+)(\d+)\}\}", text)
            if not match:
                continue

            found_placeholder = True
            category, num = match.group(1), int(match.group(2)) - 1
            images = category_map.get(category)

            if images and num < len(images):
                img_path = images[num]
                print(f"✅ Replacing {text} with {img_path}")

                # Download if URL, else use local path
                if isinstance(img_path, str) and img_path.startswith("http"):
                    try:
                        resp = requests.get(img_path, timeout=20)
                        resp.raise_for_status()
                        img_file = BytesIO(resp.content)
                    except Exception as e:
                        print(f"⚠️ Failed to download {img_path}: {e}")
                        continue
                else:
                    img_file = img_path

                # Replace placeholder with image
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                sp = shape._element
                sp.getparent().remove(sp)
                slide.shapes.add_picture(img_file, left, top, width, height)

                replaced_any = True
            else:
                print(f"⚠️ No image available for {text}, placeholder kept")

        # If slide had placeholders but no image was inserted, and no picture shapes exist → mark for deletion
        if found_placeholder and not replaced_any:
            has_pictures = any(shape.shape_type == 13 for shape in slide.shapes)  # 13 = picture
            if not has_pictures:
                print(f"🗑️ Marking slide {slide_idx+1} for deletion (no images found)")
                slides_to_delete.append(slide)

    # Delete marked slides
    for slide in slides_to_delete:
        prs.slides._sldIdLst.remove(slide._element)

    prs.save(output_path)
    print(f"💾 Saved updated presentation to {output_path}")
    return output_path

@app.post("/monday-webhook")
async def monday_webhook(request: Request):
    body = await request.json()
    if "challenge" in body:
        return JSONResponse(content={"challenge": body["challenge"]})
    if "event" in body:
        event = body["event"]
        item_id = event.get("pulseId")
        form_data = map_webhook_to_form(event)

        # --- Style Images ---
        styles_raw = None
        col_vals = event.get("columnValues", {})
        if "dropdown" in col_vals:  
            chosen = col_vals["dropdown"].get("chosenValues", [])
            styles_raw = [c["name"] for c in chosen]
        if styles_raw:
            form_data.update(get_style_images(styles_raw))
            form_data["Which style(s) do you like?"] = ", ".join(styles_raw)

        # --- Email Extraction ---
        email = None
        if "email" in col_vals:
            email_block = col_vals["email"]
            if isinstance(email_block, dict):
                email = email_block.get("email") or email_block.get("text")
        if not email:
            email = form_data.get("Email") or form_data.get("email") or "krgarav@gmail.com"

        # --- Fetch extra details ---
        user_details = fetch_user_details(email)
        if user_details.get("status") == "success":
            qd = user_details["data"][0]["quotationdetails"]
            if qd.get("project_type"):
                form_data["What is the property type?"] = qd["project_type"]
            if qd.get("area_size"):
                form_data["What is the area size?"] = qd["area_size"]
            if qd.get("project_name"):
                form_data["Project Name"] = qd["project_name"]
            if qd.get("residential_type"):
                form_data["What is the nature of your project?"] = qd["residential_type"]

        # --- Categorize and collect all images ---
        categorized_images = categorize_and_collect_images(event)
        
        # Print the categorized images in the requested format
        print("📸 Categorized Images:")
        print(json.dumps(categorized_images, indent=2))
        # --- Insert floor plans into PPT if available ---

        # Step 1: Replace text placeholders
        replace_text_in_ppt(TEMPLATE_PATH, OUTPUT_PATH, form_data)

        # Step 2: Replace image placeholders (works on already text-updated PPT)
        replace_placeholders_with_images(OUTPUT_PATH, OUTPUT_PATH, categorized_images)


        
        # Return response with categorized images
        return {"status": "ok", "categorized_images": categorized_images}

    return {"status": "ok"}