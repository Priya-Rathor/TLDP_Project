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
import time
import threading
from fastapi import BackgroundTasks

app = FastAPI()

TEMPLATE_PATH ="template.pptx"
BTEMPLATE_PATH="btemplate.pptx"
OUTPUT_PATH = "Presentation.pptx"
BOUTPUT_PATH = "Brochure.pptx"
MONDAY_API_KEY = os.getenv("MONDAY_API_KEY", "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjU0NjI5MjM1NywiYWFpIjoxMSwidWlkIjo3NDc3Njk5NywiaWFkIjoiMjAyNS0wOC0wNFQwOTo0MzowNS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTIxNDMyMDQsInJnbiI6InVzZTEifQ.yYeelRXHOZlaxwYHBAvi6eXRzD2fNn1H-jX-Pd8Ukcw")
MONDAY_API_URL = "https://api.monday.com/v2"
STYLE_FOLDER = "styleGuide"

# ---- Email Tracking to Prevent Duplicates ----
email_sent_tracker = {}
email_lock = threading.Lock()

def mark_email_sent(event_id: str):
    """Mark that email has been sent for this event"""
    with email_lock:
        email_sent_tracker[event_id] = time.time()
        
        # Clean up old entries (older than 1 hour)
        now = time.time()
        to_remove = [k for k, v in email_sent_tracker.items() if now - v > 3600]
        for k in to_remove:
            email_sent_tracker.pop(k, None)

def is_email_already_sent(event_id: str) -> bool:
    """Check if email was already sent for this event"""
    with email_lock:
        return event_id in email_sent_tracker

# ---- PDF / DOCX / ZIP Image Extraction ----

def extract_images_from_pdf(pdf_bytes_or_path):
    images = []
    doc = fitz.open(stream=pdf_bytes_or_path, filetype="pdf") if isinstance(pdf_bytes_or_path, bytes) else fitz.open(pdf_bytes_or_path)
    for page in doc:
        for img in page.get_images(full=True):
            xref = img[0]
            base_image = doc.extract_image(xref)
            images.append(BytesIO(base_image["image"]))
    return images

def extract_images_from_docx(docx_bytes_or_path):
    images = []
    doc = Document(BytesIO(docx_bytes_or_path)) if isinstance(docx_bytes_or_path, bytes) else Document(docx_bytes_or_path)
    for rel in doc.part._rels.values():
        if "image" in rel.target_ref:
            images.append(BytesIO(rel.target_part.blob))
    return images

def extract_images_from_zip(zip_bytes_or_path):
    images = []
    zip_file = ZipFile(BytesIO(zip_bytes_or_path)) if isinstance(zip_bytes_or_path, bytes) else ZipFile(zip_bytes_or_path)
    for file_name in zip_file.namelist():
        if file_name.lower().endswith((".jpg", ".jpeg", ".png")):
            images.append(BytesIO(zip_file.read(file_name)))
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
        "How old are you?": col.get("dropdown52", {}).get("chosenValues", [{}])[0].get("name"),
        "how old are you?": col.get("dropdown52", {}).get("chosenValues", [{}])[0].get("name"),
        "{{How old are you?}}": col.get("dropdown52", {}).get("chosenValues", [{}])[0].get("name"),
        "{{how old are you?}}": col.get("dropdown52", {}).get("chosenValues", [{}])[0].get("name"),

        "How many people will live in the space?": col.get("short_textzm6bosrr", {}).get("value"),
        "how many people will live in the space?": col.get("short_textzm6bosrr", {}).get("value"), 
        "{{How many people will live in the space?}}": col.get("short_textzm6bosrr", {}).get("value"),
        "{{how many people will live in the space?}}": col.get("short_textzm6bosrr", {}).get("value"),

        "What best describes your situation?": col.get("long_textcl38cdjs", {}).get("text"),
        "what best describes your situation?": col.get("long_textcl38cdjs", {}).get("text"),
        "{{What best describes your situation?}}": col.get("long_textcl38cdjs", {}).get("text"),
        "{{what best describes your situation?}}": col.get("long_textcl38cdjs", {}).get("text"),

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
        "Words Describe the feel for space?": col.get("short_text5pym043c", {}).get("value"),
        "words describe the feel for space?": col.get("short_text5pym043c", {}).get("value"),
        "words describe the feel for space": col.get("short_text5pym043c", {}).get("value"),
        "**Words Describe the feel for space?**": col.get("short_text5pym043c", {}).get("value"),
        "**words describe the feel for space?**": col.get("short_text5pym043c", {}).get("value"),
        "**{{ Words Describe the feel for space? }}**": col.get("short_text5pym043c", {}).get("value"),
        "{{ Words Describe the feel for space? }}": col.get("short_text5pym043c", {}).get("value"),
        "{{Words Describe the feel for space?}}": col.get("short_text5pym043c", {}).get("value"),
        "{{words describe the feel for space?}}": col.get("short_text5pym043c", {}).get("value"),
        "{{words describe the feel for space}}": col.get("short_text5pym043c", {}).get("value"),
        "Words Describe the feel for space": col.get("short_text5pym043c", {}).get("value"),
        "words describe the feel for space": col.get("short_text5pym043c", {}).get("value"),
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

    # ✅ Build only required text mapping
    text_map = {
        "Project Name": form_data.get("Project Name"),
        "Project Type": form_data.get("What is the nature of your project?"),
        "space To be Designed": form_data.get("Space(S) to be designed"),
        "Room size": form_data.get("What is the area size?"),
        "style(s) selected": form_data.get("Which style(s) do you like?"),
        "Location": f"{form_data.get('City', '')}, {form_data.get('Country', '')}",

        "{{Image1}}": form_data.get("{{Image1}}"),
        "{{Image2}}": form_data.get("{{Image2}}"),
        "{{Image3}}": form_data.get("{{Image3}}"),
        "{{Image4}}": form_data.get("{{Image4}}"),
        "{{Image5}}": form_data.get("{{Image5}}"),
        "{{Image6}}": form_data.get("{{Image6}}"),
        "{{Image7}}": form_data.get("{{Image7}}"),
    }

    # ✅ Generate PPT
    replace_text_and_images_in_ppt(template_path, output_path, text_map)

    print(f"✅ Short PPT generated: {output_path}")
    return output_path

def get_item_files(item_id: int):
    query = """
    query($item_id: [ID!]) {
      items(ids: $item_id) {
        assets {
          id
          name
          public_url
          file_extension
        }
      }
    }
    """
    headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json"}
    try:
        response = requests.post(MONDAY_API_URL, json={"query": query, "variables": {"item_id": str(item_id)}}, headers=headers, timeout=15)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print(f"⚠️ Request/JSON error: {e}")
        return []
    items = data.get("data", {}).get("items", [])
    if not items:
        return []
    assets = items[0].get("assets", [])
    return [{"id": a["id"], "name": a["name"], "url": a["public_url"], "ext": a["file_extension"]} for a in assets if a.get("public_url")]

# ✅ Replace placeholders in PPT and add images
def replace_text_placeholders(prs: Presentation, text_map: dict):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                for run in p.runs:
                    for placeholder, value in text_map.items():
                        if not value:
                            continue
                        if placeholder.lower() in run.text.lower() and not placeholder.startswith("{{Image"):
                            run.text = run.text.replace(placeholder, str(value))

def normalize(text: str) -> str:
    return text.strip().lower().replace(" ", "").replace("{", "").replace("}", "")

def replace_image_placeholders(prs: Presentation, image_map: dict):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            full_text = "".join([p.text for p in shape.text_frame.paragraphs])
            normalized_shape_text = normalize(full_text)
            for placeholder, source in image_map.items():
                if not source: continue
                if normalize(placeholder) == normalized_shape_text:
                    try:
                        if isinstance(source, BytesIO):
                            image_stream = source
                        elif source.startswith("http"):
                            resp = requests.get(source, stream=True, timeout=15)
                            resp.raise_for_status()
                            image_stream = BytesIO(resp.content)
                        else:
                            image_stream = open(source, "rb")

                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        for p in shape.text_frame.paragraphs:
                            for run in p.runs:
                                run.text = ""
                        slide.shapes.add_picture(image_stream, left, top, width=width, height=height)
                        if not isinstance(source, BytesIO) and not source.startswith("http"):
                            image_stream.close()
                        break
                    except Exception as e:
                        print(f"⚠️ Failed to insert {placeholder}: {e}")

def replace_text_and_images_in_ppt(template_path, output_path, text_map: dict):
    prs = Presentation(template_path)
    image_only_map = {k:v for k,v in text_map.items() if k.startswith("{{Image") or k.startswith("{{style")}
    text_only_map = {k:v for k,v in text_map.items() if k not in image_only_map}
    replace_text_placeholders(prs, text_only_map)
    replace_image_placeholders(prs, image_only_map)
    prs.save(output_path)
    return output_path

def verify_files_exist(file_paths):
    """Verify that both PPT files exist and are valid"""
    for file_path in file_paths:
        if not os.path.exists(file_path):
            print(f"❌ File does not exist: {file_path}")
            return False
        
        # Check if file size is reasonable (not empty)
        file_size = os.path.getsize(file_path)
        if file_size < 1000:  # Less than 1KB might indicate an error
            print(f"❌ File too small (possible error): {file_path} ({file_size} bytes)")
            return False
            
    print("✅ Both PPT files verified successfully")
    return True

def send_email_safely(email: str, item_id: str, file_paths: list, event_id: str):
    """Send email only once and with proper verification"""
    try:
        # Double-check if email was already sent
        if is_email_already_sent(event_id):
            print(f"📧 Email already sent for event {event_id}, skipping...")
            return

        # Verify both files exist and are valid
        if not verify_files_exist(file_paths):
            print(f"❌ Files verification failed for {item_id}")
            return

        # Send email
        print(f"📧 Sending email to {email} for item {item_id}...")
        send_email_with_ppt(
            recipient=email,
            subject=f"Your Design Presentation - {item_id}",
            body="Hello! Here are your generated presentation and brochure. Thank you for choosing our services!",
            file_paths=file_paths
        )
        
        # Mark email as sent ONLY after successful send
        mark_email_sent(event_id)
        print(f"✅ Email successfully sent to {email} for item {item_id}")
        
    except Exception as e:
        print(f"❌ Email send failed for {item_id}: {e}")
        # Don't mark as sent if it failed

# --- WEBHOOK ENDPOINT ---

# --- Deduplication store (10 min memory cache, use Redis for production) ---
processed_events = {}

def is_duplicate(event_id: str) -> bool:
    now = time.time()
    # cleanup old
    for k, v in list(processed_events.items()):
        if now - v > 600:  # keep for 10 mins
            processed_events.pop(k, None)

    if event_id in processed_events:
        return True
    processed_events[event_id] = now
    return False

def process_event(event, email):
    item_id = event.get("pulseId")
    trigger_time = event.get("triggerTime")
    event_id = f"{item_id}_{trigger_time}"
    
    print(f"🔄 Processing event {event_id} for email {email}")
    
    # Check if email was already sent for this specific event
    if is_email_already_sent(event_id):
        print(f"📧 Email already sent for event {event_id}, stopping processing...")
        return

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

    # --- Get files & extract images ---
    assets = get_item_files(int(item_id))
    image_counter = 1
    for a in assets:
        url = a["url"]
        try:
            resp = requests.get(url, stream=True, timeout=15)
            resp.raise_for_status()
            content_type = resp.headers.get("Content-Type", "").lower()
            images_to_add = []

            if "image" in content_type:
                images_to_add.append(BytesIO(resp.content))
            elif "pdf" in content_type:
                images_to_add.extend(extract_images_from_pdf(resp.content))
            elif "word" in content_type or url.lower().endswith(".docx"):
                images_to_add.extend(extract_images_from_docx(resp.content))
            elif "zip" in content_type or url.lower().endswith(".zip"):
                images_to_add.extend(extract_images_from_zip(resp.content))

            for img_stream in images_to_add:
                form_data[f"{{{{Image{image_counter}}}}}"] = img_stream
                image_counter += 1
        except Exception as e:
            print(f"⚠️ Failed to process {url}: {e}")
        
        for i in range(image_counter,22):
            form_data[f"{{{{Image{i}}}}}"] = ""

    # --- Generate unique file names to avoid conflicts ---
    timestamp = int(time.time())
    unique_output_path = f"Presentation_{item_id}_{timestamp}.pptx"
    unique_brochure_path = f"Brochure_{item_id}_{timestamp}.pptx"

    try:
        # --- Generate PPTs ---
        print(f"🔧 Generating main presentation for {item_id}...")
        replace_text_and_images_in_ppt(TEMPLATE_PATH, unique_output_path, form_data)
        
        print(f"🔧 Generating brochure for {item_id}...")
        create_short_ppt(BTEMPLATE_PATH, unique_brochure_path, form_data)

        # --- Send Email with verification ---
        send_email_safely(email, item_id, [unique_output_path, unique_brochure_path], event_id)

        
            
    except Exception as e:
        print(f"❌ PPT generation failed for {item_id}: {e}")

@app.post("/monday-webhook")
async def monday_webhook(request: Request, background_tasks: BackgroundTasks):
    body = await request.json()
    print("🚀 Webhook received:", json.dumps(body, indent=2))

    if "challenge" in body:
        return JSONResponse(content={"challenge": body["challenge"]})

    if "event" in body:
        event = body["event"]
        item_id = event.get("pulseId")
        trigger_time = event.get("triggerTime")
        event_id = f"{item_id}_{trigger_time}"

        # ✅ Enhanced deduplication check
        if is_duplicate(event_id):
            print(f"⚠️ Duplicate webhook skipped: {event_id}")
            return {"status": "duplicate_skipped"}

        # ✅ Email-specific duplication check
        if is_email_already_sent(event_id):
            print(f"📧 Email already sent for event {event_id}")
            return {"status": "email_already_sent"}

        # --- Email Extraction ---
        col_vals = event.get("columnValues", {})
        email = None
        if "email" in col_vals:
            email_block = col_vals["email"]
            if isinstance(email_block, dict):
                email = email_block.get("email") or email_block.get("text")
        if not email:
            email = "priya.rathor.266393@gmail.com"

        print(f"📧 Processing for email: {email}, event: {event_id}")

        # ✅ Background task (avoid retries)
        background_tasks.add_task(process_event, event, email)

    return {"status": "ok"}