# brochure.py

from pptx import Presentation
from pptx.util import Inches
from datetime import datetime, timedelta
from PIL import Image, ImageDraw
import platform, os, requests

# -------------------------
# Constants
# -------------------------
if platform.system() == "Windows":
    DAY_FORMAT = "%#d-%b"
else:
    DAY_FORMAT = "%-d-%b"


# -------------------------
# Helper: Load image from URL or path
# -------------------------
def get_local_image(path_or_url, tmp_name="temp_img.png"):
    if not path_or_url:
        return None
    if isinstance(path_or_url, str) and path_or_url.startswith("http"):
        try:
            resp = requests.get(path_or_url, timeout=10)
            resp.raise_for_status()
            with open(tmp_name, "wb") as f:
                f.write(resp.content)
            return tmp_name
        except Exception as e:
            print(f"⚠️ Failed to download image {path_or_url}: {e}")
            return None
    return path_or_url if os.path.exists(path_or_url) else None


# -------------------------
# Helper: Make circular crop
# -------------------------
def make_circle_image(img_path, output_path="circle_image.png"):
    im = Image.open(img_path).convert("RGBA")
    bigsize = (im.size[0] * 3, im.size[1] * 3)
    mask = Image.new("L", bigsize, 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((0, 0) + bigsize, fill=255)
    mask = mask.resize(im.size, Image.LANCZOS)
    im.putalpha(mask)
    im.save(output_path, format="PNG")
    return output_path


def replace_with_circle_image(slide, img_path):
    cropped_path = make_circle_image(img_path)
    for shape in list(slide.shapes):
        if hasattr(shape, "text") and "{{image1}}" in shape.text:
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            slide.shapes._spTree.remove(shape._element)
            slide.shapes.add_picture(cropped_path, left, top, width, height)
    return slide


# -------------------------
# Calendar + background helpers
# -------------------------
def build_mapping(start_date=None):
    if start_date is None:
        start_date = datetime.today()
    mapping = {}
    for i in range(1, 8):
        date = start_date + timedelta(days=i - 1)
        mapping[f"{{{{day{i}}}}}"] = date.strftime("%a")
    for i in range(1, 50):
        date = start_date + timedelta(days=i - 1)
        mapping[f"{{{{d{i}}}}}"] = date.strftime(DAY_FORMAT)
    return mapping


def replace_text_in_frame(text_frame, mapping):
    for para in text_frame.paragraphs:
        for run in para.runs:
            for ph, val in mapping.items():
                if ph in run.text:
                    run.text = run.text.replace(ph, val)


def iter_shapes(shapes):
    for shp in shapes:
        yield shp
        if shp.shape_type == 6:  # group shape
            yield from iter_shapes(shp.shapes)


def update_calendar_with_bg(prs, slide, image_path, start_date=None):
    mapping = build_mapping(start_date)
    slide_width, slide_height = prs.slide_width, prs.slide_height

    if image_path:
        pic = slide.shapes.add_picture(image_path, 0, 0, width=slide_width, height=slide_height)
        slide.shapes._spTree.remove(pic._element)
        slide.shapes._spTree.insert(2, pic._element)

    for shp in iter_shapes(slide.shapes):
        if getattr(shp, "has_table", False):
            for row in shp.table.rows:
                for cell in row.cells:
                    replace_text_in_frame(cell.text_frame, mapping)
        elif getattr(shp, "has_text_frame", False):
            replace_text_in_frame(shp.text_frame, mapping)
    return slide


# -------------------------
# Layout + extra images
# -------------------------
def replace_layout_and_append_images(prs, slide, layout_img_path, bg_img_path, extra_images):
    slide_width, slide_height = prs.slide_width, prs.slide_height

    # Replace Layout placeholder
    for shape in list(slide.shapes):
        if hasattr(shape, "text") and "{{Layout1}}" in shape.text:
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            slide.shapes._spTree.remove(shape._element)
            if layout_img_path:
                slide.shapes.add_picture(layout_img_path, left, top, width, height)

    # Add background
    if bg_img_path:
        pic = slide.shapes.add_picture(bg_img_path, 0, 0, slide_width, slide_height)
        slide.shapes._spTree.remove(pic._element)
        slide.shapes._spTree.insert(2, pic._element)

    # Replace {{image1}}, {{image2}}, ... with extra images
    for idx, img_path in enumerate(extra_images, start=1):
        placeholder = f"{{{{image{idx}}}}}"
        for shape in list(slide.shapes):
            if hasattr(shape, "text") and placeholder in shape.text:
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture(img_path, left, top, width, height)

    return slide



# -------------------------
# Text replacement helper
# -------------------------
def replace_text_in_ppt(slide, text_map):
    for shp in iter_shapes(slide.shapes):
        if getattr(shp, "has_text_frame", False):
            for para in shp.text_frame.paragraphs:
                for run in para.runs:
                    for ph, val in text_map.items():
                        if ph in run.text:
                            run.text = run.text.replace(ph, val)
    return slide


# -------------------------
# Main: Create Brochure
# -------------------------
def create_brochure_ppt(template_path, output_path, form_data,
                        circle_img, calendar_bg, layout_img, layout_bg,
                        categorized_images):
    """
    Create a brochure PPT with:
    1st Page  -> Replace text fields + circular cropped image
    2nd Page  -> Update calendar placeholders + set background
    3rd Page  -> Replace {{Layout1}} + add background + append extra images
    """
    print(f"🔄 Creating brochure PPT from {template_path}")
    prs = Presentation(template_path)

    # Track temp files for cleanup
    temp_files = []

    # Normalize images (URL → local)
    circle_img = get_local_image(circle_img, "circle.png")
    if circle_img: temp_files.append(circle_img)

    calendar_bg = get_local_image(calendar_bg, "calendar_bg.png")
    if calendar_bg: temp_files.append(calendar_bg)

    layout_img = get_local_image(layout_img, "layout_img.png")
    if layout_img: temp_files.append(layout_img)

    layout_bg = get_local_image(layout_bg, "layout_bg.png")
    if layout_bg: temp_files.append(layout_bg)

    # ✅ Collect ALL category images (keep order)
    all_images = (
        categorized_images.get("existing_pictures", []) +
        categorized_images.get("floor_plans", []) +
        categorized_images.get("elevation_drawings", [])
    )
    # Normalize
    extra_images = []
    for i, img in enumerate(all_images):
        local = get_local_image(img, f"extra_{i}.png")
        if local:
            extra_images.append(local)
            temp_files.append(local)

    # -------------------------
    # 1️⃣ First Page
    # -------------------------
    if len(prs.slides) > 0:
        print("🖼️ Updating first page with text + circle image...")
        slide = prs.slides[0]
        text_map = {
            "Project Name": form_data.get("Project Name", ""),
            "Project Type": form_data.get("What is the nature of your project?", ""),
            "space To be Designed": form_data.get("Space(S) to be designed", ""),
            "Room size": form_data.get("What is the area size?", ""),
            "style(s) selected": form_data.get("Which style(s) do you like?", ""),
            "Location": f"{form_data.get('City', '')}, {form_data.get('Country', '')}",
        }
        text_map = {k: v for k, v in text_map.items() if v}
        replace_text_in_ppt(slide, text_map)
        if circle_img:
            replace_with_circle_image(slide, circle_img)

    # -------------------------
    # 2️⃣ Second Page
    # -------------------------
    if len(prs.slides) > 1:
        print("📅 Updating second page with calendar + background...")
        slide = prs.slides[1]
        update_calendar_with_bg(prs, slide, calendar_bg)

    # -------------------------
    # 3️⃣ Third Page
    # -------------------------
    if len(prs.slides) > 2:
        print("📐 Updating third page with layout + extra images...")
        slide = prs.slides[2]
        replace_layout_and_append_images(prs, slide, layout_img, layout_bg, extra_images)

    # -------------------------
    # Save final brochure
    # -------------------------
    prs.save(output_path)
    print(f"✅ Brochure PPT created: {output_path}")

    # -------------------------
    # 🧹 Cleanup temporary files
    # -------------------------
    for f in temp_files:
        try:
            if os.path.exists(f):
                os.remove(f)
                print(f"🗑️ Deleted temp file: {f}")
        except Exception as e:
            print(f"⚠️ Could not delete {f}: {e}")

    return output_path
