# brochure.py

from pptx import Presentation
from pptx.util import Inches
from datetime import datetime, timedelta
from PIL import Image, ImageDraw
import platform

# -------------------------
# Constants
# -------------------------
if platform.system() == "Windows":
    DAY_FORMAT = "%#d-%b"
else:
    DAY_FORMAT = "%-d-%b"


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


def replace_with_circle_image(prs, img_path):
    cropped_path = make_circle_image(img_path)
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and "{{image1}}" in shape.text:
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture(cropped_path, left, top, width, height)
    return prs


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


def update_calendar_with_bg(prs, image_path, start_date=None):
    mapping = build_mapping(start_date)
    slide_width, slide_height = prs.slide_width, prs.slide_height
    for slide in prs.slides:
        # Insert full background image
        pic = slide.shapes.add_picture(image_path, 0, 0, width=slide_width, height=slide_height)
        slide.shapes._spTree.remove(pic._element)
        slide.shapes._spTree.insert(2, pic._element)
        # Replace text
        for shp in iter_shapes(slide.shapes):
            if getattr(shp, "has_table", False):
                for row in shp.table.rows:
                    for cell in row.cells:
                        replace_text_in_frame(cell.text_frame, mapping)
            elif getattr(shp, "has_text_frame", False):
                replace_text_in_frame(shp.text_frame, mapping)
    return prs


# -------------------------
# Layout + extra images
# -------------------------
def replace_layout_and_append_images(prs, layout_img_path, bg_img_path, extra_images):
    slide_width, slide_height = prs.slide_width, prs.slide_height
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and "{{Layout1}}" in shape.text:
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture(layout_img_path, left, top, width, height)
                # Add background
                pic = slide.shapes.add_picture(bg_img_path, 0, 0, slide_width, slide_height)
                slide.shapes._spTree.remove(pic._element)
                slide.shapes._spTree.insert(2, pic._element)

    # Append new slides with extra images
    blank_layout = prs.slide_layouts[6]  # blank layout
    for img in extra_images:
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(img, Inches(1), Inches(1), width=Inches(7), height=Inches(5))
    return prs


# -------------------------
# Text replacement helper
# -------------------------
def replace_text_in_ppt(prs, text_map):
    for slide in prs.slides:
        for shp in iter_shapes(slide.shapes):
            if getattr(shp, "has_text_frame", False):
                for para in shp.text_frame.paragraphs:
                    for run in para.runs:
                        for ph, val in text_map.items():
                            if ph in run.text:
                                run.text = run.text.replace(ph, val)
    return prs


# -------------------------
# Main: Create Brochure
# -------------------------
def create_brochure_ppt(template_path, output_path, form_data, circle_img, calendar_bg, layout_img, layout_bg, extra_images):
    """
    Create a brochure PPT with:
    1st Page  -> Replace text fields + circular cropped image
    2nd Page  -> Update calendar placeholders + set background
    3rd Page  -> Replace {{Layout1}} + add background + append extra images
    """
    print(f"🔄 Creating brochure PPT from {template_path}")
    prs = Presentation(template_path)

    # -------------------------
    # 1️⃣ First Page (Text + Circle Image)
    # -------------------------
    print("🖼️ Updating first page with text + circle image...")
    text_map = {
        "Project Name": form_data.get("Project Name", ""),
        "Project Type": form_data.get("What is the nature of your project?", ""),
        "space To be Designed": form_data.get("Space(S) to be designed", ""),
        "Room size": form_data.get("What is the area size?", ""),
        "style(s) selected": form_data.get("Which style(s) do you like?", ""),
        "Location": f"{form_data.get('City', '')}, {form_data.get('Country', '')}",
    }
    text_map = {k: v for k, v in text_map.items() if v}
    prs = replace_text_in_ppt(prs, text_map)
    prs = replace_with_circle_image(prs, circle_img)

    # -------------------------
    # 2️⃣ Second Page (Calendar + BG Image)
    # -------------------------
    print("📅 Updating second page with calendar + background...")
    prs = update_calendar_with_bg(prs, calendar_bg)

    # -------------------------
    # 3️⃣ Third Page (Layout + BG + Extra Images)
    # -------------------------
    print("📐 Updating third page with layout + extra images...")
    prs = replace_layout_and_append_images(prs, layout_img, layout_bg, extra_images)

    # -------------------------
    # Save final brochure
    # -------------------------
    prs.save(output_path)
    print(f"✅ Brochure PPT created: {output_path}")
    return output_path
