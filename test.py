from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
import requests
from io import BytesIO

prs = Presentation("template.pptx")

img_url = "https://picsum.photos/400/400"  # your image
img_data = BytesIO(requests.get(img_url).content)

for slide in prs.slides:
    for shape in slide.shapes:
        if shape.name == "CircleImage":  # give your circle shape this name in PPT
            fill = shape.fill
            fill.solid()
            fill.user_picture(img_data)

prs.save("output.pptx")
