import io
from rembg import remove
from PIL import Image

input_path = r"D:\TLDP_Project\image\Projet DUFOURCQ- 2nd Floor.jpg"
output_path = "output.png"

with open(input_path, "rb") as inp_file:
    input_bytes = inp_file.read()

# Remove background (returns bytes because input is bytes)
output_bytes = remove(input_bytes)

# Convert back to Image
output_img = Image.open(io.BytesIO(output_bytes)).convert("RGBA")

# Save result
output_img.save(output_path, format="PNG")

print("✅ Background removed and saved as", output_path)

