from rembg import remove
from PIL import Image
import io

# Use raw string for Windows paths
input_path = r"D:\TLDP_Project\image\Projet DUFOURCQ- 2nd Floor.jpg"
output_path = "output.png"

# Open image
img = Image.open(input_path)

# Remove background (returns bytes)
output = remove(img)

# Convert bytes to PIL Image
output_img = Image.open(io.BytesIO(output)).convert("RGBA")

# Save result
output_img.save(output_path)

print("✅ Background removed and saved as", output_path)
