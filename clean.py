import cv2
from PIL import Image

def remove_color_keep_lines(input_path, output_path):
    # Read image
    img = cv2.imread(input_path)

    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Invert image (optional, makes lines darker)
    inverted = cv2.bitwise_not(gray)

    # Apply adaptive threshold (extract lines)
    thresh = cv2.adaptiveThreshold(
        inverted, 255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY,
        15, -2
    )

    # Invert back to black lines on white background
    cleaned = cv2.bitwise_not(thresh)

    # Save result
    cv2.imwrite(output_path, cleaned)
    print(f"✅ Saved cleaned line drawing to {output_path}")

# Example usage
remove_color_keep_lines("D:\TLDP_Project\image\Projet DUFOURCQ - 1st FLOOR WIP.jpg", "floorplan_clean.png")
