from PIL import Image, ImageDraw, ImageFont
import os


def create_centered_pe_jpg():
    # 1. Setup Canvas
    W, H = 600, 600
    img = Image.new('RGB', (W, H), color="#003366")  # Navy Blue
    draw = ImageDraw.Draw(img)

    # 2. Draw Border
    draw.rectangle([0, 0, W - 1, H - 1], outline="white", width=25)

    # 3. Draw Centered Text
    try:
        # Load Font
        font = ImageFont.truetype("arial.ttf", 360)

        # KEY FIX: anchor="mm" aligns the text's "Middle-Middle" to the coordinates provided
        # We set the coordinates to W/2 (300) and H/2 (300) - the exact center.
        # Note: 'pb' adjusts for the font baseline to make it vertically optical centered
        draw.text((W / 2, H / 2), "PE", fill="white", font=font, anchor="mm")

    except Exception as e:
        print(f"Font error: {e}. Using fallback shapes.")
        # Fallback Manual Shapes (Calculated to be centered)
        # Total width of PE approx 420px. Start X ~ 90.

        # P
        draw.rectangle([90, 120, 140, 480], fill="white")  # Vertical
        draw.rectangle([90, 120, 270, 170], fill="white")  # Top
        draw.rectangle([220, 120, 270, 300], fill="white")  # Right Loop
        draw.rectangle([90, 250, 270, 300], fill="white")  # Bottom Loop

        # E
        draw.rectangle([330, 120, 380, 480], fill="white")  # Vertical
        draw.rectangle([330, 120, 510, 170], fill="white")  # Top
        draw.rectangle([330, 275, 470, 325], fill="white")  # Middle
        draw.rectangle([330, 430, 510, 480], fill="white")  # Bottom

    # 4. Save
    output_path = "PE_Icon_600_Centered.jpg"
    img.save(output_path, "JPEG", quality=100)
    print(f"Success! Centered image saved to: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    create_centered_pe_jpg()