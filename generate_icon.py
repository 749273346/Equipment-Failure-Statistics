from PIL import Image, ImageDraw

def create_app_icon(output_path="app_icon.ico"):
    # Sizes for the icon
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = []

    for size in sizes:
        width, height = size
        # Create a new image with a dark blue background
        img = Image.new('RGBA', size, color=(44, 62, 80, 255)) # #2C3E50
        draw = ImageDraw.Draw(img)

        # Draw a border (lighter blue)
        border_width = max(1, width // 32)
        draw.rectangle(
            [0, 0, width - 1, height - 1], 
            outline=(52, 152, 219), # #3498DB
            width=border_width
        )

        # Draw 3 bars (Statistics)
        # Margins
        margin_x = width // 5
        margin_y = height // 5
        bar_width = (width - 2 * margin_x) // 3 - (width // 20)
        
        # Bar 1 (Red)
        b1_h = height // 3
        x1 = margin_x
        y1 = height - margin_y - b1_h
        draw.rectangle([x1, y1, x1 + bar_width, height - margin_y], fill=(231, 76, 60)) # #E74C3C

        # Bar 2 (Yellow)
        b2_h = height // 2
        x2 = x1 + bar_width + (width // 20)
        y2 = height - margin_y - b2_h
        draw.rectangle([x2, y2, x2 + bar_width, height - margin_y], fill=(241, 196, 15)) # #F1C40F

        # Bar 3 (Green)
        b3_h = int(height * 0.6)
        x3 = x2 + bar_width + (width // 20)
        y3 = height - margin_y - b3_h
        draw.rectangle([x3, y3, x3 + bar_width, height - margin_y], fill=(46, 204, 113)) # #2ECC71

        images.append(img)

    # Save as ICO
    images[0].save(
        output_path, 
        format='ICO', 
        sizes=sizes, 
        append_images=images[1:]
    )
    print(f"Icon saved to {output_path}")

if __name__ == "__main__":
    create_app_icon()
