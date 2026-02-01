from html2image import Html2Image
import os

# Set up paths
html_file = r"G:\My Drive\School\IIC\infographic_IIH.html"
output_dir = r"G:\My Drive\School\IIC"
output_file = "infographic_IIH.png"

# Read HTML content
with open(html_file, 'r', encoding='utf-8') as f:
    html_content = f.read()

# Create Html2Image instance
hti = Html2Image(output_path=output_dir)

# Set browser size for the infographic
hti.screenshot(
    html_str=html_content,
    save_as=output_file,
    size=(1200, 2200)
)

print(f"Infographic saved to: {os.path.join(output_dir, output_file)}")
