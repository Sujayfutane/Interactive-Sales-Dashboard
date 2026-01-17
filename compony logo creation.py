from PIL import Image, ImageDraw, ImageFont

# Create a blank image
img = Image.new('RGB', (300, 100), color=(30, 144, 255))  # Dodger Blue background

# Draw text
draw = ImageDraw.Draw(img)
font = ImageFont.load_default()
draw.text((50, 35), "My Company", fill="white", font=font)

# Save the logo
img.save(r"G:\Data Science Intership\Interactive Sales Dashboard\company_logo.png")
print("Logo created successfully!")
