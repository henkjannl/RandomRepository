from pptx import Presentation
from pptx.util import Cm
from pathlib import Path

# ================= USER SETTINGS =================
IMAGE_DIR = Path(r"C:\Users\henkj\OneDrive\03 HenkJan\_ASMPT\Collet\images")
OUTPUT_PPTX = r"C:\Users\henkj\OneDrive\03 HenkJan\_ASMPT\Collet\layers.pptx"     # <-- output PowerPoint filename
# ================================================

prs = Presentation()

# 16:9 slide size
prs.slide_width  = Cm(25.4)
prs.slide_height = Cm(14.288)

# Use a blank slide layout
title_only_layout = prs.slide_layouts[5]

# Collect PNG files, sorted alphabetically
images = sorted(IMAGE_DIR.glob("*.png"))


for idx, img_path in enumerate(images, start=1):
    slide = prs.slides.add_slide(title_only_layout)

    # ---- Title ----
    # title_box = slide.shapes.add_textbox(
    #     Cm(1),    # left
    #     Cm(0.5),  # top
    #     Cm(20),   # width
    #     Cm(1.5)   # height
    # )
    # title_tf = title_box.text_frame
    # title_tf.clear()
    # title_tf.text = f"Layer {idx:02d}"
    slide.shapes.title.text = f"Layer {idx:02d}"

    # ---- Image ----
    slide.shapes.add_picture(
        str(img_path),
        left=Cm(1),        # 1 cm from left margin
        top=Cm(2.2),       # 2.2 cm from top margin
        height=Cm(11)      # fixed height, width auto by aspect ratio
    )

prs.save(OUTPUT_PPTX)
print(f"Saved presentation with {len(images)} slides to '{OUTPUT_PPTX}'")

