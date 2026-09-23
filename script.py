import os
import shutil

import pandas as pd
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(BASE_DIR, "Report")
IMAGES_DIR = os.path.join(BASE_DIR, "RiteshHandwritten")
SOURCE_REPORT = os.path.join(REPORT_DIR, "ocr_accuracy_report.csv")
OUTPUT_REPORT = os.path.join(REPORT_DIR, "ocr_accuracy_report_with_images.xlsx")

ROW_HEIGHT = 90
IMG_WIDTH = 120
IMG_HEIGHT = 90


def build_report_with_images(source_csv=SOURCE_REPORT, images_dir=IMAGES_DIR,
                              output_path=OUTPUT_REPORT):
    # keep an untouched copy of the source report alongside the new one
    copy_path = os.path.join(REPORT_DIR, "ocr_accuracy_report_copy.csv")
    shutil.copyfile(source_csv, copy_path)

    df = pd.read_csv(source_csv)
    df.insert(1, "image", "")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="OCR Accuracy")
        sheet = writer.sheets["OCR Accuracy"]

        image_col_letter = get_column_letter(df.columns.get_loc("image") + 1)
        sheet.column_dimensions[image_col_letter].width = 18

        for row_idx, image_name in enumerate(df["image_name"], start=2):
            image_path = os.path.join(images_dir, str(image_name))
            if not os.path.isfile(image_path):
                print(f"Missing image for {image_name}, skipping")
                continue

            with PILImage.open(image_path) as im:
                width, height = im.size
            scale = min(IMG_WIDTH / width, IMG_HEIGHT / height)

            img = XLImage(image_path)
            img.width = width * scale
            img.height = height * scale
            sheet.add_image(img, f"{image_col_letter}{row_idx}")
            sheet.row_dimensions[row_idx].height = ROW_HEIGHT

    print(f"Wrote report with images to {output_path}")


def rename_files(folder):
    entries = sorted(
        f for f in os.listdir(folder)
        if os.path.isfile(os.path.join(folder, f))
    )

    # two-pass rename avoids collisions when new names overlap old ones
    temp_names = []
    for i, filename in enumerate(entries, start=1):
        ext = os.path.splitext(filename)[1]
        src = os.path.join(folder, filename)
        tmp = os.path.join(folder, f"__tmp_{i}{ext}")
        os.rename(src, tmp)
        temp_names.append((tmp, ext))

    for i, (tmp, ext) in enumerate(temp_names, start=1):
        dst = os.path.join(folder, f"Test_img{i}{ext}")
        os.rename(tmp, dst)
        print(f"{tmp} -> {dst}")


if __name__ == "__main__":
    build_report_with_images()
