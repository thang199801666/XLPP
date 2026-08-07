"""Example: add an image to a workbook and save it."""

import os
import sys
from pathlib import Path

sys.path.insert(0, r"D:\Temp")
os.add_dll_directory(r"D:\Temp")

import xlpp


def test_add_image_and_save():
    image_path = Path(__file__).with_name("image1.png")
    output_path = Path(__file__).with_name("image_result.xlsx")

    workbook = xlpp.Workbook()
    worksheet = workbook.add_worksheet("Images")
    image = xlpp.Image.from_file(str(image_path), "D2")
    worksheet.add_image(image)

    workbook.save(str(output_path))
    assert output_path.exists()


if __name__ == "__main__":
    test_add_image_and_save()
    print("Saved:", Path(__file__).with_name("image_result.xlsx"))
