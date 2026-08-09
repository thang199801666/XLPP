"""Test adding an image to a workbook and saving it."""

from pathlib import Path
import xlpp


def test_add_image_and_save(tmp_path: Path):
    image_path = Path(__file__).with_name("image1.png")
    output_path = tmp_path / "image_result.xlsx"

    workbook = xlpp.Workbook()
    worksheet = workbook.add_worksheet("Images")
    image = xlpp.Image.from_file(str(image_path), "D2")
    worksheet.add_image(image)

    workbook.save(str(output_path))
    assert output_path.exists()

    loaded = xlpp.Workbook()
    loaded.load(str(output_path))
    assert loaded["Images"].image_count == 1
