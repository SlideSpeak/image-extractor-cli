import logging
import os
import subprocess
from uuid import uuid4
from pypdf import PdfReader
from PIL import Image
import io
import sys


def generate_uuid_filename(extension):
    return f"{uuid4()}{extension.lower()}"


def ensure_directory_exists(path):
    if not os.path.exists(path):
        os.makedirs(path)


def extract_images(file_path: str):
    output_path = os.path.join(os.path.dirname(file_path), "extracted_images")
    _, file_extension = os.path.splitext(file_path)
    file_extension = file_extension.lower()

    if file_extension == ".pdf":
        extract_images_from_pdf(file_path, output_path)
    elif file_extension == ".docx":
        extract_images_from_docx(file_path, output_path)
    elif file_extension == ".pptx":
        extract_images_from_pptx(file_path, output_path)
    else:
        logging.error(f"Unsupported file type: {file_extension}")


def extract_images_from_pdf(pdf_file_path: str, output_path: str):
    try:
        reader = PdfReader(pdf_file_path)
        ensure_directory_exists(output_path)
        for page in reader.pages:
            for image in page.images:
                print("image")
                ext = os.path.splitext(image.name)[1].lower()
                image_data = image.data
                if ext == ".jpeg":
                    ext = ".jpg"
                elif ext == ".jp2":
                    try:
                        with Image.open(io.BytesIO(image_data)) as img:
                            if img.mode == "RGBA":
                                img = img.convert("RGB")
                            ext = ".png"
                            image_data = io.BytesIO()
                            img.save(image_data, format="PNG")
                            image_data = image_data.getvalue()
                    except Exception as e:
                        logging.error(f"Failed to convert JP2 to PNG: {e}")
                        continue

                image_filename = generate_uuid_filename(ext)
                file_path = os.path.join(output_path, image_filename)
                with open(file_path, "wb") as fp:
                    fp.write(image_data)
    except Exception as e:
        logging.error(f"Failed to extract images from {pdf_file_path}: {e}")


def extract_images_from_docx(docx_file_path: str, output_path: str):
    try:
        ensure_directory_exists(output_path)
        subprocess.run(
            ["unzip", "-j", docx_file_path, "word/media/*", "-d", output_path],
            check=True,
            shell=False,
        )
        for filename in os.listdir(output_path):
            original_path = os.path.join(output_path, filename)
            _, ext = os.path.splitext(filename)
            if ext.lower() == ".jpeg":
                ext = ".jpg"
            elif ext.lower() == ".jp2":
                try:
                    with Image.open(original_path) as img:
                        if img.mode == "RGBA":
                            img = img.convert("RGB")
                        ext = ".png"
                        img.save(original_path, format="PNG")
                except Exception as e:
                    logging.error(f"Failed to convert JP2 to PNG: {e}")
                    os.remove(original_path)
                    continue

            new_filename = generate_uuid_filename(ext)
            new_path = os.path.join(output_path, new_filename)
            os.rename(original_path, new_path)
    except Exception as e:
        logging.error(f"Failed to extract images from {docx_file_path}: {e}")


def extract_images_from_pptx(pptx_file_path: str, output_path: str):
    try:
        ensure_directory_exists(output_path)
        subprocess.run(
            ["unzip", "-j", pptx_file_path, "ppt/media/*", "-d", output_path],
            check=True,
            shell=False,
        )

        valid_extensions = {".jpg", ".jpeg", ".png", ".jp2"}
        for filename in os.listdir(output_path):
            original_path = os.path.join(output_path, filename)
            _, ext = os.path.splitext(filename)
            if ext.lower() == ".jpeg":
                ext = ".jpg"
            elif ext.lower() == ".jp2":
                try:
                    with Image.open(original_path) as img:
                        if img.mode == "RGBA":
                            img = img.convert("RGB")
                        ext = ".png"
                        img.save(original_path, format="PNG")
                except Exception as e:
                    logging.error(f"Failed to convert JP2 to PNG: {e}")
                    os.remove(original_path)
                    continue

            if ext.lower() in valid_extensions:
                new_filename = generate_uuid_filename(ext)
                new_path = os.path.join(output_path, new_filename)
                os.rename(original_path, new_path)
            else:
                os.remove(original_path)
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to unzip and extract images from {pptx_file_path}: {e}")
    except Exception as e:
        logging.error(f"An error occurred while processing {pptx_file_path}: {e}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python extract_images.py <file_path>")
        sys.exit(1)

    file_path = sys.argv[1]
    extract_images(file_path)
