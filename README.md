# Image Extractor for PowerPoint, Word and PDF

![powerpoint-pdf-image-extractor-github](https://github.com/user-attachments/assets/c177bb10-7da0-418d-8a74-15ea2fe3239f)


A CLI tool to extract images from PowerPoint, Word and PDF files written in Python 🐍. This script extract all images in your .pptx, .docx, or .pdf file into a local folder. The benefit of using this tool to extract images over taking screenshots is that you get the highest resolution possible.

## Use Cases

- 1️⃣ Extract images from PowerPoint presentations
- 2️⃣ Extract images from Word (doc/docx) documents
- 3️⃣ Extract images from PDF files

## Features

- ⬇️ Extract and download all images within a PowerPoint, Word or PDF
- 📁 Supports all image file types (jpg, png, jp2, gif, tiff, ...)
- 📑 Supports extracing images from: PowerPoint (.pptx, .ppt), Word (.docx, .doc) and PDF (.pdf)
- 📸 High resolution images: Images are not compressed
- 📀 Runs locally: Keep your data

## Setup

Create a virtual Python env
```
python3 -m venv env
```

Activate the virtual env
```
source env/bin/activate
```

Using [pip](https://pip.pypa.io/en/stable/installation/) install all dependencies
```
pip3 install -r requirements.txt
```

## Requirements

You need to have [7Zip](https://www.7-zip.org) installed because under the hood `unzip` is used to unarchive and archive the pptx files.

## Usage

```
python3 image_extractor.py <INPUT_FILE_PATH>
```

_⚠️ Note:_ All images of the PowerPoint, PDF or Word document will be extracted to a folder called `extracted_images` in the same folder as the original document.

## License 

Apache License 2.0: See `LICENSE` file

## Author

Written and maintained by [SlideSpeak.co](https://slidespeak.co)
