"""
Google Cloud Vision API - Bengali OCR Script
BTRC Meeting Portal - Image to Text Converter
=============================================
Usage:
  python ocr_bengali.py                          # Current folder images
  python ocr_bengali.py --input ./images         # Specific input folder
  python ocr_bengali.py --input img.jpg          # Single image file
"""

import os
import sys
import json
import argparse
import glob
from pathlib import Path
from datetime import datetime

# Service account key (same folder as this script)
KEY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "influential-kit-319418-d642233e873f.json")
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = KEY_FILE

SUPPORTED_FORMATS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp")


def get_vision_client():
    try:
        from google.cloud import vision
        return vision.ImageAnnotatorClient()
    except ImportError:
        print("google-cloud-vision not installed!")
        print("Run: pip install google-cloud-vision")
        sys.exit(1)


def ocr_image(client, image_path):
    from google.cloud import vision

    image_path = Path(image_path)
    print(f"  Processing: {image_path.name}")

    with open(image_path, "rb") as f:
        content = f.read()

    image = vision.Image(content=content)

    response = client.document_text_detection(
        image=image,
        image_context=vision.ImageContext(
            language_hints=["bn", "en"]
        )
    )

    result = {
        "file": str(image_path),
        "filename": image_path.name,
        "text": "",
        "words": [],
        "blocks": [],
        "error": None
    }

    if response.error.message:
        result["error"] = response.error.message
        print(f"    API Error: {response.error.message}")
        return result

    if response.full_text_annotation:
        result["text"] = response.full_text_annotation.text

        for page in response.full_text_annotation.pages:
            for block in page.blocks:
                block_text = ""
                for para in block.paragraphs:
                    for word in para.words:
                        word_text = "".join([s.text for s in word.symbols])
                        confidence = word.confidence if hasattr(word, 'confidence') else 0
                        result["words"].append({
                            "text": word_text,
                            "confidence": round(confidence * 100, 1)
                        })
                        block_text += word_text + " "
                result["blocks"].append(block_text.strip())

    word_count = len(result["text"].split()) if result["text"] else 0
    print(f"    Extracted {word_count} words")
    return result


def save_results(results, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    combined_lines = [
        "=" * 60,
        "BTRC OCR Results - Bengali Text Extraction",
        "Generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Total files: " + str(len(results)),
        "=" * 60,
        ""
    ]

    successful = 0
    for r in results:
        fname = Path(r["filename"]).stem

        # Individual text file per image
        out_file = output_dir / (fname + "_ocr.txt")
        with open(out_file, "w", encoding="utf-8") as f:
            f.write("File: " + r["filename"] + "\n")
            f.write("Date: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
            f.write("-" * 40 + "\n\n")
            if r["error"]:
                f.write("ERROR: " + r["error"] + "\n")
            else:
                f.write(r["text"] if r["text"] else "(No text detected)\n")

        combined_lines.append("FILE: " + r["filename"])
        combined_lines.append("-" * 40)
        if r["error"]:
            combined_lines.append("ERROR: " + r["error"])
        else:
            combined_lines.append(r["text"] if r["text"] else "(No text detected)")
            successful += 1
        combined_lines.append("")

    # Save combined file
    combined_file = output_dir / ("combined_ocr_" + timestamp + ".txt")
    with open(combined_file, "w", encoding="utf-8") as f:
        f.write("\n".join(combined_lines))

    # Save JSON data
    json_file = output_dir / ("ocr_data_" + timestamp + ".json")
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 50)
    print("OCR Complete!")
    print("  Processed  : " + str(len(results)) + " file(s)")
    print("  Successful : " + str(successful))
    print("  Failed     : " + str(len(results) - successful))
    print("\nOutput saved to: " + str(output_dir))
    print("  Combined text : " + combined_file.name)
    print("  JSON data     : " + json_file.name)
    return combined_file


def collect_images(input_path):
    p = Path(input_path)
    if p.is_file():
        if p.suffix.lower() in SUPPORTED_FORMATS:
            return [str(p)]
        else:
            print("Unsupported format: " + p.suffix)
            sys.exit(1)
    elif p.is_dir():
        images = []
        for ext in SUPPORTED_FORMATS:
            images.extend(glob.glob(str(p / ("**/*" + ext)), recursive=True))
            images.extend(glob.glob(str(p / ("**/*" + ext.upper())), recursive=True))
        return sorted(set(images))
    else:
        print("Path not found: " + str(input_path))
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="Bengali OCR using Google Cloud Vision API")
    parser.add_argument("--input", "-i", default=".", help="Input image file or folder")
    parser.add_argument("--output", "-o", default="ocr_output", help="Output folder")
    args = parser.parse_args()

    print("\nBTRC Bengali OCR Tool")
    print("=" * 50)

    if not os.path.exists(KEY_FILE):
        print("Service account key not found: " + KEY_FILE)
        sys.exit(1)
    print("Using credentials: " + Path(KEY_FILE).name)

    images = collect_images(args.input)
    if not images:
        print("No images found in: " + args.input)
        print("Supported: " + ", ".join(SUPPORTED_FORMATS))
        sys.exit(0)

    print("Found " + str(len(images)) + " image(s) to process\n")

    client = get_vision_client()

    results = []
    for i, img_path in enumerate(images, 1):
        print("[" + str(i) + "/" + str(len(images)) + "]", end=" ")
        result = ocr_image(client, img_path)
        results.append(result)

    output_base = Path(args.input) if Path(args.input).is_dir() else Path(".")
    save_results(results, str(output_base / args.output))


if __name__ == "__main__":
    main()
