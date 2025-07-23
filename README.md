This project is a comprehensive AI-powered Business Card Scanner and Contact Extractor, developed using Python, OpenCV, and EasyOCR, integrated with Google Custom Search API for intelligent company name correction. The tool automates the process of extracting structured information from scanned business cards and organizes it into a clean, formatted Excel spreadsheet.

The workflow begins with advanced image preprocessing techniques using OpenCV, including grayscale conversion, adaptive thresholding, morphological operations for line removal, noise cleaning, image sharpening, and resolution upscaling. These steps significantly enhance OCR accuracy by preparing the card images for robust text recognition.

Text extraction is powered by EasyOCR, which reads the processed images and extracts all visible text. The script then employs sophisticated regex patterns and natural language heuristics to parse out meaningful information such as:

Company Name

Owner Name(s)

Email Address

Phone Numbers

Physical Address

For ambiguous or incomplete company names, the system uses the Google Custom Search API to intelligently infer the most likely official company name by cross-referencing the extracted data (like owner names, email domains, or address) with live search results.

The final structured data is then exported into an Excel workbook, with styling and formatting handled via OpenPyXL, including colored headers, column auto-sizing, and consistent data alignment.

Additionally, the script is modular, reusable, and maintains clean logging to help track every step of the OCR, parsing, and correction pipeline. It supports batch processing for multiple images and stores intermediate and final outputs for debugging and review.

This project showcases expertise in computer vision, OCR, data cleaning, regular expressions, Excel automation, and API integration, making it a powerful real-world tool for digitizing contact information from physical media.
