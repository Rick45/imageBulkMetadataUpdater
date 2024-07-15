
## Warning

**Important:** Running this script may result in the modification or loss of data. It is highly recommended to perform a backup of your data before executing the script. The author of this script is not responsible for any data loss or damage that may occur during the execution of this script.

Please ensure that you have a reliable backup of your data before proceeding with the script.


# Image Metadata Extractor

This Python script, `metadataExtractor.py`, is designed to extract metadata from image files, with a focus on converting GPS coordinates to decimal degrees. It provides the option to extract only recommended metadata based on a predefined pattern or to extract all available metadata.

## Features

- Extracts metadata from image files, including EXIF data.
- Converts GPS coordinates (`GPSLatitude` and `GPSLongitude`) from degrees, minutes, and seconds (DMS) to decimal degrees.
- Offers the option to extract only recommended metadata based on a pattern for decimal values.
- Skips unnecessary metadata like thumbnail data, `ComponentsConfiguration`, and `SceneType` to streamline the output.

## Requirements

- Python 3.x
- `piexif`: For handling EXIF data in image files.
- `re`: For regular expression operations, used in filtering recommended metadata.

## Installation

Make sure Python 3.x is installed on your system, then install the required Python package using pip:

```bash
pip install piexif
```

## Usage

Run the script in a terminal or command prompt:
```bash
python metadataExtractor.py
```
1. When prompted, enter the folder path containing the photos from which you want to extract metadata.
2. Choose whether to extract only recommended metadata by entering `y` (yes) or `n` (no) when prompted.


# Image Metadata Bulk Inserter

This Python script, `metadatabulkInserter.py`, automates the process of inserting metadata into image files. It reads the metadata from an Excel file and updates the images' EXIF data accordingly.

## Features

- Reads and Converts GPS coordinates from decimal degrees to degrees, minutes, and seconds (DMS), the format required by EXIF standards.
- Updates all the EXIF data present in the excel columns in the specified image files.

## Requirements

- Python 3.x
- `pandas`: For reading and manipulating data from Excel files.
- `piexif`: For handling EXIF data in image files.
- `re`: For regular expression operations.
- `ast`: For converting string representations of Python data structures into actual Python objects.

## Installation

Make sure Python 3.x is installed on your system, then install the required Python packages using pip:

```bash
pip install pandas piexif
```

## Usage

1. **Prepare an Excel File:** Use the metadataExtractor to generate an Excel file containing metadata. Update the desired metadata. if there arent any GPS coordinates you will need to add them manually by including columns for `GPSLatitude` and `GPSLongitude` alongside the filenames of the images.
2. **Place the Excel File:** Move the Excel file to a known directory. As an example, the script includes a commented path: `D:\DEV\imageBulkMetadataUpdater\samples\images_metadata.xlsx`.
3. **Run the Script:** Execute the script in your terminal or command prompt by running:

```bash
python metadatabulkInserter.py
```

1. When prompted, enter the folder path containing the excel file with the new metadata and the photos name.
2. When prompted, enter the folder path containing the photos from which you want to update the metadata.

