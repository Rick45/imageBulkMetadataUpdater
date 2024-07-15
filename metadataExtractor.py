import os
import pandas as pd
import piexif
import re

patternForDecimals = r"\(\s*\d+\s*,\s*\d+\s*\)"

grabOnlyRecommended = 'y'

def sanitize_for_excel(value):
	"""
	Sanitize a string for Excel.
	Removes or replaces characters that are illegal in Excel cells.
	"""
	if isinstance(value, str):
		# Remove non-printable characters
		value = ''.join(char for char in value if char.isprintable())
		# Replace other potentially problematic characters
		value = re.sub(r'[^\x20-\x7E]+', '_', value)  # Replace non-ASCII characters with '_'
		value = re.sub(r'[:*?\\/\[\]]+', '_', value)  # Replace illegal Excel characters with '_'
	return value

def dms_to_decimal(degrees, minutes, seconds):
	"""
	Convert degrees, minutes, and seconds to decimal degrees.

	Parameters:
	- degrees (int or float): The degrees part of the DMS.
	- minutes (int or float): The minutes part of the DMS.
	- seconds (int or float): The seconds part of the DMS.

	Returns:
	- float: The decimal degrees.
	"""
	return degrees + (minutes / 60) + (seconds / 3600)


def get_image_metadata(image_path):
	exif_dict = piexif.load(image_path)
	metadata = {}
	for ifd_name in exif_dict:
		if ifd_name == "thumbnail":
			continue  # Skip the thumbnail data
		for tag, value in exif_dict[ifd_name].items():
			tag_name = piexif.TAGS[ifd_name][tag]["name"]
			
			if tag_name == 'GPSLatitude' or tag_name == 'GPSLongitude':
				# Convert GPSLatitude to decimal degrees
				degrees, minutes, seconds = value
				metadata[tag_name] = dms_to_decimal(degrees[0], minutes[0], seconds[0])

			match = False
			if grabOnlyRecommended == 'y':
				match = re.search(patternForDecimals, str(value))
			if tag_name == 'ComponentsConfiguration' or tag_name == 'SceneType' or match:
				continue
			if tag_name not in ('MakerNote'):
				# Convert bytes to string if necessary ComponentsConfiguration
				if isinstance(value, bytes):
					try:
						value = value.decode()
					except UnicodeDecodeError:
						value = "<binary data>"
				metadata[tag_name] = value
	return metadata

def main():
	global grabOnlyRecommended
	folder_path = input("Please enter the folder path with the photos: ")
	grabOnlyRecommended = input("Do you want to grab only recommended metadata? (y/n): ")
	if grabOnlyRecommended == '':
		grabOnlyRecommended = 'y'
	image_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]

	all_metadata = []
	for image_file in image_files:
		full_path = os.path.join(folder_path, image_file)
		print(f"Processing {image_file}...")  # Print the current image being processed
		metadataDirty = get_image_metadata(full_path)
		# Exclude 'ExifVersion' from the metadata
		metadata = {k: sanitize_for_excel(v) for k, v in metadataDirty.items() if k != 'ExifVersion'}
		metadata['Filename'] = image_file  # Add filename to metadata
		all_metadata.append(metadata)

	# Create DataFrame
	df = pd.DataFrame(all_metadata)

	# Ensure 'Filename' is the first column
	mandatory_cols = ['Filename']
	optional_cols = ['GPSLatitude', 'GPSLongitude']
	existing_optional_cols = [col for col in optional_cols if col in df.columns]
	cols = mandatory_cols + existing_optional_cols + [col for col in df.columns if col not in mandatory_cols + optional_cols]
	df = df[cols]

	excel_filename = os.path.join(folder_path, 'images_metadata.xlsx')
	df.to_excel(excel_filename, index=False)
	print(f"Excel file '{excel_filename}' has been created with the images' metadata.")

if __name__ == "__main__":
	main()