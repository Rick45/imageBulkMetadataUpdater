import os
import pandas as pd
import piexif
import re
import ast


#D:\DEV\imageBulkMetadataUpdater\samples\images_metadata.xlsx

patternForDecimals = r"\(\s*\d+\s*,\s*\d+\s*\)"

def decimal_degrees_to_dms(degrees):
	"""Convert decimal degrees to degrees, minutes, seconds (as required by EXIF)"""
	is_positive = degrees >= 0
	degrees = abs(degrees)
	d = int(degrees)
	m = int((degrees - d) * 60)
	s = (degrees - d - m/60) * 3600.0
	return d, m, s, is_positive


def apply_metadata_to_image(image_path, metadata):
	try:
		exif_dict = piexif.load(image_path)
		
		latitude = metadata.get('GPSLatitude')
		longitude = metadata.get('GPSLongitude')
		
		if latitude and longitude:
			# Convert latitude and longitude from decimal degrees to EXIF format
			lat_d, lat_m, lat_s, lat_positive = decimal_degrees_to_dms(latitude)
			lon_d, lon_m, lon_s, lon_positive = decimal_degrees_to_dms(longitude)

			# Format for EXIF
			exif_dict['GPS'][piexif.GPSIFD.GPSLatitude] = [(lat_d, 1), (lat_m, 1), (int(lat_s * 100), 100)]
			exif_dict['GPS'][piexif.GPSIFD.GPSLatitudeRef] = 'N' if lat_positive else 'S'
			exif_dict['GPS'][piexif.GPSIFD.GPSLongitude] = [(lon_d, 1), (lon_m, 1), (int(lon_s * 100), 100)]
			exif_dict['GPS'][piexif.GPSIFD.GPSLongitudeRef] = 'E' if lon_positive else 'W'
		
		# Apply other metadata as before
		for ifd in ("0th", "Exif", "GPS", "1st"):
			for tag in exif_dict[ifd]:
				tag_name = piexif.TAGS[ifd][tag]["name"]
				if tag_name in metadata and tag_name not in ["GPSLatitude", "GPSLatitudeRef","GPSLongitude", "GPSLongitudeRef"]:

					este = metadata[tag_name]
					match = re.search(patternForDecimals, str(metadata[tag_name]))

					if match:
						print(f"Skipping tag: {tag_name} for now, this tags are a mess to calculate")

						if tag_name =="ExposureTime":							
							print(f"Applying tag: {tag_name} with value: {metadata[tag_name]}")
							exif_dict[ifd][tag] = metadata[tag_name]

						#avoid this tags, they are a mess to calculate as they return a tuple and require to be inserted as multiple types depending on the tag
						if "GPS" not in tag_name:
							# Proceed with operations if it is a tuple
							# For example, converting value_tuple to an EXIF-compatible format
							numeric_tuple = ast.literal_eval(str(metadata[tag_name]))
							# Unpack the tuple into two variables
							value1, value2 = numeric_tuple
							finalValue = value1 / value2
							print(f"Applying tag: {tag_name} with value: {finalValue}")
							exif_dict[ifd][tag] = finalValue
					else:
						# Handle the case where value_tuple is not a tuple
						#print("value_tuple is not a tuple.")
						if metadata[tag_name] is not None:
							print(f"Applying tag: {tag_name} with value: {metadata[tag_name]}")
							exif_dict[ifd][tag] = metadata[tag_name]
						else:
							print(f"Skipping tag {tag_name} because is empty")
							#exif_dict[ifd][tag] = ""
		
		exif_bytes = piexif.dump(exif_dict)
		piexif.insert(exif_bytes, image_path)
		print(f"Metadata applied to {image_path}")
	except Exception as e:
		print(f"Error applying metadata to {image_path}: {e}")

def main():
	excel_path = input("Please enter the path to the Excel file: ")
	photos_folder = input("Please enter the path to the photos folder: ")
	
	# Read the Excel file
	df = pd.read_excel(excel_path)
	
	# Iterate through the DataFrame
	for index, row in df.iterrows():
		filename = row['Filename']
		image_path = os.path.join(photos_folder, filename)
		
		# Prepare metadata, excluding the 'Filename' column
		metadata = {col: row[col] for col in df.columns if col != 'Filename' and row[col] != ''}
		
		# Apply metadata to image
		if os.path.exists(image_path):
			print(f"Processing {image_path}...")  # Print the current image being processed
			apply_metadata_to_image(image_path, metadata)
		else:
			print(f"File not found: {image_path}")

if __name__ == "__main__":
	main()