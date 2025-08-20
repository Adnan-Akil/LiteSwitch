import sys
import os

from converter.document_converter import CONVERSION_MAP

def main():
  if len(sys.argv) < 2:
    print("No file path provided")
    sys.exit(1)

  input_path = os.path.abspath(sys.argv[1].strip('"'))
  file_dir = os.path.dirname(input_path)
  os.chdir(file_dir)

  if not os.path.exists(input_path):
    print(f"File not found: {input_path}")
    sys.exit(1)
  
  file_extension = os.path.splitext(input_path)[1].lower()
  source_ext=file_extension.lstrip('.')
  target_ext="pdf"

  if source_ext in CONVERSION_MAP and target_ext in CONVERSION_MAP[source_ext]:
    try:
      converter = CONVERSION_MAP[source_ext][target_ext]
      output_path = converter(input_path)
      print(f"Conversion Successful {output_path}")
    except Exception as e:
      print(f"Error during conversion: {e}")
      sys.exit(1)
  else:
    print(f"No conversion available from {source_ext} to {target_ext}")

if __name__ == "__main__":
  main()