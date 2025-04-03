import originpro as op
import pandas as pd
import os
import sys
import re

# Ensures that Origin gets shut down if an uncaught exception
def origin_shutdown_exception_hook(exctype, value, traceback):
    '''Ensures Origin gets shut down if an uncaught exception'''
    op.exit()
    sys.__excepthook__(exctype, value, traceback)
if op and op.oext:
    sys.excepthook = origin_shutdown_exception_hook

# Set Origin instance visibility.
if op.oext:
    op.set_show(True)

# Directory containing .opju files
src_dir = r'src_folder'

# Base output directory for the exported CSV files
base_output_dir = r'output_folder'

# Sanitize filenames for Windows compatibility (remove special characters)
def sanitize_filename(filename):
    return re.sub(r'[<>:"/\/|?*]', '_', filename)

# Loop through all .opju files in the source directory
for filename in os.listdir(src_dir):
    if filename.endswith('.opju'):
        opju_path = os.path.join(src_dir, filename)
        
        # Open the .opju file
        op.open(file=opju_path, readonly=True)
        
        # Create a sanitized folder name based on the opju file name (without extension)
        sanitized_folder_name = sanitize_filename(filename.replace('.opju', ''))
        output_dir = os.path.join(base_output_dir, sanitized_folder_name)
        
        # Create the directory for this opju file, if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Loop through all worksheets in the project
        for wb in op.pages('w'):
            for wks in wb:
                # Convert the worksheet to a pandas DataFrame
                df = wks.to_df()
                
                # Set the output CSV filename
                csv_filename = os.path.join(output_dir, f'{wks.name}.csv')
                
                # Export the DataFrame to a CSV file
                df.to_csv(csv_filename, index=False)
                print(f"Exported {wks.name} to {csv_filename}")
        
        # Exit running instance of Origin for this file
        if op.oext:
            op.exit()
