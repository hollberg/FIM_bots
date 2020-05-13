"""
build_dc_zips.py

Build *.zip files and associated 1003.txt file for
Donor Central fund statement uploads.
By: Mitch Hollberg (mhollberg@gmail.com)
Created: 4/30/2020
Last Update: 4/30/2020
"""

# Imports
import zipfile
import os
from os.path import basename

# Configuration Values
source_dir = r'P:\00_Common\FinanceCommon\DonorStatements\2020\2020Q1\2020Q1_Manual'
destination_dir = r'N:\found\eAdvisor\Statements\DonorCentral\2020Q1_V2'
zipfile_prefix = r'2020Q1_DC_Zip_'
begin_date = '1/1/2020'
end_date = '3/31/2020'

fname = 'contents1003.txt'
contents = ''

max_size = 15_000_000   # Max size of all files in zip
running_size = 0
zipfile_number = 1


def make_contents_line(fundid, fname, begin_date, end_date):
    """ Create a new line of the 'contents1003.txt' file """
    values = ['0142', fundid, fname, begin_date, end_date, '1003']
    return '\t'.join(values) + '\n'


# print(make_contents_line('fund', 'foo.txt', '1/1/2020', '3/31/2020'))

# Loop over files in 'source_dir'; Make zip file with 'contents1003.txt' file
number_of_files = len(os.listdir(source_dir))
for i, file in enumerate(os.listdir(source_dir)):
    # Define output zip file
    output_zip_filepath = os.path.join(destination_dir,
                                       zipfile_prefix+str(zipfile_number)+'.zip')
    with zipfile.ZipFile(output_zip_filepath, 'a') as output_zip:

        # Get next statement.pdf
        filepath = os.path.join(source_dir, file)   # Format of (C:\mytext.txt)
        fundid = file.split('_')[0]    # fname of format 'FUNDID_BEGINDATE.pdf'
        output_zip.write(filepath, basename(file))
        contents += make_contents_line(fundid=fundid, fname=file,
                                       begin_date=begin_date, end_date=end_date)

        file_stats = os.stat(filepath)
        running_size += file_stats.st_size      # In bytes; 1,000,000 = 1MB
        if (running_size >= max_size) \
                or (i + 1 == number_of_files):  # File's big enough/last loop
            with open(fname, 'w') as contents_file:
                contents_file.write(contents)
            # Write file (without full filepath) to zip
            # See stackoverflow.com/questions/16091904
            output_zip.write(fname)

            # Reset/Update values
            zipfile_number += 1
            running_size = 0
            contents = ''

