import os
import sys
import shutil
import logging
import json
import datetime as dt
import pandas as pd
import time
# import uuid
# import argparse
import re
import json


# from dusit.utils import *

# logging.basicConfig(format=LOG_FORMAT, stream=sys.stdout, level=logging.DEBUG)
# logging.basicConfig(format=LOG_FORMAT, filename='..\log\ideas_{}'.format(dt.datetime.today().strftime('%Y%m%d%H%M%S_%f')),level=logging.INFO)
def _change_filename(old_filename, property_code, load_date):
    """ Replace filename to proper filename

    Args:
        old_filename (str)
        property_code (str): eg. DTHH, DTPA, All, ...
        load_date (str): yyyyMMdd
    Returns:
        new_filename (str)
    """
    new_filename=old_filename
    if property_code == 'All':
        if 'Cmp_Daily' in old_filename:
            new_filename='{}_Cmp_Daily_All_{}.xlsx'.format(property_code, load_date)
        elif 'Cmp_Monthly' in old_filename:
            new_filename='{}_Cmp_Monthly_All_{}.xlsx'.format(property_code, load_date)
    else:
        if 'Cmp_Daily' in old_filename:
            new_filename='{}_Cmp_Daily_{}.xlsx'.format(property_code, load_date)
        elif 'Cmp_Monthly' in old_filename:
            new_filename='{}_Cmp_Monthly_{}.xlsx'.format(property_code, load_date)
        elif ('DailydSTAR' in old_filename) or ('MonthlydSTAR' in old_filename):
            new_filename='{}_{}_{}.xlsx'.format(
                                    property_code,
                                    old_filename.split('_')[0] + 'survey',
                                    load_date
                                    )
    return new_filename

def convert_xls_to_xlsx(path, property_code, load_date):
    files=[]
    for filename in filter(lambda x: x.endswith('xls') or x.endswith('xlsx'), os.listdir(path)):
        infile = os.path.join(path, filename)
        print(infile)
        outfile = os.path.join('\\\\stdiazddpblobteam.file.core.windows.net\\dev\\dev\\source\\Ideas\\Logic Apps\\data_out', filename)
        # logger.info('Converting {} to {}'.format(infile, outfile))
        xls = pd.ExcelFile(infile)
        writer = pd.ExcelWriter(outfile, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
        cell_format = writer.book.add_format({'num_format': '@'})
        for sheet in xls.sheet_names:
            df = xls.parse(sheet, header=None)
            df.to_excel(writer, sheet_name=sheet, header=False, index=False)
            worksheet = writer.sheets[sheet]
            worksheet.set_column('A:AAA', 12, cell_format)
        writer.save()
        files.append(outfile)
    return files

def move_to_download_folder(to_path, extension, files):
    for file in files:
        filename = file.split('\\')[-1]
        prop_code = filename[0:4]
        if not os.path.exists(os.path.join(to_path, prop_code)):
            os.makedirs(os.path.join(to_path, prop_code))

        check_size = os.path.getsize(file)
        shutil.move(file, os.path.join(to_path, prop_code, filename))
        while (check_size != os.path.getsize(os.path.join(to_path, prop_code, filename))):
            time.sleep(1)

if __name__ == '__main__':
    print('test')
    copy_path = 'tset'
    staging_path = "\\\\stdiazddpblobteam.file.core.windows.net\\dev\\dev\\source\\Ideas\\Logic Apps\\data_staging"
    copy_path = '\\\\stdiazddpblobteam.file.core.windows.net\\dev\\dev\\source\\Ideas\\Logic Apps\\data_out'
	# timeout = config['action_timeout']
    load_date = dt.date.today().strftime('%Y%m%d') # current date
    
    files = convert_xls_to_xlsx(staging_path, '', load_date)
    print('convert success')
    print(files)
    move_to_download_folder(copy_path, '', files)
