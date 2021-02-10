import os
import shutil
import time
import logging
import pandas as pd

LOG_FORMAT = '%(asctime)s [%(levelname)s] %(name)s: %(message)s'

logger = logging.getLogger('DUSIT')

def convert_xlsx(source_path):
    files=[]
    if not os.path.exists(os.path.join(source_path, 'converted')):
        os.makedirs(os.path.join(source_path, 'converted'))

    for filename in filter(lambda x: x.endswith('xls') or x.endswith('xlsx'), os.listdir(source_path)):
        infile = os.path.join(source_path, filename)
        print(infile)
        outfile = os.path.join(source_path, 'converted', filename)
        logger.info('Converting: {} to {}'.format(infile, outfile))
        xls = pd.ExcelFile(infile)
        writer = pd.ExcelWriter(outfile, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
        cell_format = writer.book.add_format({'num_format': '@'})
        for sheet in xls.sheet_names:
            df = xls.parse(sheet, header=None)
            df.to_excel(writer, sheet_name=sheet, header=False, index=False)
            worksheet = writer.sheets[sheet]
            worksheet.set_column('A:Z', 12, cell_format)
        writer.save()
        files.append(outfile)
        logger.info('Deleting: {}'.format(infile))
        os.remove(infile)
    return files

def move_to_download_folder(dest_path, files):
    for file in files:
        filename = file.split('\\')[-1]
        prop_code = filename[0:4]
        if not os.path.exists(os.path.join(dest_path, prop_code)):
            os.makedirs(os.path.join(dest_path, prop_code))
        logger.info('Moving: {} to {}'.format(file, os.path.join(dest_path, prop_code, filename)))
        check_size = os.path.getsize(file)
        shutil.move(file, os.path.join(dest_path, prop_code, filename))
        while (check_size != os.path.getsize(os.path.join(dest_path, prop_code, filename))):
            time.sleep(1)