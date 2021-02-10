import os
# import sys
import shutil
import logging
import json
import datetime as dt
import pandas as pd
# import time
# import uuid
# import argparse
# import re
# import json


from dusit.utils import *

logging.basicConfig(format=LOG_FORMAT, filename='../log/ideas_{}'.format(dt.datetime.today().strftime('%Y%m%d%H%M%S_%f')),level=logging.INFO)

if __name__ == '__main__':
    # read config
    logger.info('Reading configuration file ...')
    with open('../config/ideas_conf.json') as f:
        config = json.load(f)
    logger.info('IDEAS config: {}'.format(json.dumps(config, indent=4)))

    source_path = config['source_path']
    dest_path = config['dest_path']

    load_date = dt.date.today().strftime('%Y%m%d') # current date
    try:
        logger.info('Convert files from Logic Apps ...')
        files = convert_xlsx(source_path)
        logger.info('Convert files from Logic Apps ... Done')
        # print(files)
        logger.info('Move converted files to Azure File Storage ...')
        move_to_download_folder(dest_path, files)
        logger.info('Move converted files to Azure File Storage ... Done')
    except Exception as e:
        logger.error('An error occurs while converting the file. \nReason: {}'.format(e))
    logger.info('========= COMPLETED =========')
