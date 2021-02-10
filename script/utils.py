import os
import shutil
import time
import logging
import pandas as pd

import re
import json
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

LOG_FORMAT = '%(asctime)s [%(levelname)s] %(name)s: %(message)s'

logger = logging.getLogger('DUSIT')

def find_variable(markup, re_pattern):
    '''
    to find variable value in html code
    '''
    soup = BeautifulSoup(markup=markup, features='html.parser')
    p = re.compile(re_pattern)
    output = []
    for script in soup.find_all('script'):
        matched = p.search(script.text)
        if matched:
            output = json.loads(matched.group(1))
            break
    return output

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

def download_wait(directory, timeout, nfiles=None):
    """Wait for downloads to finish with a specified timeout.

    Args:
        directory (str): The path to the folder where the files will be downloaded.
        timeout (int): How many seconds to wait until timing out.
        nfiles (int): If provided, also wait for the expected number of files. defaults to None

    Returns:
        The return value. True for success, False otherwise.
    """
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False     
        files = list(filter(lambda file: os.path.isfile(file), map(lambda filename: os.path.join(directory, filename), os.listdir(directory))))
        logger.info('downloaded files: {}'.format(files))        
        if nfiles and len(files) < nfiles:
            dl_wait = True

        for fname in files:
            if fname.endswith('.crdownload'):
                dl_wait = True

        seconds += 1
    return files

def convert_xls_to_xlsx(path, property_code, load_date):
    files=[]
    for filename in filter(lambda x: x.endswith('xls') or x.endswith('xlsx'), os.listdir(path)):
        infile = os.path.join(path, filename)
        outfile = os.path.join(path, _change_filename(filename, property_code, load_date))
        logger.info('Converting {} to {}'.format(infile, outfile))
        xls = pd.ExcelFile(infile)
        writer = pd.ExcelWriter(outfile, engine='xlsxwriter')
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
        if not os.path.exists(os.path.join(to_path, extension)):
            os.makedirs(os.path.join(to_path, extension))

        check_size = os.path.getsize(file)
        shutil.move(file, os.path.join(to_path, extension, filename))
        while (check_size != os.path.getsize(os.path.join(to_path, extension, filename))):
            time.sleep(1)

def get_chrome_options(download_path, headless=True):
    chrome_options = webdriver.ChromeOptions()
    if headless:
        chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-notifications')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--verbose')
    chrome_options.add_experimental_option("prefs", {
            "download.default_directory": download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing_for_trusted_sources_enabled": False,
            "safebrowsing.enabled": False,
            "credentials_enable_service": False,
            "profile.default_content_setting_values.notifications" : 2,
            "profile": {
                'password_manager_enabled': False
            }
    })
    return chrome_options

def str_login(driver, str_username, str_password):
    '''
    shortcut to login to dstar.str.com
    '''
    driver.find_element_by_xpath('//*[@id="username"]').clear()
    driver.find_element_by_xpath('//*[@id="username"]').send_keys(str_username)
    driver.find_element_by_xpath('//*[@id="password"]').clear()
    driver.find_element_by_xpath('//*[@id="password"]').send_keys(str_password)
    driver.find_element_by_xpath('//button[@class="md-btn-raised"]').click()