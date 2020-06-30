
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
# from webdriver_manager.chrome import ChromeDriverManager
import os
import time
from datetime import datetime
import pytz
import shutil
import zipfile
import re
import csv

import PyPDF2
from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser

def getDownloadsFolderPath():
    currentpath = os.path.realpath(__file__)
    print(currentpath)
    # add_to_startup(currentpath)

    patharr = currentpath.split('\\')
    patharr[len(patharr)-1] = 'downloads'
    ptr = ('\\').join(patharr)
    return ptr
def initDriver():
  options = webdriver.ChromeOptions()
  download_dir = getDownloadsFolderPath()
  profile = {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}], 
              "plugins.always_open_pdf_externally": True,
              "download.default_directory": download_dir}
  # profile = {"plugins.always_open_pdf_externally": True}
  options.add_experimental_option("prefs", profile)
  
  driver = webdriver.Chrome(options=options)

  return driver

driver = initDriver()

def writeUrls(item_desc):
  filename = "itemurls.csv"
  try: 
    f = open(filename, "a")
    for itemurl in item_desc:
      f.write(itemurl+ "\n")
    f.close()
  except:
    print("There was an error writing to the CSV data file.")

def writeProductsInfo(productsinfo):
    filename = "gjoco_no.csv"
    file_exists = os.path.isfile(filename)
    header = ['source',	'manufacturer_name', 'product_url',	'product_article_id',	'sds_pdf',	'sds_source',	'sds_language',	'product_name',	'sds_pdf_product_name',
              'sds_pdf_published_date',	'sds_pdf_revision_date',	'sds_pdf_manufacture_name',	'sds_pdf_Hazards_identification',	'sds_filename',	'crawl_date']
    with open(filename, 'a', encoding="utf-8-sig") as csvfile:
        # writer = csv.writer(csvfile,quotechar='|',  quoting=csv.QUOTE_ALL)
        writer = csv.DictWriter(csvfile, delimiter=',', lineterminator='\n',fieldnames=header)

        if not file_exists:
            writer.writeheader()  # file doesn't exist yet, write a header
        for row in productsinfo:
            writer.writerow({
              'source':row[0],	'manufacturer_name':row[1], 'product_url':row[2],	'product_article_id':row[3],	'sds_pdf':row[4],	'sds_source':row[5],	'sds_language':row[6],	'product_name':row[7],	'sds_pdf_product_name':row[8], 'sds_pdf_published_date':row[9],	'sds_pdf_revision_date':row[10],	'sds_pdf_manufacture_name':row[11],	'sds_pdf_Hazards_identification':row[12],	'sds_filename':row[13],	'crawl_date':row[14]
            })
        
        
def getProductSdsInfo(file_path):
    print('getProductSdsInfo')
    output_string = StringIO()
    with open(file_path, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)

    content = output_string.getvalue()
    pattern0 = re.compile(r'[^\+]H[1-9][1-9][1-9][^\+]')
    pattern1 = re.compile(r'H[1-9][1-9][1-9]')
    pattern2 = re.compile(r'H[1-9][1-9][1-9][\+]H[1-9][1-9][1-9]')
    harzards_arr = []
    initarrforfirst = re.findall(pattern0, content)
    for init in initarrforfirst:
        final = re.findall(pattern1, init)[0]
        harzards_arr.append(final)
    secondarr = re.findall(pattern2, content)
    harzards_arr.extend(secondarr)
    harzards_arr = list(set(harzards_arr))
    sds_pdf_Hazards_identification = ', '.join(harzards_arr)
    
    # print(content)
    content_arr = content.split('\n')


    sds_pdf_product_name = content_arr[6]
    

    try:
        ind = content_arr.index('Utgitt dato')
        ptr2 = content_arr[ind+2].split('.')
        sds_pdf_published_date = '{}/{}/{}'.format(ptr2[0], ptr2[1], ptr2[2])
    except:
        sds_pdf_published_date = None


    try:
        ind = content_arr.index('Revisjonsdato')
        ptr2 = content_arr[ind+2].split('.')
        sds_pdf_revision_date = '{}/{}/{}'.format(ptr2[0], ptr2[1], ptr2[2])
    except:
        sds_pdf_revision_date = None


    try:
        ind = content_arr.index('Firmanavn')
        ptr2 = content_arr[ind+2].split('.')
        sds_pdf_manufacture_name = '{}/{}/{}'.format(ptr2[0], ptr2[1], ptr2[2])
    except:
        sds_pdf_manufacture_name = None

    return [sds_pdf_product_name,	sds_pdf_published_date,	sds_pdf_revision_date,	sds_pdf_manufacture_name,	sds_pdf_Hazards_identification]




def compressToZip(filename):
    zipname = filename.split('.')[0]+'.zip'
    zf = zipfile.ZipFile(
      'downloads/'+zipname, "w", zipfile.ZIP_DEFLATED)
    zf.write('downloads/'+filename, filename)
    zf.close()
    return 'downloads/'+zipname

def getProductsUrl(site_links):
  productsurl_arr = []
  for link in site_links:
    driver.get(link)
    while True:
      soup = BeautifulSoup(driver.page_source, 'html.parser')
      rowli_arr = soup.find("ul", {"class": "list_pro clo3"}).findChildren("li" , recursive=False)
      for rowli in rowli_arr:
        ptr = rowli.find("a").get('href')
        producturl = 'https://gjoco.no/' + ptr
        productsurl_arr.append(producturl)
      break
      try:
        nextbut = driver.find_element_by_class_name('go-next')
        nextbut.click()
        # break
      except:
        break
      
    break
    
  return productsurl_arr
    
def getProductInfo(link):
  # driver = initDriver()
  print('get product info')
  driver.get(link)
  time.sleep(3)
  soup = BeautifulSoup(driver.page_source, 'html.parser')
  
  # print(soup)
  titleptr = soup.find('div', {'class':'breadcum_name'}).text
  title = titleptr.split('\n')[2].strip()
  print('title', title)
  
  downloadbutts = driver.find_elements_by_xpath("//a[contains(text(), 'Sikkerhetsdatablad')]")
  print(downloadbutts)
  downloadbutt = downloadbutts[len(downloadbutts)-1]
  pdfurl = downloadbutt.get_attribute('href')
  print(pdfurl)
  # downloadbutt.click()
  
  driver.get(pdfurl)
  
  
  # try:
  #   downloadbutts = driver.find_elements_by_xpath("//a[contains(text(), 'Sikkerhetsdatablad')]")
  #   downloadbutt = downloadbutts[len(downloadbutts)-1]
  #   # print(downloadbutt)
  #   pdfurl = downloadbutt.get_attribute('href')
  #   # soup = BeautifulSoup(driver.page_source, 'html.parser')
  #   # pdfurl = soup.find('div', {'data-component-uuiddata-component-uuiddata-component-uuid':'NO'}).find('a').['href']
  
  #   print(pdfurl)
  #   downloadbutt.click()

  #   print('success download')
  # except:
  #   return None
  def getCurrentFileNmae(path_to_downloads):
    id = 0
    while True:
        time.sleep(0.1)      
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                filename = fname.split('.')[0] + '.pdf'
                return filename
        id += 1
    if id > 100:
        return None                
  pdffilename = getCurrentFileNmae('downloads')
  print('pdffilename: ', pdffilename)
  if pdffilename is None:
    print('return None')
    return None
    
  file_path = 'downloads/' + pdffilename
  print('filepath', file_path)
  while True:
      try:
          open(file_path, 'r')
          break
      except:
          time.sleep(1)
  # driver.close()
  # driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 'w') 
  # soup = BeautifulSoup(driver.page_source, 'html.parser')

  # pdffilename = pdfurl.split('/')[4].split('?')[0]

  print(pdffilename)
  Sds_info = getProductSdsInfo(file_path)
  print(pdfurl)
  source = 'gjoco.no.py'
  manufacturer_name = 'gjoco.no'
  product_url = link
  product_article_id = None
  sds_pdf = pdfurl
  sds_source = link
  sds_language = "Norwegian"
  product_name = title

  sds_filename_in_zip = file_path


  tz = pytz.timezone('Europe/Oslo')
  today = datetime.now(tz=tz)
  crawl_date = today.strftime("%d/%m/%Y")
  productinfo = [source, manufacturer_name, product_url,	product_article_id,	sds_pdf, sds_source,	sds_language,	product_name]
  productinfo.extend(Sds_info)
  productinfo.extend([sds_filename_in_zip,	crawl_date])
  return productinfo


site_links = ['https://gjoco.no/no/kategori/produkt']
products_url_arr = getProductsUrl(site_links)
print(products_url_arr)
productsinfo_arr = []
# link = products_url_arr[0]
# product_info = getProductInfo(link)
# if product_info is None:
#   print(link)
# else:
#   productsinfo_arr.append(product_info)
ind = 0
for link in products_url_arr:
  if ind > 3:
    break
  product_info = getProductInfo(link)
  if product_info is None:
    print('product link: ', link)
  else:
    print('product link: ', link)
    productsinfo_arr.append(product_info)
  ind += 1
  print('product num => ', ind)
print(productsinfo_arr)
writeProductsInfo(productsinfo_arr)

time.sleep(3)
driver.close()
