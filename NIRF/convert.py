#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Apr  4 20:47:34 2018

@author: Admin
"""

#import pdftables_api
#c = pdftables_api.Client('e9y684294pqa')
#c.xml('IR-5-O-OEMDP-U-0107.pdf', 'IR-5-O-OEMDP-U-0107.xml')

'''import requests # get it from http://python-requests.org or do 'pip install requests'

url = "http://pdfx.cs.man.ac.uk"

def pypdfx(filename):
    fin = open(filename + '.pdf', 'rb')
    files = {'file': fin}
    try:
        print('Sending', filename, 'to', url)
        r = requests.post(url, files=files, headers={'Content-Type':'application/pdf'})
        print('Got status code', r.status_code)
    finally:
        fin.close()
    fout = open(filename + '.xml', 'wb')
    fout.write(r.content)
    fout.close()
    print('Written to', filename + '.xml')

if __name__ == '__main__':
    filename = 'IR-5-O-OEMDP-U-0107'
    pypdfx(filename)'''

import os
import tabula

pdfs = ['/Users/Admin/Downloads/NIRF/Architecture/pdf', '/Users/Admin/Downloads/NIRF/Colleges/pdf', '/Users/Admin/Downloads/NIRF/Engineering/pdf', '/Users/Admin/Downloads/NIRF/Law/pdf', '/Users/Admin/Downloads/NIRF/Management/pdf', '/Users/Admin/Downloads/NIRF/Medical/pdf', '/Users/Admin/Downloads/NIRF/Overall/pdf', '/Users/Admin/Downloads/NIRF/Pharmacy/pdf', '/Users/Admin/Downloads/NIRF/University/pdf']
csvs = ['/Users/Admin/Downloads/NIRF/Architecture/csv', '/Users/Admin/Downloads/NIRF/Colleges/csv', '/Users/Admin/Downloads/NIRF/Engineering/csv', '/Users/Admin/Downloads/NIRF/Law/csv', '/Users/Admin/Downloads/NIRF/Management/csv', '/Users/Admin/Downloads/NIRF/Medical/csv', '/Users/Admin/Downloads/NIRF/Overall/csv', '/Users/Admin/Downloads/NIRF/Pharmacy/csv', '/Users/Admin/Downloads/NIRF/University/csv']

#base_command = 'tabula.convert_into("{}", "{}", output_format="csv", mutiple_tables=true)'
#tabula.convert_into("/Users/Admin/Downloads/NIRF/Architecture/pdf/IR-1-A-A-C-46330.pdf", "/Users/Admin/Downloads/NIRF/Architecture/csv/IR-1-A-A-C-46330.csv", output_format="csv", mutiple_tables="true")
base_command = 'java -jar tabula-1.0.1-jar-with-dependencies.jar -p all -f CSV -r -o {} {}'
sed_cmd = r"sed -i -e $'s/\r/ /g' {}"
print(sed_cmd)

for pdf_folder, csv_folder in zip(pdfs, csvs):
    for filename in os.listdir(pdf_folder):
        pdf_path = os.path.join(pdf_folder, filename)
        csv_path = os.path.join(csv_folder, filename.replace('.pdf', '.csv'))
        command = base_command.format(csv_path, pdf_path)
        os.system(command)
        command = sed_cmd.format(csv_path)
        os.system(command)

