#!/usr/bin/env python
# -*- coding: utf-8 -*-

import httplib2
from bs4 import BeautifulSoup, SoupStrainer
import win32clipboard
import xlsxwriter
http = httplib2.Http()

result_url = "http://blv-sport.de/service/msonline/files/20145720-e.htm"

print("Retrieving URL: "+result_url)

outfile = open('output.txt', 'w')
workbook = xlsxwriter.Workbook('result_export.xlsx')
worksheet = workbook.add_worksheet()

status, response = http.request(result_url, "GET")
response = response.decode('iso-8859-1')

soup = BeautifulSoup(response)

Meisterschaft = soup.select(".KopfZ11")[0].get_text()
Datum_Ort = soup.select(".KopfZ12")[0].get_text()
# Format Datum_Ort: "am DATUM in ORT"
position_in = Datum_Ort.find("in")
Datum = Datum_Ort[3:position_in]
Ort = Datum_Ort[position_in+3:]

print("Extracting: ")
print("Meisterschaft: " + Meisterschaft)
print("Datum: " + Datum)
print("Ort: " + Ort)

soup.style.decompose()
soup.head.decompose()
soup.select(".KopfZ1")[0].parent.parent.decompose()
soup.select(".KopfZ21")[0].parent.parent.decompose()
soup.select(".KopfZ2")[0].parent.parent.decompose()

data = soup.prettify()
pos_bericht = data.find('name="VERANSTALTUNGS-BERICHT"')
data = data[:pos_bericht]
soup = BeautifulSoup(data)

# remove all <br/>
brs = soup.find_all("br")
for tag in brs:
	soup.br.unwrap()

		

worksheet.write(0, 0, Meisterschaft)
worksheet.write(0, 1, Datum)
worksheet.write(0, 2, Ort)

worksheet.set_column('A:A', 10)
worksheet.set_column('B:B', 30)
worksheet.set_column('E:E', 30)

row_excel=1
tables = soup.find_all("table")
for table in tables:
	row_excel = row_excel + 1
	rows = table.findAll('tr')

	for tr in rows:
		cols = tr.findAll('td')
		column_excel = 0
		for td in cols:
			text = td.get_text()
			text = text.rstrip()
			text = text.lstrip()
			text = text.strip(' \t\n\r')
			worksheet.write(row_excel, column_excel, text)
			column_excel = column_excel + 1

workbook.close()	

outfile.write(soup.prettify())	
outfile.close()

input("Export successful. Press Enter to close.")
