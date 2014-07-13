#!/usr/bin/env python
# -*- coding: utf-8 -*-

import httplib2
from bs4 import BeautifulSoup, SoupStrainer
import win32clipboard
import xlsxwriter


http = httplib2.Http()

outfile = open('output.txt', 'w', encoding='utf-8')
workbook = xlsxwriter.Workbook('result_export.xlsx')
worksheet = workbook.add_worksheet()


status, response = http.request('http://blv-sport.de/service/msonline/files/20145720-e.htm', "GET")
response = response.decode('iso-8859-1')

soup = BeautifulSoup(response)

soup.style.decompose()
soup.head.decompose()


Meisterschaft = soup.select(".KopfZ11")[0].get_text()
Datum_Ort = soup.select(".KopfZ12")[0].get_text()
# Format Datum_Ort: "am DATUM in ORT"
position_in = Datum_Ort.find("in")
Datum = Datum_Ort[3:position_in]
Ort = Datum_Ort[position_in+3:]

print("Meisterschaft: " + Meisterschaft)
print("Datum: " + Datum)
print("Ort: " + Ort)

soup.select(".KopfZ1")[0].parent.parent.decompose()
soup.select(".KopfZ21")[0].parent.parent.decompose()
soup.select(".KopfZ2")[0].parent.parent.decompose()


heading_veranstaltungs_bericht = soup.find("a", {"name":"VERANSTALTUNGS-BERICHT"}).parent.parent.parent.parent

data = soup.prettify()

pos_bericht = data.find('name="VERANSTALTUNGS-BERICHT"')

print(pos_bericht)

data = data[:pos_bericht]

soup = BeautifulSoup(data)
# remove all <br/>
brs = soup.find_all("br")
for tag in brs:
		soup.br.unwrap()

worksheet.write(0, 0, Meisterschaft)
worksheet.write(0, 1, Datum)
worksheet.write(0, 2, Ort)

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
