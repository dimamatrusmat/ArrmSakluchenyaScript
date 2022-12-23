from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from docxtpl import DocxTemplate
import requests as req
import pandas as pd
import os

ua = UserAgent()

def getRota(fio):

	os.chdir("../")

	file = 'spel.xlsx'
	df = pd.read_excel(file)

	os.chdir("date")
	
	for i in range(len(df['fio'])):
		if fio == df['fio'][i]:
			return df['rot'][i], df['zv'][i], df['name_zv'][i], i

	

	return "Неверно фио", "", "", ""

def go_genetif(fio):
	fio = fio.split(" ")

	lastname = fio[0].strip()
	name = fio[1].strip()
	middlename = fio[2].strip()

	url = 'https://surnameonline.ru/inflect.php'

	context = {
		'surname': lastname,
		'name': name,
		'patronymic': middlename
	}

	h = ua.random
	headers = {'User-Agent': h}

	resp = req.post(url, context,  headers) 
	soup = BeautifulSoup(resp.text, 'lxml')
	text = soup.html.body.ul.find_all('li')[1]
	text = text.text[5:]


	return text

def read_file(filename, all_rota):

	with open(filename) as file:
		i = 0

		for line in file:
			
			if i == 1:
				fio = line
				fio = fio.strip()
				fio1 = fio
				
			if i == 2:
				year1 = line
				year= year1[12:16]
				year1 = year1[6:16]
			if i == 6:
				text0 = line.strip()
			if i == 7:
				text1 = line.strip()	
			if i == 8:
				text2 = line.strip()
			if i == 9:
				text3 = line.strip()
			if i == 10:
				text4 = line.strip()	
			if i == 11:
				text5 = line.strip()
			if i == 12:
				text6 = line.strip()
			if i == 13:
				text7 = line.strip()	
			if i == 14:
				text8 = line.strip()
				
			name_zv = ''
			zv = ''
			i+=1


		if i >= 15:

			rota, zv, name_zv, number = getRota(fio)
			getdan = [fio,year1,rota, filename, number]
			all_rota.append(getdan)

		else:
			print("Не тот тест" + filename)		


if __name__ == "__main__":
	os.chdir("date")
	allfiles = os.listdir('.')
	all_rota = []
	
	i = 1
	for file in allfiles:
		read_file(file, all_rota)
		print(str(100 - round(len(allfiles)/i, 2))+ "%")

		i+=1

	df1 = pd.DataFrame(all_rota, columns=['fio','year','rota', 'filename', 'number'])
	os.chdir("../")
	df1.to_excel("otchet_play.xlsx", index=False)
