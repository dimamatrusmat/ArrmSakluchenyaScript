from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from docxtpl import DocxTemplate
import requests as req
import pandas as pd
import os

ua = UserAgent()
df_fio = pd.read_excel('fio.xlsx')
os.chdir("date")
new_fio = []

def get_fio_in_file(fio):
	
	list_fio = len(df_fio['fio'])
	for i in range(list_fio):
		if df_fio['fio'][i] == fio:
			return df_fio['fio_r'][i]
	
	return 0

def get_fio():

	list_fio = len(df_fio['fio'])
	for i in range(list_fio):
		new_fio.append([df_fio['fio'][i],df_fio['fio_r'][i]])

def save_fio():
			
	df_fio1 = pd.DataFrame(new_fio, columns=['fio','fio_r'])
	df_fio1.to_excel('fio.xlsx', index=False)
			

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
	ans = get_fio_in_file(fio)

	if ans == 0:
		fio1 = fio.split(" ")

		lastname = fio1[0].strip()
		name = fio1[1].strip()
		middlename = fio1[2].strip()

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

		new_fio.append([fio, text])

		return text

	else:
		return ans

def read_file(filename, all_rota):

	with open(filename) as file:
		i = 0

		for line in file:
			
			if i == 1:
				fio = line
				fio = fio.strip()
				fio1 = fio
				try:
					fio1 = go_genetif(fio)
				except:
					print(fio)
				
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
			os.chdir("../")

			doc = DocxTemplate("shablon.docx")
			context = {'fio_rr': fio1, 'year': year,'text0': text0,'text1': text1,'text2': text2,'text3': text3, 'text4': text4, 'text5': text5, 'text6': text6, 'text7': text7, 'text8': text8, 'zv': zv, 'name_zv': name_zv}
			doc.render(context)

			newpath = "final"
			if not os.path.exists(newpath):
				os.makedirs(newpath)

			os.chdir("final")
			rota = str(rota)
			if "/" in str(rota):
				newpath = rota.split("/")[0] + " рота"

				if not os.path.exists(newpath):
					os.makedirs(newpath)

				os.chdir(newpath)	

				newpath = rota.split("/")[1] + " взвод"

				if not os.path.exists(newpath):
					os.makedirs(newpath)

				os.chdir(newpath)	

				doc.save(fio + ".docx")

				os.chdir("../")

				os.chdir("../")

			else:

				newpath = rota
				if not os.path.exists(newpath):
					os.makedirs(newpath)

				os.chdir(newpath)

				doc.save(fio + ".docx")

				os.chdir("../")


			os.chdir("../")

			os.chdir("date")

		else:
			print("Не тот тест" + filename)		


if __name__ == "__main__":
	allfiles = os.listdir('.')
	all_rota = []
	get_fio()

	i = 1
	for file in allfiles:
		read_file(file, all_rota)
		print(str(100 - round(len(allfiles)//i, 2))+ "%")

		i+=1
	os.chdir("../")
	save_fio()	
	
	df1 = pd.DataFrame(all_rota, columns=['fio','year','rota', 'filename', 'number'])
	
	df1.to_excel("otchet.xlsx", index=False)
