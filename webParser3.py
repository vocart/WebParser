from html.parser import HTMLParser
import requests
import xlsxwriter


#preparing the file
def makeXlsxFile():
	file = xlsxwriter.Workbook('Lista książek.xlsx')
	addSheet = file.add_worksheet()
	# write headers
	addSheet.write('A1', 'Tytuł')
	addSheet.write('B1', 'Autor')
	addSheet.write('C1', 'Seria')
	addSheet.write('D1', 'Półka')

	#adding books to excel
	listOfBooks = makeBookObjects()

	for i in range(len(listOfBooks)):
		addSheet.write(i + 1, 0, listOfBooks[i].t)
		addSheet.write(i + 1, 1, listOfBooks[i].a)
		addSheet.write(i + 1, 2, listOfBooks[i].s)
		addSheet.write(i + 1, 3, listOfBooks[i].p)

	file.close()
	print('- Info uploaded to excel')


def makeBookObjects():
	fullPageContent = stringFormatting()
	listOfBooks = []
	class Book():
		def __init__(self, title, author, series, place):
			self.t = title
			self.a = author
			self.s = series
			self.p = place

			def __str__(self):
				return self.t

	for i in range (len(fullPageContent)):
		if fullPageContent[i].startswith('Na półkach:') is True:
			if fullPageContent[i-2] == 'Cykl: ' :
				correctName = fullPageContent[i+1].replace(fullPageContent[i-3], '')
				correctName2 = replacer(correctName)
				bookName = fullPageContent[i-4]
				bookName = Book(fullPageContent[i-4], fullPageContent[i-3], fullPageContent[i-1], correctName2)
				listOfBooks.append(bookName)
			else:
				correctName = fullPageContent[i+1].replace(fullPageContent[i-1], '')
				correctName2 = replacer(correctName)
				bookName = fullPageContent[i-2]
				bookName = Book(fullPageContent[i-2], fullPageContent[i-1], '---', correctName2)
				listOfBooks.append(bookName)
		else:
			pass

	print('- Made list of books ')
	return listOfBooks

	


def replacer(name):
	correctName2 = name.replace("Inne wydanie","").replace("David Weber","").replace("Lektury","")\
		.replace("Biografie sławnych ludzi","").replace("B. V. Larson","").replace("George R. R. Martin","")\
		.replace("Harlan coben","").replace("Ken Follet","").replace("Chris Bunch","").replace("Aleksander Dumas","")\
		.replace("Klub Interesującej...","").replace("Tom Clancy","").replace("R. J. Pineiro","")\
		.replace("James Patterson","").replace("Harry Harrison","").replace("Alistair MacLean","")\
		.replace("Współczesna proza...","").replace("Carl Hiaasen","").replace("Clive Cussler","")\
		.replace("PODMIEŃ..."," PODMIEŃ...").replace("Przeczytane","").replace("Posiadam","")\
		.replace("Pożyczone"," Pożyczone ! ").replace(",","")
	return correctName2


def stringFormatting():
	fullPageContent = extractFromPageContent()
	fullPageContent = fullPageContent.replace('  ', '').replace('\t', '').split('\n')
	fullPageContent = [line for line in fullPageContent if line != '']
	fullPageContent = [line for line in fullPageContent if not line[0].isdigit() and line != 'Ocenił na:' and line != 'Średnia ocen:']

	print('- Formatting data')
	return fullPageContent


# removes html formatting, leaves content
def extractFromPageContent():
	fullPageContent = makeFullPageContent()
	class HTMLFilter(HTMLParser):
	    text = ""
	    def handle_data(self, data):
	        self.text += data

	f = HTMLFilter()
	f.feed(fullPageContent)
	fullPageContent = f.text
	print('- Extraction of non-html data')
	return fullPageContent


def makeFullPageContent():

	fullPageContent = ''

	for pageNumber in range(1,65):

		data = {"page": pageNumber, "listId" : "booksFilteredList", "shelfs": "5796074", "showFirstLetter": "0", "paginatorType": "Standard", "objectId": "1770105"}
		custom_header = {"X-Requested-With": "XMLHttpRequest"}
		url = "https://lubimyczytac.pl/profile/getLibraryBooksList"

		response = requests.post(url, data, headers=custom_header)
		content = response.json()["data"]["content"]
		fullPageContent += content
		print('page:', pageNumber, 'added')

	print('- Adding all data from site')
	return fullPageContent
	



makeXlsxFile()
