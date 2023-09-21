import requests,openpyxl,re,time,os,traceback
from datetime import datetime
from bs4 import BeautifulSoup
from random import randrange
from os import getcwd,path,mkdir
# from py_mail import send_mail_python

start_time = datetime.now()
current_directory = os.getcwd()+"/"

import sys
file_path_inputs = sys.argv
file_path = file_path_inputs[1].replace('"','') if len(file_path_inputs) > 1 else "C:\bss_pm_auto"
inp_filename = file_path_inputs[2] if len(file_path_inputs) > 2 else "1.xlsx"
file = f"{file_path}\{inp_filename}"

print(start_time)

runtime_status = f"""
Start time: {start_time}
Script Status: 'Script Started'
"""

# try:
# 	mail_status = send_mail_python(f"Jackit - {datetime.now().strftime('%d_%m_%Y')}", "", file, "", runtime_status)
# 	print('Mail Send Successfully')
# except:
# 	print('Mail Send Failed')
# 	pass


def file_name_checker(name):
	for char in ['@','$','%','&','\\','/',':','*','?','"',"'",'<','>','|','~','`','#','^','+','=','{','}','[',']',';','!']:
		if char in name:
			name = str(name).replace(char, "__")
	return name
# inupt_files = os.getcwd() +"/HTML FILES"
# if not os.path.isdir(inupt_files):
# 	os.mkdir(inupt_files)

# Checking New Files folder exists and creating New Files folder Windows
inupt_files = f"{file_path}\\HTML Files\\"
if not path.isdir(inupt_files):
	mkdir(inupt_files)

# file_path = getcwd()
output_filename = '{}\Jackit_{}_{}.xlsx'.format(file_path,datetime.now().strftime('%d_%m_%Y'),inp_filename)
overall_prdt_list = []
script_error = ''
error_status = ''

def script_status(input_file_name,input_count,start_time,end_time,time_taken,error_status):
	if os.path.exists("Script Status.xlsx"):
		wb1 = openpyxl.load_workbook('Script Status.xlsx')
		ws1 = wb1.active
	else:
		wb1 = openpyxl.Workbook()
		ws1 = wb1.active
		ws1['A1'].value = "Input Filename"
		ws1['B1'].value = "Input Count"
		ws1['C1'].value = "Start Time"
		ws1['D1'].value = "End Time"
		ws1['E1'].value = "Processed Time"
		ws1['F1'].value = "Error Status"
		ws1['G1'].value = "Row Number"
	maxxx = ws1.max_row+1
	ws1.cell(maxxx,1).value = str(input_file_name)
	ws1.cell(maxxx,2).value = str(input_count)
	ws1.cell(maxxx,3).value = str(start_time)
	ws1.cell(maxxx,4).value = str(end_time)
	ws1.cell(maxxx,5).value = str(time_taken)
	ws1.cell(maxxx,6).value = str(error_status)
	wb1.save("Script Status.xlsx")
		
# while True:
# 	try:
# 		filepath = input('Enter File Name without extension  :   ')
# 		file = '{}.xlsx'.format(filepath)
# 		if os.path.isfile(file) == True:
# 			break
# 		else:
# 			print('Entered File does not Exists...Please Enter Valid File Name  :   ')
# 	except:
# 		filepath = input('Enter File Name without extension  :   ')
# 		file = '{}.xlsx'.format(filepath)
# file_name = input('Enter File name without Extension :') + '.xlsx'
# if os.path.exists(current_directory + file_name + ".xlsx") == True:
wb = openpyxl.load_workbook(inp_filename)
data_sheet = wb.active
row_max = data_sheet.max_row+1
data_sheet['A1'] = 'Input Brand'
data_sheet['B1'] = 'Input Sku'
data_sheet['C1'] = 'Brand'
data_sheet['D1'] = 'Sku'
data_sheet['E1'] = 'Domain Url'
data_sheet['F1'] = 'Sale Price'
data_sheet['G1'] = 'Product Url'
data_sheet['H1'] = 'Status'
data_sheet['I1'] = 'Scrape Price'
data_sheet['J1'] = 'Crawl Timestamp'
data_sheet['K1'] = 'Filename'
sleep_range = 500
try:
	for row_num in range(2,row_max):
		input_brand = data_sheet.cell(row_num,1).value
		input_sku = data_sheet.cell(row_num,2).value
		input_url = 'https://www.jackit.com/catalogsearch/result/?q={}'.format(input_sku)
		print(f'>>\-\-\-\>>..........{row_num-1}------------{input_sku}------>>)))------------->>>>\n')
		url = requests.get(input_url)
		soup = BeautifulSoup(url.content, 'html.parser')
		if str(url) == '<Response [200]>':
			page_check = soup.select('#product-wrapper')
			if page_check == []:
	#$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$-----------------< Land Page >------------------ $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
				sku = soup.select('.product-info-main .product-info-price .sku')[0].text.strip() if soup.select('.product-info-main .product-info-price .sku') != [] else ''
				price = soup.select('.product-info-price .price')[0].text.strip() if soup.select('.product-info-price .price') != [] else ''
				brand = soup.select('.product-info-main span strong')[0].text.strip() if soup.select('.product-info-main span strong') != [] else ''
				product_url = soup.select('link[rel="canonical"]')[0]['href'] if soup.select('link[rel="canonical"]') != [] else ''
				data_sheet.cell(row_num,3).value = brand
				data_sheet.cell(row_num,4).value = sku
				if str(input_sku).lower().strip() == str(sku).strip().lower() and str(input_brand).lower().strip() in str(brand).strip().lower():
					data_sheet.cell(row_num,6).value = price.replace('$', '')
					data_sheet.cell(row_num,9).value = price
					data_sheet.cell(row_num,7).value = product_url
					data_sheet.cell(row_num,8).value = 'Found'
					data_sheet.cell(row_num,11).value = str(file_name_checker(str(input_sku)))+'.html'
					data_sheet.cell(row_num,5).value = 'https://www.jackit.com'
					data_sheet.cell(row_num,10).value = datetime.now().strftime('%d-%m-%Y %I:%M %p')
					print('Found')
					with open(f'{inupt_files}/{str(file_name_checker(str(input_sku)))}.html', 'w+',encoding='utf-8') as file:
						file.write(str(soup))
						file.close()
				else:
					data_sheet.cell(row_num,8).value = 'Not Found'
					print('Not Found')
	#$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$----------------< List page >-------------------$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
			elif page_check != []:
				list_data = soup.select('.product.product-item-info')
				if len(list_data) != 0:
					loop_break = 0
					for itemer in list_data:
						sku_data = itemer.select('.product .list-manufacturer')[0].text.strip() if itemer.select('.product .list-manufacturer') != [] else ''
						sku_re = re.findall(r'(.*?)\|(.*)',str(sku_data))
						if sku_re != []:
							list_sku = str(sku_re[0][1]).strip()
							list_brand = str(sku_re[0][0]).strip()
							list_price = itemer.select('.product .price')[0].text.strip() if itemer.select('.product .price') != [] else ''
							list_url = itemer.select('.product .product-item-link')[0]['href'] if itemer.select('.product .product-item-link') != [] else ''
							data_sheet.cell(row_num,3).value = list_brand
							data_sheet.cell(row_num,4).value = list_sku
							if str(input_brand).strip().lower() in str(list_brand).lower().strip() and str(input_sku).strip().lower() == str(list_sku).lower().strip():
								print('Found')
								data_sheet.cell(row_num,6).value = list_price.replace('$','')
								data_sheet.cell(row_num,9).value = list_price
								data_sheet.cell(row_num,7).value = list_url
								# data_sheet.cell(row_num,8).value = 'Found'
								data_sheet.cell(row_num,11).value = str(file_name_checker(str(input_sku)))+'.html'
								data_sheet.cell(row_num,5).value = 'https://www.jackit.com'
								data_sheet.cell(row_num,10).value = datetime.now().strftime('%d-%m-%Y %I:%M %p')
								with open(f'{inupt_files}/{str(file_name_checker(str(input_sku)))}.html', 'w+',encoding='utf-8') as file:
									file.write(str(soup))
									file.close()
								loop_break = 1
						if loop_break == 1:
							break
					if loop_break == 1:
						data_sheet.cell(row_num,8).value = 'Found'
					else:
						data_sheet.cell(row_num,8).value = 'Not Found'
				else:	
					print('Not Found')
					data_sheet.cell(row_num,8).value = 'Not Found'

				# sku_data = soup.select('.product .list-manufacturer')[0].text.strip() if soup.select('.product .list-manufacturer') != [] else ''
				# sku_re = re.findall(r'(.*?)\|(.*)',str(sku_data))
				# # list_brand = ''
				# if sku_re != []:
				# 	list_sku = str(sku_re[0][1]).strip()
				# 	list_brand = str(sku_re[0][0]).strip()
				# 	list_price = soup.select('.product .price')[0].text.strip() if soup.select('.product .price') != [] else ''
				# 	list_url = soup.select('.product .product-item-link')[0]['href'] if soup.select('.product .product-item-link') != [] else ''
				# 	data_sheet.cell(row_num,3).value = list_brand
				# 	data_sheet.cell(row_num,4).value = list_sku
				# 	if str(input_brand).strip().lower() in str(list_brand).lower().strip() and str(input_sku).strip().lower() == str(list_sku).lower().strip():
				# 		data_sheet.cell(row_num,6).value = list_price.replace('$','')
				# 		data_sheet.cell(row_num,9).value = list_price
				# 		data_sheet.cell(row_num,7).value = list_url
				# 		data_sheet.cell(row_num,8).value = 'Found'
				# 		data_sheet.cell(row_num,11).value = str(file_name_checker(str(input_sku)))+'.html'
				# 		data_sheet.cell(row_num,5).value = 'https://www.jackit.com'
				# 		data_sheet.cell(row_num,10).value = datetime.now().strftime('%d-%m-%Y %I:%M %p')
				# 		print('Found')
				# 		with open(f'{inupt_files}/{str(file_name_checker(str(input_sku)))}.html', 'w+',encoding='utf-8') as file:
				# 			file.write(str(soup))
				# 			file.close()
				# 	else:
				# 		data_sheet.cell(row_num,8).value = 'Not Found'
				# 		print('Not Found')
			else:
				data_sheet.cell(row_num,8).value = 'Not Found'
				print('Not Found')
		else:
			print('Failed.......')
			data_sheet.cell(row_num,8).value = 'Not Found'
		if row_num == sleep_range:
			time.sleep(randrange(5,10))
			sleep_range += 500
		try:
			wb.save(output_filename)
			output_error = f"{output_filename} saved succesfully"
		except:
			input('-----> Please close the file <-----')
			wb.save(output_filename)
			output_error = f"Output Error \n {str(traceback.format_exc())}"
	script_error = "No Script Errors"
except:
	script_error = f"Script Error \n {str(traceback.format_exc())}"
try:
	wb.save(output_filename)
	output_error = f"{output_filename} saved succesfully"
except:
	input('-----> Please close the file <-----')
	wb.save(output_filename)
	output_error = f"Output Error \n {str(traceback.format_exc())}"
error_status = "\n\n".join([script_error, output_error])
end_time = datetime.now()
total_time = end_time - start_time

# try:
# 	mail_status = send_mail_python(f"Jackit - {current_date}", error_status, output_filename)
# except:
# 	pass

print('----------------------------------------------------------------------------------------------------')
print('Script Start Time : ',start_time)
print('Script End Time : ',end_time)
print('Total Time : ', total_time)
script_status(file_path,row_max,start_time,end_time,total_time,error_status)

runtime_status = f"""
Start time: {start_time}
End time: {end_time}
Time Taken: {end_time-start_time}
Script Status: 'Script Completed'
"""

# try:
# 	mail_status = send_mail_python(f"Jackit - {datetime.now().strftime('%d_%m_%Y')}", error_status, file_path, output_filename, runtime_status)
# 	print('Mail Send Successfully')
# except:
# 	print('Mail Send Failed')
# 	print(traceback.format_exc())
# 	pass
# input('Enter Any Key :')