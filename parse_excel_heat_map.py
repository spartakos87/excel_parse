import xlrd
from pprint import pprint as pp
from collections import OrderedDict
from json import dump

def parse_excel(f):
	book = xlrd.open_workbook('new_for_parsing.xlsx')
	sheet = book.sheet_by_name('Summary (3 options)')
	data_dict = OrderedDict()	
	for model in range(3,20+1):
		#temp = {}
		temp = OrderedDict()
		model_title = sheet.cell(1,model).value
		if model_title != '':
			if model_title == 42.0:
				model_title = '42'
			for sdg in range(2,54+1):
				value = sheet.cell(sdg,model).value
				sdg_title = sheet.cell(sdg,1).value
				if sdg_title != '':
					if value !='':
						value =int(value)
					temp[sdg_title] = value
			data_dict[model_title] = temp
	with open('file_parsing_data.json','w') as fp:
		dump(data_dict, fp)
	return data_dict
