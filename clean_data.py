# Note: Manual setting ASIN col of Simulate sheet in clean_ASIN function
import json
import datetime
# from set_spreadsheet import set_spreadsheet, sheet_write, set_spreadsheet_o_o, set_spreadsheet_Research, add_many_list

##################################################
# Ex: 37.74/ea > 37.74
# Convert str_ to float
def to_float(str_):

	if type(str_) in [float,int] or str_ == '':
		return str_
	else:
		res = ''
		for i in str_:
			if i.isdigit() or i in ['.', '-']:
				res += i
		try:
			res = float(res)
		except:
			res = ''
			print('   to_float error. Can not convert to fload')
		return res

##################################################
if __name__ == '__main__':
	print('    Have not main function.')
	# row = get_next_row(spreadsheet= 'Products Research',sheet='eBay_Scraping', col='Lower_Price_1', row ='')
	# print(type(row))
	# copy_history()
	# print(to_float('123'))
	# clean_ASIN()
	# clean_NKV()

