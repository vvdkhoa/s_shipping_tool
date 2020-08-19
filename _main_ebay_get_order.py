from ebaysdk.trading import Connection as Trading
from ebaysdk.exception import ConnectionError
from trading import init_options # trading.py in same folder
from common import dump
from datetime import datetime, timedelta #, timezone
import unicodedata
from set_spreadsheet import set_spreadsheet_ship_table, sheet_write_ship_table

###########################################################################
# https://developer.ebay.com/devzone/xml/docs/reference/ebay/getorders.html
# https://a-zumi.net/ebay-traging-api-unshipped-orders/
# Get max 100 order, Priority for earlier time
# Date time obj: https://qiita.com/t-iguchi/items/a0bb8a5f273b319e5755#5datetime%E3%81%AE%E5%80%A4%E3%82%92%E5%8F%96%E3%82%8A%E5%87%BA%E3%81%99
def getOrders(opts, time_from, time_to):

	orders = {}
	try:
		api = Trading(debug=opts.debug, config_file=opts.yaml, appid=opts.appid,
		              certid=opts.certid, devid=opts.devid, warnings=True, timeout=120)

		api.execute('GetOrders', {
			"CreateTimeFrom": time_from,
			"CreateTimeTo": time_to,
			# "NumberOfDays": 5,
			"OrderRole": "Seller",
			"OrderStatus": "Completed",
			"IncludeFinalValueFee": True
			})
		
		if api.response.reply.OrderArray != None:

			for order in api.response.reply.OrderArray.Order: # 取得した注文データから各種データを取得

	        	# ShippedTime属性がある場合は発送済みなのでスルー
	        	# if hasattr(order, "ShippedTime") is True:
	        	# 	continue

				for txn in order.TransactionArray.Transaction:

					try: 	# Get variation info, if variation_name = '' => single list
						sku = txn.Variation.SKU
						variation_name = txn.Variation.VariationSpecifics.NameValueList.Value
						variation_url = txn.Variation.VariationViewItemURL
					except:
						try:
							sku = txn.Item.SKU
						except:
							print('    No SKU, Please check buyer ID: {}'.format(order.BuyerUserID))
							sku = ''
						variation_name = ''
						variation_url = ''
						pass

					try: 	# Get tracking info, tracking != '' or tracking = 'No Tracking' => shipped
						tracking = txn.ShippingDetails.ShipmentTrackingDetails.ShipmentTrackingNumber
					except:
						if hasattr(order, "ShippedTime") is True:
							tracking = 'No Tracking'
						else:
							tracking = ''
					try:
						shipped_time = order.ShippedTime
					except:
						shipped_time = ''
						pass

					try:	# Buyer message
						message = order.BuyerCheckoutMessage
					except:
						message = ''

					try:
						PaidTime = order.PaidTime
					except:
						PaidTime = ''

					data = {
						'order_id': order.OrderID, #注文ID
						'record': txn.ShippingDetails.SellingManagerSalesRecordNumber,
						'sku': sku, # SKU
						'buyer_id': order.BuyerUserID,
						'message': message,
						'order_item_id': txn.Item.ItemID, #商品ID
						'created_time': order.CreatedTime, #注文日時
						'paid_time': PaidTime,
						'product_name': txn.Item.Title, # 商品名
						'variation_name': variation_name,
						'variation_url': variation_url,
						'quantity_purchased': txn.QuantityPurchased, # 数量
						'price': txn.TransactionPrice.value, # 単価
						'fee': txn.FinalValueFee.value, # 手数料
						'shipping_cost': txn.ActualShippingCost.value, # 送料
						'ship_name': order.ShippingAddress.Name,
						'ship_address1': order.ShippingAddress.Street1,
						'ship_address2': order.ShippingAddress.Street2,
						'ship_city': order.ShippingAddress.CityName,
						'ship_state': order.ShippingAddress.StateOrProvince,
						'ship_postal_code': order.ShippingAddress.PostalCode,
						'ship_country': order.ShippingAddress.CountryName,
						'ship_phone': order.ShippingAddress.Phone,
						'ship_service': order.ShippingServiceSelected.ShippingService,
						'ship_paid': order.ShippingServiceSelected.ShippingServiceCost.value,
						'tracking': tracking,
						'shipped_time': shipped_time
						}

					orders[data['record']] = data

		else:
			print('     > There are no orders for this period')

	except ConnectionError as e:
	    print(e)
	    print(e.response.dict())

	return orders

###########################################################################
def get_all_order(days_ago):

	period = 1	# = n*0.5 Chia nho khoang thoi gian lay thong tin, min = 0.5 days, max 1day, > 1day khong lay duoc het thong tin gan nhat
	get_time = int(days_ago/period)

	now = datetime.now()

	if days_ago <= period:
		time_from = now  - timedelta(days=days_ago)
		time_to = now 
		all_orders = getOrders(opts, time_from, time_to)
		print('    Get orders data from {} to {}'.format(time_from, time_to))
	else:
		all_orders= {}
		for n in range(get_time):
			time_from = now  - timedelta(days_ago - period*n)
			time_to = now  - timedelta(days_ago - period*(n+1))
			print('    Get orders from {} to {}'.format(time_from, time_to))
			all_orders.update(getOrders(opts, time_from, time_to))

	return all_orders


###########################################################################
# Get record list from all orders
def get_record_list(all_orders):

	record_list = []
	for i in all_orders:
		record_list.append(i)

	record_list.sort()

	return record_list

###########################################################################
# Get shipping service from all orders
# return {'12059': 'Standard', '12060': 'Expedited', '12061': 'Economy',...}
def get_ship_service(all_orders):

	res_ship_service ={}
	for record in all_orders:

		# Shipping service
		ship_service = all_orders[record]['ship_service']

		if ship_service[:5] == 'Other' or ship_service[:7] == 'Economy':
			ship_service = 'Economy'
		elif ship_service[:8] == 'Standard' or ship_service[-8:] == 'Standard':
			ship_service = 'Standard'
		elif ship_service[:9] == 'Expedited' or ship_service[-7:] == 'Express':
			ship_service = 'Expedited'
		else:
			print(    'Please check shipping service, record {}'.format(record))

		res_ship_service[record] = ship_service

	return res_ship_service


###########################################################################
# Get and check the address length
def check_address(all_orders):

	address_check = {}
	country_dic = {}

	for record in all_orders:

		name = all_orders[record]['ship_name']			# Max 80 characters
		address1 = all_orders[record]['ship_address1']	# Max 80
		address2 = all_orders[record]['ship_address2']	# Max 80
		city = all_orders[record]['ship_city']			# Max 36
		state = all_orders[record]['ship_state']		# Max 24
		postal_code = all_orders[record]['ship_postal_code']	# Max 20
		phone = all_orders[record]['ship_phone']		# Max 20
		country = all_orders[record]['ship_country']

		check = []

		# Check the address length https://note.nkmk.me/python-unicodedata-east-asian-width-count/
		if (0 if name is None else len_check(name)) > 80:
			check.append('Name too long')
		if (0 if address1 is None else len_check(address1)) > 80:
			check.append('Address 1 too long')
		if (0 if address2 is None else len_check(address2)) > 80:
			check.append('Address 2 too long')
		if (0 if city is None else len_check(city)) > 36:
			check.append('City too long')
		if (0 if state is None else len_check(state)) > 24:
			check.append('State too long')
		if (0 if postal_code is None else len_check(postal_code)) > 20:
			check.append('Postal Code too long')
		if (0 if phone is None else len_check(phone)) > 20:
			check.append('Tel too long')
		check = ', '.join(check)
		address_check[record] = check

		country_dic[record] = country

	return (address_check, country_dic)

###########################################################################
def len_check(text): # Alphanumeric characters length
    count = 0
    for c in text:
        if unicodedata.east_asian_width(c) in 'FWAN':  # N: Thai, Arabic, W: Japanese, A: Vietnamese, 
            count += 2
        else:
            count += 1
    return count


###########################################################################
# Write shipping service to 差出表_作業指示書
def write_ship_info(ship_service, check_address, country_dic):

	print('Write shipping info to 差出表_作業指示書')

	# 差出表_作業指示書 > Country sheet
	Service_table = get_ship_service_table()

	# Old recode in 差出表_作業指示書 > CheckShipping sheet
	CheckShipping = set_spreadsheet_ship_table('CheckShipping')
	Old_Record = CheckShipping.col_values(1)

	Add_Record = []
	for record in ship_service: # New_Record
		if record not in Old_Record:

			ship_service_jp = get_jppost_service(
								Service_table, service_name=ship_service[record], country_name=country_dic[record])
			Add_Record.append([ 
				int(record), ship_service[record], country_dic[record], check_address[record], ship_service_jp ])

	if Add_Record != []:
		start_row = len(Old_Record) + 1
		w_data = []
		for row in Add_Record:
			w_data += row
		sheet_write_ship_table(SHEET_NAME='CheckShipping', Start_row=start_row, Start_col=1, 
			End_row=start_row+len(Add_Record)-1, End_col=len(Add_Record[0]), update_list=w_data)


###########################################################################
# Return Service_table = 
# {'Afghanistan': {'EMS': 'x', 'e_packet': '', 'e_packet_light': 'x', 'sal': 'x', 'postal_packet': 'x'}...} 'x': No service
def get_ship_service_table():

	Countries = set_spreadsheet_ship_table('Countries')
	Service_data = Countries.get_all_values()
	Service_table = {}
	for i in range(1, len(Service_data)):
		row = Service_data[i]
		Service_table[row[0]] = {
			'EMS':row[1], 'e_packet': row[2], 'e_packet_light': row[3],'small_packet_sal': row[4], 'postal_packet': row[5]}

	return Service_table


###########################################################################
# Shipping service Englist > Japan post service
def get_jppost_service(Service_table, service_name, country_name):

	try:
		if service_name == 'Standard':
			if Service_table[country_name]['e_packet'] != 'x':
				res = 'ｅパケット'
			else:
				res = ''

		elif service_name == 'Expedited':
			if Service_table[country_name]['EMS'] != 'x':
				res = 'EMS'
			elif Service_table[country_name]['postal_packet'] != 'x':
				res = '国際小包'
			else:
				res = ''

		elif service_name == 'Economy':
			if Service_table[country_name]['e_packet_light'] != 'x':
				res = 'ｅパケットライト'
			elif Service_table[country_name]['small_packet_sal'] != 'x':
				res = '小形包装物(書留)'
			else:
				res = ''

		else:
			res = ''

	except:
		res = ''
		print('    Please confirm country: %s' % country_name)

	return res


###########################################################################
if __name__ == '__main__':

	(opts, args) = init_options()
	all_orders = get_all_order(days_ago=3)

	ship_service = get_ship_service(all_orders)
	(check_address, country_dic) = check_address(all_orders)

	write_ship_info(ship_service, check_address, country_dic)

	
	




