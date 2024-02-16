from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from xl import get_num_cols, get_num_rows, copy_row
import datetime


hotspot = {'trackResponse': {'shipment': [{'inquiryNumber': '1Z3124Y80339645436', 'package': [{'trackingNumber': '1Z3124Y80339645436', 'deliveryDate': [], 'activity': [{'location': {'address': {'countryCode': 'US', 'country': 'US'}}, 'status': {'type': 'M', 'description': 'Shipper created a label, UPS has not received the package yet. ', 'code': 'MP', 'statusCode': '003'}, 'date': '20240209', 'time': '075350', 'gmtDate': '20240209', 'gmtOffset': '-08:00', 'gmtTime': '15:53:50'}], 'currentStatus': {'description': 'Shipment Ready for UPS', 'code': '003'}, 'packageAddress': [{'type': 'ORIGIN', 'name': 'ABS KIDS', 'attentionName': '', 'address': {'addressLine1': ' 16255 VENTURA BLVD ', 'addressLine2': '', 'city': 'ENCINO', 'stateProvince': 'CA', 'postalCode': '914362317', 'countryCode': 'US', 'country': 'US'}}, {'type': 'DESTINATION', 'name': 'ABS KIDS', 'attentionName': 'MILES FORET', 'address': {'addressLine1': '258 EAST GARRISON BLVD', 'addressLine2': '', 'city': 'GASTONIA', 'stateProvince': 'NC', 'postalCode': '28054', 'countryCode': 'US', 'country': 'US'}}], 'weight': {'unitOfMeasurement': 'LBS', 'weight': '5.00'}, 'service': {'code': '518', 'levelCode': '003', 'description': 'UPS Ground'}, 'referenceNumber': [{'type': 'SHIPMENT', 'number': 'HOTSPOT RETURN'}, {'type': 'SHIPMENT', 'number': 'BILL TO IT'}, {'type': 'PACKAGE', 'number': 'HOTSPOT RETURN'}, {'type': 'PACKAGE', 'number': 'BILL TO IT'}], 'deliveryInformation': {'deliveryPhoto': {'isNonPostalCodeCountry': False}}, 'dimension': {'height': '6.00', 'length': '20.00', 'width': '10.00', 'unitOfDimension': 'IN'}, 'packageCount': 1}], 'userRelation': ['SHIPPER']}]}}
dalton = {'trackResponse': {'shipment': [{'inquiryNumber': '1Z3124Y81211512775', 'package': [{'trackingNumber': '1Z3124Y81211512775', 'deliveryDate': [{'type': 'SDD', 'date': '20240216'}], 'deliveryTime': {'startTime': '140000', 'type': 'EDW', 'endTime': '160000'}, 'activity': [{'location': {'address': {'city': 'Greensboro', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2749'}, 'status': {'type': 'I', 'description': 'Departed from Facility', 'code': 'DP', 'statusCode': '010'}, 'date': '20240214', 'time': '010100', 'gmtDate': '20240214', 'gmtOffset': '-05:00', 'gmtTime': '06:01:00'}, {'location': {'address': {'city': 'Greensboro', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2802'}, 'status': {'type': 'I', 'description': 'Arrived at Facility', 'code': 'AR', 'statusCode': '005'}, 'date': '20240213', 'time': '231100', 'gmtDate': '20240214', 'gmtOffset': '-05:00', 'gmtTime': '04:11:00'}, {'location': {'address': {'city': 'Gastonia', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2802'}, 'status': {'type': 'I', 'description': 'Departed from Facility', 'code': 'DP', 'statusCode': '005'}, 'date': '20240213', 'time': '210900', 'gmtDate': '20240214', 'gmtOffset': '-05:00', 'gmtTime': '02:09:00'}, {'location': {'address': {'city': 'Gastonia', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2802'}, 'status': {'type': 'I', 'description': 'We have your package', 'code': 'OR', 'statusCode': '005'}, 'date': '20240213', 'time': '182500', 'gmtDate': '20240213', 'gmtOffset': '-05:00', 'gmtTime': '23:25:00'}, {'location': {'address': {'countryCode': 'US', 'country': 'US'}}, 'status': {'type': 'M', 'description': 'Shipper created a label, UPS has not received the package yet. ', 'code': 'MP', 'statusCode': '003'}, 'date': '20240213', 'time': '063909', 'gmtDate': '20240213', 'gmtOffset': '-08:00', 'gmtTime': 
'14:39:09'}], 'currentStatus': {'description': 'On the Way', 'code': '010'}, 'packageAddress': [{'type': 'ORIGIN', 'name': 'ABS KIDS', 'attentionName': '', 'address': {'addressLine1': ' 16255 VENTURA BLVD ', 'addressLine2': '', 'city': 'ENCINO', 'stateProvince': 'CA', 'postalCode': '914362317', 'countryCode': 'US', 'country': 'US'}}, {'type': 'DESTINATION', 'name': 'DALTON GRANGE', 'attentionName': '', 'address': {'addressLine1': '2417 WEST 4975 SOUTH', 'addressLine2': '', 'city': 'ROY', 'stateProvince': 'UT', 'postalCode': '84067', 'countryCode': 'US', 'country': 'US'}}], 'weight': {'unitOfMeasurement': 'LBS', 'weight': '4.00'}, 'service': {'code': '546', 'levelCode': '012', 'description': 'UPS 3 Day Select®'}, 'referenceNumber': [{'type': 'SHIPMENT', 'number': 'DALTON GRANGE LAPTOP'}, {'type': 'SHIPMENT', 'number': 'BILL TO IT'}, {'type': 'PACKAGE', 'number': 'DALTON GRANGE LAPTOP'}, {'type': 'PACKAGE', 'number': 'BILL TO IT'}], 'deliveryInformation': {'deliveryPhoto': {'isNonPostalCodeCountry': False}}, 'dimension': {'height': '3.00', 'length': '17.00', 'width': '13.00', 'unitOfDimension': 'IN'}, 'packageCount': 2}], 'userRelation': ['SHIPPER']}]}}
fmaudit = {'trackResponse': {'shipment': [{'inquiryNumber': '1Z3124Y80203095153', 'package': [{'trackingNumber': '1Z3124Y80203095153', 'deliveryDate': [{'type': 'DEL', 'date': '20240212'}], 'deliveryTime': {'type': 'DEL', 
'endTime': '152415'}, 'activity': [{'location': {'address': {'city': 'SALT LAKE CITY', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '8410'}, 'status': {'type': 'D', 'description': 'DELIVERED ', 'code': '9E', 'statusCode': '011'}, 'date': '20240212', 'time': '152415', 'gmtDate': '20240212', 'gmtOffset': '-07:00', 'gmtTime': '22:24:15'}, {'location': {'address': {'city': 'Salt Lake City', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '8410'}, 'status': {'type': 'I', 'description': 'Out For Delivery Today', 'code': 'OT', 'statusCode': '021'}, 'date': '20240212', 'time': '092113', 'gmtDate': '20240212', 'gmtOffset': '-07:00', 'gmtTime': '16:21:13'}, {'location': {'address': {'city': 'Salt Lake City', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '8419'}, 'status': {'type': 'I', 'description': 'Processing at UPS Facility', 'code': 'DS', 'statusCode': '087'}, 'date': '20240210', 'time': '063008', 'gmtDate': '20240210', 'gmtOffset': '-07:00', 'gmtTime': '13:30:08'}, {'location': 
{'address': {'city': 'Salt Lake City', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '8419'}, 'status': {'type': 'X', 'description': 'The receiving business was closed and delivery has been rescheduled for the next business day.', 'code': '55', 'statusCode': '055'}, 'date': '20240209', 'time': '091216', 'gmtDate': '20240209', 'gmtOffset': '-07:00', 'gmtTime': '16:12:16'}, {'location': {'address': 
{'city': 'Salt Lake City', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '8419'}, 'status': {'type': 'X', 'description': 'The receiving business was closed and delivery has been rescheduled for the next business day.', 'code': '55', 'statusCode': '055'}, 'date': '20240209', 'time': '070917', 'gmtDate': '20240209', 'gmtOffset': '-07:00', 'gmtTime': '14:09:17'}, {'location': {'address': {'city': 'Salt Lake City', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '8419'}, 'status': {'type': 'I', 'description': 'Processing at UPS Facility', 'code': 'YP', 'statusCode': '071'}, 'date': '20240209', 'time': '055340', 'gmtDate': '20240209', 'gmtOffset': '-07:00', 'gmtTime': '12:53:40'}, {'location': {'address': {'city': 'Salt Lake City', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '4009'}, 'status': {'type': 'I', 'description': 'Arrived at Facility', 'code': 'AR', 'statusCode': '005'}, 'date': '20240208', 'time': '215900', 'gmtDate': '20240209', 'gmtOffset': '-07:00', 'gmtTime': '04:59:00'}, {'location': {'address': {'city': 'Salt Lake City', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '4009'}, 'status': {'type': 'I', 'description': 'Departed from Facility', 'code': 
'DP', 'statusCode': '005'}, 'date': '20240208', 'time': '202200', 'gmtDate': '20240209', 'gmtOffset': '-07:00', 'gmtTime': '03:22:00'}, {'location': {'address': {'city': 'Salt Lake City', 'stateProvince': 'UT', 'countryCode': 'US', 'country': 'US'}, 'slic': '4009'}, 'status': {'type': 'I', 'description': 'Arrived at Facility', 'code': 'AR', 'statusCode': '005'}, 'date': '20240208', 'time': '181800', 'gmtDate': '20240209', 'gmtOffset': '-07:00', 'gmtTime': '01:18:00'}, {'location': {'address': {'city': 'Louisville', 'stateProvince': 'KY', 'countryCode': 'US', 'country': 'US'}, 'slic': '4009'}, 'status': {'type': 'I', 'description': 'Departed from Facility', 'code': 'DP', 'statusCode': '005'}, 'date': '20240208', 'time': '164300', 'gmtDate': '20240208', 'gmtOffset': '-05:00', 'gmtTime': '21:43:00'}, {'location': {'address': {'city': 'Louisville', 'stateProvince': 'KY', 'countryCode': 'US', 'country': 'US'}, 'slic': '2749'}, 'status': {'type': 'I', 'description': 'Arrived at Facility', 'code': 'AR', 'statusCode': '005'}, 'date': '20240208', 'time': '120300', 'gmtDate': '20240208', 'gmtOffset': '-05:00', 'gmtTime': '17:03:00'}, {'location': {'address': {'city': 'South Charleston', 'stateProvince': 'WV', 'countryCode': 'US', 'country': 'US'}, 'slic': '2749'}, 'status': {'type': 'I', 'description': 'Departed from Facility', 'code': 'DP', 'statusCode': '005'}, 'date': '20240208', 'time': '074500', 'gmtDate': '20240208', 'gmtOffset': '-05:00', 'gmtTime': '12:45:00'}, 
{'location': {'address': {'city': 'South Charleston', 'stateProvince': 'WV', 'countryCode': 'US', 'country': 'US'}, 'slic': '2749'}, 'status': {'type': 'I', 'description': 'Arrived at Facility', 'code': 'AR', 'statusCode': '005'}, 'date': '20240208', 'time': '071700', 'gmtDate': '20240208', 'gmtOffset': '-05:00', 'gmtTime': '12:17:00'}, {'location': {'address': {'city': 'Greensboro', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2749'}, 'status': {'type': 'I', 'description': 'Departed from Facility', 'code': 'DP', 'statusCode': '005'}, 'date': '20240208', 'time': '030200', 'gmtDate': '20240208', 'gmtOffset': '-05:00', 'gmtTime': '08:02:00'}, {'location': {'address': {'city': 'Greensboro', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2802'}, 'status': {'type': 'I', 'description': 'Arrived at Facility', 'code': 'AR', 'statusCode': '005'}, 'date': '20240207', 'time': '232800', 'gmtDate': '20240208', 'gmtOffset': '-05:00', 'gmtTime': '04:28:00'}, {'location': {'address': {'city': 'Gastonia', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2802'}, 'status': {'type': 'I', 'description': 'Departed from Facility', 'code': 'DP', 'statusCode': '005'}, 'date': '20240207', 'time': '211900', 
'gmtDate': '20240208', 'gmtOffset': '-05:00', 'gmtTime': '02:19:00'}, {'location': {'address': {'city': 'Gastonia', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2802'}, 'status': {'type': 'I', 'description': 'We have your package', 'code': 'OR', 'statusCode': '005'}, 'date': '20240207', 'time': '182643', 'gmtDate': '20240207', 'gmtOffset': '-05:00', 'gmtTime': '23:26:43'}, {'location': {'address': {'city': 'Gastonia', 'stateProvince': 'NC', 'countryCode': 'US', 'country': 'US'}, 'slic': '2802'}, 'status': {'type': 'I', 'description': 'Drop-Off', 'code': 'XD', 'statusCode': '005'}, 'date': '20240207', 'time': '133800', 'gmtDate': '20240207', 'gmtOffset': '-05:00', 'gmtTime': '18:38:00'}, {'location': {'address': {'countryCode': 'US', 'country': 'US'}}, 'status': {'type': 'M', 'description': 'Shipper created a label, UPS has not received the package yet. ', 'code': 'MP', 'statusCode': '003'}, 'date': '20240207', 'time': '101646', 'gmtDate': '20240207', 'gmtOffset': '-08:00', 'gmtTime': '18:16:46'}], 'currentStatus': {'description': 'Delivered', 'code': '011'}, 'packageAddress': [{'type': 'ORIGIN', 'name': 'ABS KIDS', 'attentionName': '', 'address': {'addressLine1': ' 16255 VENTURA BLVD ', 'addressLine2': '', 'city': 'ENCINO', 'stateProvince': 'CA', 'postalCode': '914362317', 'countryCode': 'US', 'country': 'US'}}, {'type': 'DESTINATION', 'name': '', 'attentionName': '', 'address': {'addressLine1': '515 S 700 E', 'addressLine2': '', 'city': 'SALT LAKE CITY', 'stateProvince': 'UT', 'postalCode': '84102', 'countryCode': 'US', 'country': 'US'}}], 'weight': {'unitOfMeasurement': 'LBS', 'weight': '5.00'}, 'service': {'code': '545', 'levelCode': '002', 'description': 'UPS 2nd Day Air®'}, 'referenceNumber': [{'type': 'SHIPMENT', 'number': 'FMAUDIT'}, {'type': 'SHIPMENT', 'number': 'BILL TO IT'}, {'type': 'PACKAGE', 'number': 'FMAUDIT'}, {'type': 'PACKAGE', 'number': 'BILL TO IT'}], 'deliveryInformation': {'receivedBy': 'ROMAN', 'location': 'Front Desk', 'deliveryPhoto': {'photoCaptureInd': 'N', 'isNonPostalCodeCountry': False}}, 'dimension': {'height': '6.00', 'length': '20.00', 'width': '20.00', 'unitOfDimension': 'IN'}, 'packageCount': 1}], 'userRelation': ['SHIPPER']}]}}



def trackPackage(api_data):
    #CURRENT STATUS OF PACKAGE
    curr_status = api_data['trackResponse']['shipment'][0]['package'][0]['currentStatus']['description']

    obj = {
        'label_created':'',
        'status': curr_status,
        'shipped_date': '',
        'last_scan_time': '',
        'last_scan_location': '',
        'shipping_to': '',
        'delivery_date': '',
        'reference': '',
        'bill': '',
        'ship_service': ''
    }

    obj['reference'] = api_data["trackResponse"]["shipment"][0]['package'][0]['referenceNumber'][0]['number']
    obj['bill'] = api_data["trackResponse"]["shipment"][0]['package'][0]['referenceNumber'][1]['number']
    obj['ship_service'] = api_data["trackResponse"]["shipment"][0]['package'][0]['service']['description']

    #LABEL CREATED DATE
    label_date = api_data["trackResponse"]["shipment"][0]['package'][0]['activity']
    label_date = api_data["trackResponse"]["shipment"][0]['package'][0]['activity'][len(label_date)-1]['date']
    ld_formatted = label_date[4:]+label_date[:4]
    ld_formatted = ld_formatted[:2]+'/'+ld_formatted[2:4]+'/'+ld_formatted[4:]
    obj['label_created'] = ld_formatted

    #SHIPPING TO
    shipping_to = api_data["trackResponse"]["shipment"][0]['package'][0]['packageAddress'][1]['address']
    obj['shipping_to'] = shipping_to['addressLine1'] + ' '+ shipping_to['city']+ ', ' +shipping_to['stateProvince']+' '+shipping_to['postalCode'] 

    #IF STATUS IS UPS WAITING FOR PACKAGE
    if(curr_status == 'Shipment Ready for UPS'):
        obj['shipped_date']='N/A'
        obj['last_scan_time']='N/A'
        obj['last_scan_location'] = 'N/A'
        obj['delivery_date'] = 'N/A'

    #IF PACKAGE IS IN TRANSIT
    elif(curr_status == 'On the Way' or 'Preparing for Delivery'):

        #SHIPDATE
        ship_date = api_data["trackResponse"]["shipment"][0]['package'][0]['activity']
        ship_date = api_data["trackResponse"]["shipment"][0]['package'][0]['activity'][len(ship_date)-2]['date']
        sd_formatted = ship_date[4:]+ship_date[:4]
        sd_formatted = sd_formatted[:2]+'/'+sd_formatted[2:4]+'/'+sd_formatted[4:]
        obj['shipped_date']=sd_formatted
        

        #LAST SCAN TIME/DATE
        last_scan_time = api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]['time']
        last_scan_time = last_scan_time[:2]+':' +last_scan_time[2:4]
        last_scan_time = datetime.datetime.strptime(last_scan_time,'%H:%M').strftime('%I:%M %p')

        last_scan_date = api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]['date']
        last_scan_date = last_scan_date[4:]+last_scan_date[:4]
        last_scan_date = last_scan_date[:2]+'/'+last_scan_date[2:4]+'/'+last_scan_date[4:]
        obj['last_scan_time'] = last_scan_date+' - '+last_scan_time

        #LAST SCAN LOCATION
        location = api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]["location"]["address"]["city"] + ", " + api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]["location"]["address"]["stateProvince"]
        status = api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]["status"]["description"] + f" ({location})" 
        obj['last_scan_location'] = status
        obj['delivery_date'] = 'N/A'
        
        

    #IF PACKAGE IS DELIVERED
    elif(curr_status == 'Delivered'):

        #SHIPDATE
        ship_date = api_data["trackResponse"]["shipment"][0]['package'][0]['activity']
        ship_date = api_data["trackResponse"]["shipment"][0]['package'][0]['activity'][len(ship_date)-2]['date']
        sd_formatted = ship_date[4:]+ship_date[:4]
        sd_formatted = sd_formatted[:2]+'/'+sd_formatted[2:4]+'/'+sd_formatted[4:]
        obj['shipped_date']=sd_formatted


        #LAST SCAN TIME/DATE
        last_scan_time = api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]['time']
        last_scan_time = last_scan_time[:2]+':' +last_scan_time[2:4]
        last_scan_time = datetime.datetime.strptime(last_scan_time,'%H:%M').strftime('%I:%M %p')

        last_scan_date = api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]['date']
        last_scan_date = last_scan_date[4:]+last_scan_date[:4]
        last_scan_date = last_scan_date[:2]+'/'+last_scan_date[2:4]+'/'+last_scan_date[4:]
        obj['last_scan_time'] = last_scan_date+' - '+last_scan_time



        obj['last_scan_time']=last_scan_time 

        #LAST SCAN LOCATION
        location = api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]["location"]["address"]["city"] + ", " + api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]["location"]["address"]["stateProvince"]
        status = api_data["trackResponse"]["shipment"][0]['package'][0]["activity"][0]["status"]["description"] + f"({location})" 
        obj['last_scan_location'] = status

        #DELIVERY DATE/FORMATTING
        delivery_date = api_data["trackResponse"]["shipment"][0]['package'][0]['deliveryDate'][0]['date']
        dd_formatted = delivery_date[4:]+delivery_date[:4]
        dd_formatted = dd_formatted[:2]+'/'+dd_formatted[2:4]+'/'+dd_formatted[4:]
        obj['delivery_date'] = dd_formatted

    #else
    else:
        print(curr_status)

    return obj
