from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from xl import get_num_cols, get_num_rows, copy_row
import datetime

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
