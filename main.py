import configparser
import requests
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
import calendar
import json
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
import pandas as pd
import io

def get_site_status(config): 
    url = config['DEFAULT']['VerkadaURL']    

    headers = {
        "accept": "application/json",
        "x-api-key": config['DEFAULT']['VerkadaAPIKey']
    }

    response = requests.get(url, headers=headers)
    return json.loads(response.text)["sites"]
    

def get_schedule_file(config):
    authcookie = Office365('https://mywirelessgroup.sharepoint.com', 
                           username=config['DEFAULT']['BotUsername'], 
                           password=config['DEFAULT']['BotPassword']).GetCookies()
    site = Site('https://mywirelessgroup.sharepoint.com/sites/GMETeam/', version=Version.v365, authcookie=authcookie)

    folder = site.Folder('Shared Documents/General/Operations Dashboard/Store Hours of Operation - AP')

    schedule_file = folder.get_file('Store Schedules.xlsx')

    df = pd.read_excel(io.BytesIO(schedule_file))

    return df[df['End Date'].isnull()]


def market_to_timezone(marketname):
    match marketname:
        case "AZPHX Market": 
            return 'US/Arizona' 
        case "CABAY Market": 
            return 'US/Pacific'
        case "CAGLA Market": 
            return 'US/Pacific'
        case "CASAN Market": 
            return 'US/Pacific'
        case "ILCHI Market": 
            return 'US/Central'
        case "MIDET Market": 
            return 'US/Eastern'
        case "ININD Market": 
            return 'US/Eastern'
        case "NVLAS Market": 
            return 'US/Pacific'
        case "ORPTL Market": 
            return 'US/Pacific'
        case "WASEA Market": 
            return 'US/Pacific'
        case _: 
            return ''
        

def get_cinglepointid(sitename):
    try: 
        cpid = int(sitename.split('~')[1].strip())
    except IndexError: 
        cpid = -1
    except ValueError: 
        cpid = -1
    return cpid


def get_open_close_columns(validation_day):
    match validation_day: 
        case "Monday": 
            return 4, 5
        case "Tuesday": 
            return 6, 7
        case "Wednesday": 
            return 8, 9
        case "Thursday": 
            return 10, 11
        case "Friday": 
            return 12, 13
        case "Saturday": 
            return 14, 15
        case "Sunday": 
            return 16, 17 



def site_validation(verkadafile, schedulefile, validation_time, validation_day):
    skipped_locations = []
    for site in verkadafile:
        cpid = get_cinglepointid(site["site_name"])

        if cpid == -1:
            skipped_locations.append(site["site_name"])
            continue

        site_state = site["site_state"]

        siterow = schedulefile[schedulefile['Cinglepoint ID'] == cpid]
        market_name = siterow.iat[0,0]
        site_timezone = market_to_timezone(market_name)

        if site_timezone == '':
            skipped_locations.append(site["site_name"])
            continue

        open_column, close_column = get_open_close_columns(validation_day)
        
        try: 
            open_time = datetime.strptime(siterow.iat[0, open_column], "%I:%M %p").time()
            close_time = datetime.strptime(siterow.iat[0,close_column], "%I:%M %p").time()
        except ValueError:
            skipped_locations.append(site["site_name"]) 
            continue
        except TypeError: 
            open_time = siterow.iat[0, open_column]
            close_time = siterow.iat[0, close_column]
        
        site_local_time = validation_time.astimezone(ZoneInfo(site_timezone)).time()

        print(f'{site["site_name"]} cpid {cpid}, site_state {site_state}, open time {open_time}, close time {close_time}, tz = {site_timezone}')
        print(f'validation time {validation_time.time()}')
        print(f'site local time {site_local_time}')
        print(f'send alert {site_local_time > open_time and site_local_time < close_time and site_state == 'armed'}')
    
    print(skipped_locations)
    
    

if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read('config.ini')

    # Get current time
    my_time = datetime.now(ZoneInfo('US/Pacific'))
    # Get weekday
    my_weekday = calendar.day_name[datetime.now().weekday()]
    
    print(f"Time is {my_time}. Weekday is {my_weekday}")

    site_list = get_site_status(config)
    print(site_list)
    schedule_file = get_schedule_file(config)
    site_validation(site_list, schedule_file, my_time, my_weekday)

