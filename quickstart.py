import pygsheets
import datetime
from watchdog.observers import Observer
from watchdog.events import LoggingEventHandler
import getpass
import logging
from time import sleep
import os
from os.path import splitext
import pytz

#Get username of PC
username = getpass.getuser()

# Folder(s) to monitor || Remove username if required.
dirs_array = [
    ["Example 1","Example1\\to\\path\\directory\\{0}".format(username)],   
    ["Example 2","Example2\\to\\path\\directory\\{0}".format(username)],
    ["Example 3","Example3\\to\\path\\directory\\{0}".format(username)]
]

# Custom DLCs maps, MUST be at start of file name
# Example (GM co 10 example mission v10.Altis)
map_dlcs = ["GM", "SOG", "WS"]

# Google service account authorization file (Json file).
CREDS_FILE_PATH = os.environ.get("CREDS_FILE_PATH")
gc = pygsheets.authorize(service_file=CREDS_FILE_PATH)

# Get and open the spreadsheet
GSHEET_ID = os.environ.get("GSHEET_ID")
spreadsheet = pygsheets.Spreadsheet(client=gc, id=GSHEET_ID)
sh = gc.open(spreadsheet.title)

# Get all worksheets in the spreadsheet
workSheets = []
for sheet in spreadsheet.worksheets(sheet_property=None, value=None, force_fetch=False):
    workSheets.append(sheet.title)

if len(workSheets) < len(dirs_array) and len(dirs_array) > 1:
    # Add new sheets for each directory
    for diretory in dirs_array:
        if not diretory[0] in workSheets:
            new_sheet = sh.add_worksheet(diretory[0], rows=100, cols=26)
            new_sheet.update_row(2,["Status", "Type", "Min", "Max", "Mission Name", "Version", "Island"])
            workSheets.append(diretory[0])

    # If worksheet is new, delete default sheet1
    if workSheets[0] == "Sheet1":
        sh.del_worksheet(sh.worksheet_by_title("Sheet1"))

# Insert new row after row index
def add_row(wksheet, index, data):
    wksheet.insert_rows(index, 1, data, False)
    logging.info(f"New mission added: {data[0:7]}")

def get_all_wksheet_values(wksheet):
    missions = []
    missions = wksheet.get_all_values(include_tailing_empty_rows=False)
    if len(missions) > 1:
        row_index = 1
        temp_hold = []
        count = 0
        for i in missions:
            # Remove any rows without data
            if i == [] or i[4] == "" or i[0] == "Status":
                temp_hold.append(i)
                row_index += 1
                count +=1
            else:
                #Assign row indexes to each
                i.insert(7,row_index)
                row_index += 1
                
        # Remove rows without data from the missions list
        for b in temp_hold:
            if count > 0:
                missions.pop(missions.index(b))   
                count += 1 
            else:
                missions.pop(missions.index(b))
                count += 1

        return missions

def get_matches(data_list, val_to_match):
    matches = [x for x in data_list if val_to_match in x]
    return matches

def get_row_index(data_list, new_player_count, old_player_count, mission_type):
    new_player_count = int(new_player_count)
    old_player_count = int(old_player_count)
    row_index = 0
    last_index = 0
    row_found = False
    last_player_count = 0

    for x in data_list:
        x_player_count = int(x[3])
        x_row = int(x[7])
        if new_player_count >= x_player_count and mission_type == x[1].upper():
            row_index = x_row
            row_found = True
            last_player_count = int(x[3])

        if new_player_count <= x_player_count and mission_type == x[1].upper() and row_found is not True:
            row_found = True
            row_index = x_row-1

        if mission_type == x[1].upper():
            last_index = x_row
        else:
            last_index = x_row

    # Old mission row will be deleted, so adjust row index 
    if old_player_count == 0:
        row_index = row_index+1
    if old_player_count > new_player_count and last_player_count > new_player_count:
        row_index = row_index-1
    elif old_player_count < new_player_count:
        row_index = row_index-1

    # If no row index is set, must be new so set index to last.
    if row_index <= 0:
        row_index = last_index
    return row_index


#Update row in worksheet
def update_row(wksheet, data):
    #Get date
    today = datetime.datetime.now()
    update_date = f"{today.strftime('%d')} {today.strftime('%b')} {today.strftime('%Y')}"
    
    get_version = int(data[3].upper().replace("V", ""))
    data[3] = f"V{get_version}"
    new_file_name = data
    new_file_name[0] = data[0].upper()
    

    # Get wksheet data
    list_data = get_all_wksheet_values(wksheet)
    if list_data is not None:
        # Find if new file name is in list
        matches = get_matches(list_data, new_file_name[2])
        # Check if more than 1 mission found in sheet
        if len(matches) > 1:
            to_remove_matches = []
            # If so, compare map names and remove non matching results
            for x in matches:
                if x[6] != new_file_name[4] or x[1] != new_file_name[0]:
                    to_remove_matches.append(matches.index(x))
                elif x[1] == new_file_name[0] and x[6] != new_file_name[4]:
                    to_remove_matches.append(matches.index(x))

            count = 0
            for b in to_remove_matches:
                if count > 0:
                    matches.pop(b-count)   
                    count += 1 
                else:
                    matches.pop(b)
                    count += 1
    else:
        matches = []
        # If no mission found on sheet, add it. Else update mission row
    if matches == []:
        new_file_name.insert(1, 0)
        new_file_name.insert(0, "NEW")
        if list_data is None or list_data == []:
            row = 3
        else:
            row = get_row_index(list_data, new_file_name[3],0, new_file_name[1])
        add_row(wksheet, row, new_file_name)
    else:
        # Select the matching file name
        wks_file_name = matches[0]
        # set last known mission status
        new_file_name.insert(0,wks_file_name[0])

        # Set the row index
        new_file_name.append(wks_file_name[7])
        #Insert Min player count into new mission info
        new_file_name.insert(2, wks_file_name[2])

        # Sort wksheet version number
        get_version = int(wks_file_name[5].upper().replace("V", ""))
        wks_file_name[5] = f"V{get_version-1}"
        
        # If mission info matches old then update row
        if wks_file_name == new_file_name:
            # Delete old
            wksheet.delete_rows(new_file_name[7], 1)
            new_file_name[0] = "UPDATED"
            wksheet.update_row(get_row_index(list_data, new_file_name[3], new_file_name[1]), new_file_name)
            logging.info(f"Mission on spreadsheet updated: {new_file_name}")

        # Else determine if player change or version changed and update
        else:
            player_change = False

            new_row_index = get_row_index(list_data, new_file_name[3], wks_file_name[3], new_file_name[1])

            # If the player count changed, then update min player count
            if wks_file_name[2] != new_file_name[2] or wks_file_name[3] != new_file_name[3]:
                old_player_count = wks_file_name[3]
                wks_file_name[2] = new_file_name[2]
                wks_file_name[3] = new_file_name[3]
                player_change = True
            
            # If the version changed, update the version
            if wks_file_name[5] != new_file_name[5]:
                wks_file_name[5] = new_file_name[5]

            for i in wks_file_name[8:len(wks_file_name)]:
                new_file_name.append(i)

            # Make one final check to compare file names and update else add new row
            if wks_file_name == new_file_name:
                new_file_name[0] = "UPDATED"
                # Delete old
                if player_change:
                    wksheet.delete_rows(new_file_name[7], 1)
                    new_file_name[11] = update_date
                    new_file_name.pop(7)
                    add_row(wksheet, new_row_index, new_file_name)
                else: 
                    new_file_name[11] = update_date
                    new_file_name.pop(7)
                    wksheet.update_row(new_row_index, new_file_name)
                    logging.info(f"Mission on spreadsheet updated: {new_file_name}")
                
            else:
                new_file_name[0] = "NEW"
                new_file_name.pop(7)
                add_row(wksheet, new_row_index, new_file_name)
                
            

# Delete row index
def delete_row(wksheet, row):
    wksheet.delete_rows(row, 1)

def get_file_info(raw_file_name):
    data = []
    #Get file name and fix user format
    name, mapName = splitext(raw_file_name)

    name = name.translate({ord('_'): " "})
    name = name.translate({ord('@'): ""})
    mapName = mapName.translate({ord('_'): " "})
    mapName = mapName.translate({ord('.'): ""})

    #Create file details in array
    fileDetails = []
    for t in name.split():
        try:
            if not t.upper() in map_dlcs:
                fileDetails.append(t)
        except ValueError:
            pass

    #Extract Mission info
    missionType = fileDetails[0].upper()
    maxPlayers = fileDetails[1]
    missionName = ""
    version = int(fileDetails[len(fileDetails)-1].upper().replace("V", ""))
    version = f"V{version}"

    #Remove extracted data from array
    fileDetails.pop(0)
    fileDetails.pop(0)
    fileDetails.pop(len(fileDetails)-1)
    
    #Combine mission names into single string
    for t in fileDetails:
        missionName = missionName + t + " "
    
    data = [missionType, maxPlayers,missionName,version,mapName]
    return data

class Watcher(LoggingEventHandler):
    def on_created(self, event):
        path, extension = splitext(event.src_path)
        for i in dirs_array:
            if i[1] in str(path):
                diretory = i[1]
                active_worksheet = i[0]
        
        raw_mission_name = path.replace(diretory+"\\", "")

        if extension == ".pbo":
            data = get_file_info(raw_mission_name)
            update_row(sh[workSheets.index(active_worksheet)], data)
        else:
            logging.info(f"New File: Not pbo, not handling file: {raw_mission_name}{extension}")

    def on_deleted(self, event):
        path, extension = splitext(event.src_path)
       
        for i in dirs_array:
            if i[1] in str(path):
                diretory = i[1]
                active_worksheet = i[0]
        raw_mission_name = path.replace(diretory+"\\", "")
        if extension == ".pbo":
            data = get_file_info(raw_mission_name)
            search = sh[workSheets.index(active_worksheet)].find(data[2])
            row = 0
            for i in search:
                search_values = sh[workSheets.index(active_worksheet)].get_row(i.row)
                if search_values[1] == data[0] and search_values[3] == data[1] and search_values[6] == data[4] and search_values[5] == data[3]:
                    row = i.row
            if row > 0:
                delete_row(sh[workSheets.index(active_worksheet)], row)
                logging.info(f"Deleted File: Mission removed from spreadsheet: {raw_mission_name}")
            else:
                logging.info(f"Deleted File: Mission not found in spreadsheet: {raw_mission_name}")
        else:
            logging.info(f"Deleted File: Not pbo, not handling file: {raw_mission_name}{extension}")

# Formatter for logger. Set to London date/time format
class Formatter(logging.Formatter):
    def converter(self, timestamp):
        dt = datetime.datetime.fromtimestamp(timestamp, pytz.timezone('Europe/London'))
        return dt
        
    def formatTime(self, record, datefmt=None):
        dt = self.converter(record.created)
        if datefmt:
            s = dt.strftime(datefmt)
        else:
            try:
                s = dt.isoformat(timespec='milliseconds')
            except TypeError:
                s = dt.isoformat()
        return s


# Observer Start up and watch
if __name__ == "__main__":
    logger = logging.root
    handler = logging.StreamHandler()
    handler.setFormatter(Formatter("%(asctime)s| %(message)s", datefmt="%d-%m-%Y %H:%M"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    event_handler = Watcher()
    observer = Observer()
    for i in dirs_array:
        logging.info(f"Watching: {i}")
        observer.schedule(event_handler, i[1], recursive=True)

    # Start the Observer
    observer.start()
    try:
        while True:
            sleep(10)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


