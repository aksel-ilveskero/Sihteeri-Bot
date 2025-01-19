import datetime
import time
import asyncio
import json

from tg_bot import publish_empty_agenda

def cleaning_date(week):
    """Returns start and end date of week in brackets

    Args:
        week (_int_): Week number

    Returns:
        str: Date range
    """
    
    first_day = datetime.datetime.strptime(f'2025-W{int(week)-1}-1', "%Y-W%W-%w").date()
    last_day = first_day + datetime.timedelta(days=6.9)

    return f"({first_day.strftime('%d.%m.')} - {last_day.strftime('%d.%m.')})"
    

def distribute_agenda_penalties(agenda_contents, sheet_service):
    num_lines = len(agenda_contents)

    # Get person lookup table
    person_lookup_table = sheet_service.spreadsheets().values().get(spreadsheetId="1-YJpvi8mxBGKeQyMuEk4b_d6F4vIu_2MvsQuCLr3CKU", range="A:G").execute()

    parse_return = []
    penalty_list = []

    # Find missing information, then add that area to penalty list
    for i in range(num_lines):
        try:
            if "xxx" in agenda_contents[i]["paragraph"]["elements"][0]["textRun"]["content"]:
                position = agenda_contents[i-1]["paragraph"]["elements"][0]["textRun"]["content"]
                parse_return.append(position.rstrip("\n"))
        except:
            pass

    # Add responsible persons to penalty list based on missing area
    for pair in person_lookup_table["values"]:
        if pair[0] in parse_return:
            for name in pair[1:]:
                penalty_list.append(name)
    
    # Turn list to dictionary, with key as name and value as number of penalties
    penalty_dict = {}
    for name in penalty_list:
        if name in penalty_dict:
            penalty_dict[name] += 1
        else:
            penalty_dict[name] = 1
    
    # Get the corresponing row for every person in the penalty sheet
    penalty_sheet_id = "1ESoB5X5K67OLtlVpnSbsiWoYUbUVVe1r9--JRB1_BI8"
    penalty_sheet = sheet_service.spreadsheets()
    sheet_result = (
        penalty_sheet.values()
        .get(spreadsheetId=penalty_sheet_id, range="A:G")
        .execute()
    )
    penalty_info = sheet_result["values"]

    # Ask for confirmation before adding penalties
    check_str = "Seuraavat henkilöt ovat saamassa sakkoja:\n"
    for name, count in penalty_dict.items():
        check_str += f"{name}: {count} sakkoa\n"
    check_str += "Lisätäänkö sakot? (Y/N)\n"
    check_input = input(check_str)
    if check_input == "Y" or check_input == "y":
        pass
    else:
        return None

    # For every name in the penalty list, add their penalty count to the corresponding cell in F column
    print("Lisätään sakot...")
    for name, count in penalty_dict.items():
        # Find index of row where first cell matches name
        row_index = next((i for i, row in enumerate(penalty_info) if row[0] == name), None)
        if row_index is not None:
            # Get current value in B column
            try:
                current_value = int(penalty_info[row_index][1])
            except (IndexError, ValueError):
                current_value = 0
                
            # Update value
            new_value = current_value + count
            
            # Update cell in spreadsheet
            penalty_sheet.values().update(
                spreadsheetId=penalty_sheet_id,
                range=f"B{row_index + 1}",
                valueInputOption="RAW",
                body={"values": [[new_value]]}
            ).execute()
        
    print("Sakot lisätty!")


def distribute_attendance_penalties(in_attendance, inhibited, sheet_service):
    # Get list of all KIKHT25 members
    member_sheet = sheet_service.spreadsheets()
    sheet_result = (
        member_sheet.values()
        .get(spreadsheetId="15Uq97hRNPYEHY04-esWWcTYxnhaUALrFT8IGfmDLX7o", range="A:B")
        .execute()
    )
    member_info = sheet_result["values"]

    # Init list of persos
    penalty_list = []

    # If person is not in attedance or inhibited, add to list
    for member in member_info:
        if member[0] in in_attendance or member[0] in inhibited:
            continue
        else:
            penalty_list.append(member[0])
    
    # Deal appropriate number of penalties
    board_members = ["Valtteri Erkkilä", "Aksel Ilveskero", "Lassi Kortesniemi", "Henri Vilenius", "Eemil Erkkilä", "Anna Passila", "Lauri Vuorjoki", "Nuppu Tikkakoski", "Pessi Fabritius", "Anniina Haka", "Jere Markkinen", "Joonas Keskitalo", "Elia Mäki"]

    penalty_dict = {}
    for name in penalty_list:
        if name in board_members:
            penalty_dict[name] = 6
        else:
            penalty_dict[name] = 2
    
    # Get the corresponing row for every person in the penalty sheet
    penalty_sheet_id = "1ESoB5X5K67OLtlVpnSbsiWoYUbUVVe1r9--JRB1_BI8"
    penalty_sheet = sheet_service.spreadsheets()
    sheet_result = (
        penalty_sheet.values()
        .get(spreadsheetId=penalty_sheet_id, range="A:G")
        .execute()
    )
    penalty_info = sheet_result["values"]

    # Ask for confirmation before adding penalties
    check_str = "Seuraavat henkilöt ovat saamassa sakkoja:\n"
    for name, count in penalty_dict.items():
        check_str += f"{name}: {count} sakkoa\n"
    check_str += "Lisätäänkö sakot? (Y/N)\n"
    check_input = input(check_str)
    if check_input == "Y" or check_input == "y":
        pass
    else:
        return None

    # For every name in the penalty list, add their penalty count to the corresponding cell in F column
    print("Lisätään sakot...")
    for name, count in penalty_dict.items():
        # Find index of row where first cell matches name
        row_index = next((i for i, row in enumerate(penalty_info) if row[0] == name), None)
        if row_index is not None:
            # Get current value in B column
            try:
                current_value = int(penalty_info[row_index][1])
            except (IndexError, ValueError):
                current_value = 0
                
            # Update value
            new_value = current_value + count
            
            # Update cell in spreadsheet
            penalty_sheet.values().update(
                spreadsheetId=penalty_sheet_id,
                range=f"B{row_index + 1}",
                valueInputOption="RAW",
                body={"values": [[new_value]]}
            ).execute()
        
    print("Sakot lisätty!")

def check_present(sheet_service, meeting_number):
    # Load results of form
    inhibited_sheet = sheet_service.spreadsheets()
    sheet_result = (
        inhibited_sheet.values()
        .get(spreadsheetId="1rjk41WD8mjdIXPCfyQkOmqeSGIPO0br2sFAX4Ks3DRQ", range="A:D")
        .execute()
    )
    inhibited_info = sheet_result["values"]

    # Pad all rows to a length of 4
    inhibited_info = [sublist + [""] * (4 - len(sublist)) for sublist in inhibited_info]

    # Add all permanently inhibited to list
    inhibited_persons = []
    permanently_inhibited = [x[1] for x in inhibited_info if x[2] == "Nykyisen lukukauden ajaksi"]
    for person in permanently_inhibited:
        inhibited_persons.append(person)

    # Add all that are inhibited from specific meeting
    number_lookup = [meeting_number, f"0{meeting_number}", f"0{meeting_number}/2025"]
    inhibited_from_meeting = [x[1] for x in inhibited_info if x[3] in number_lookup]
    for person in inhibited_from_meeting:
        inhibited_persons.append(person)
    
    return inhibited_persons


def create_agenda(drive_service, sheet_service, doc_service) -> int:
    """Creates new agenda in drive and publishes it in Telegram

    Args:
        drive_service (_Resource_): Google API service for accessing drive files
        sheet_service (_Resource_): Google API service for accessing spreadsheets
        doc_service (_Resource_): Google API service for accessing docs files

    Raises:
        ValueError: If an agenda with the same name already exists

    Returns:
        result (_int_): Result code
    """

    current_week = datetime.date.today().isocalendar().week
    meeting_week = current_week + 1
    no_cleaning_next_week = False
    
    # Get meeting information from spreadsheet
    info_sheet = sheet_service.spreadsheets()
    sheet_result = (
        info_sheet.values()
        .get(spreadsheetId="1nS0NjD0YIfj1OxszkGOaOwHLgRsnkcc3FOExrKlQK5I", range="A:H")
        .execute()
    )
    meeting_info = sheet_result["values"]

    # Get meeting information corresponding to week number
    next_meeting_data = next(data_list for data_list in meeting_info if data_list[0] == str(meeting_week))
    
    try:
        next_week_cleaner_data = next((data_list for data_list in meeting_info if data_list[0] == str(meeting_week+1)))
    except:
        no_cleaning_next_week = True
    
    meeting_number = int(next_meeting_data[1])
    earlier_meeting_number = meeting_number - 1
    next_meeting_number = meeting_number + 1
    meeting_number = str(meeting_number).zfill(2)
    earlier_meeting_number = str(earlier_meeting_number).zfill(2)
    next_meeting_number = str(next_meeting_number).zfill(2)

    meeting_date = next_meeting_data[2]
    meeting_time = next_meeting_data[3]
    meeting_location = next_meeting_data[4]

    current_cleaner = [next_meeting_data[5], next_meeting_data[6]]
    if len(next_meeting_data) > 7:
        current_cleaner.append(next_meeting_data[7])
    current_week_dates = cleaning_date(meeting_week)

    if no_cleaning_next_week == False:
        next_week_cleaner = [next_week_cleaner_data[5], next_week_cleaner_data[6]]
        if len(next_week_cleaner_data) > 7:
            next_week_cleaner.append(next_week_cleaner_data[7])
        next_week_dates = cleaning_date(meeting_week+1)

    # New agenda filename:
    filename = f"Esityslista {meeting_number}/2025"
    print(f"Luodaan tiedosto {filename}:\n")

    # Check if file already exists:
    duplicate_check_results = (
        drive_service.files()
        .list(fields="files(id, name)", q=f"name = '{filename}' and trashed = false")
        .execute()
    )
    
    if duplicate_check_results["files"] != []:
        raise ValueError("Esityslista on jo olemassa.")

    # Check information
    print(f"Kokouksen viikko: {meeting_week}")
    print(f"Kokouksen päivämäärä ja aika: {meeting_date} klo {meeting_time}")
    print(f"Kokouspaikka: {meeting_location}")
    print(f"Tulevat siivoajat: {' ja '.join(current_cleaner)}")
    if no_cleaning_next_week == False:
        print(f"Sitä seuraavat siivoajat: {' ja '.join(next_week_cleaner)}")
    print()
    info_correct = input("Onko tiedot oikein? (Y/N):\n")

    if info_correct == 'Y' or info_correct == 'y':
        pass
    else:
        return None
    
    # Choose further information
    check_input = input("Tuleeko estyneisyydestä ilmoittaa? (Y/N):\n")
    if check_input == 'Y' or check_input == 'y':
        inhibited = True
    else:
        inhibited = False

    check_input = input("Onko edellinen pöytäkirja tarkastettava? (Y/N):\n")
    if check_input == 'Y' or check_input == 'y':
        check_proceedings = True
    else:
        check_proceedings = False

    check_input = input("Tämä pöytäkirja tarkastetaan seuraavassa kokouksessa? (Y/N):\n")
    if check_input == 'Y' or check_input == 'y':
        check_later = True
    else:
        check_later = False

    
    # Create new agenda file from template
    print("Kopioidaan uusi tiedosto...")
    copy_file_result = (
        drive_service.files()
        .copy(fileId="1860qE1XUQ2pb2m37cTARWX8EEeKzg2OG7xHPxRBn5Tk", supportsAllDrives=True)
        .execute()
    )

    # Rename file to next meeting number
    print("Nimetään tiedosto kokousnumeron mukaisesti...")
    meeting_file_id = copy_file_result["id"]
    name_data = {'name': filename}

    renamed_file_result = (
        drive_service.files()
        .update(fileId=meeting_file_id, body=name_data)
        .execute()
    )

    agenda_id = renamed_file_result["id"]

    # Move file to correct folder
    print("Siirretään tiedosto oikealle kansioon...")
    drive_service.files().update(fileId=agenda_id, addParents="1pFHCklXOvA5_WbVJU-7G939jF1V29Q6M", removeParents="10NEk6LRHKlja7h6AoDAry7ChD1snbU-8", supportsAllDrives=True).execute()

    # Set permissions of agenda file
    print("Asetetaan tiedoston oikeudet...")
    drive_service.permissions().create(fileId=agenda_id, body={"role": "writer", "type": "anyone"}, supportsAllDrives=True).execute()

    # Fill in meeting information:
    print("Täytetään pohja annetuilla tiedoilla...")
    requests = [
        {'replaceAllText': {'replaceText': f"{meeting_number}/2025", "containsText": {"text": "&Kokousnumero&", "matchCase": False}}},
        {'replaceAllText': {'replaceText': f"ti {meeting_date} klo {meeting_time}", "containsText": {"text": "&Aika&", "matchCase": False}}},
        {'replaceAllText': {'replaceText': f"{meeting_location}", "containsText": {"text": "&Paikka&", "matchCase": False}}},
        {'replaceAllText': {'replaceText': f"{meeting_week} {current_week_dates} ovat {current_cleaner[0]} ja {current_cleaner[1]}.", "containsText": {"text": "&Siivousvuoro&", "matchCase": False}}}
    ]
    if no_cleaning_next_week == False:
        requests.append({'replaceAllText': {'replaceText': f"{meeting_week+1} {next_week_dates} ovat {next_week_cleaner[0]} ja {next_week_cleaner[1]}.", "containsText": {"text": "&Tuleva siivousvuoro&", "matchCase": False}}})
    else:
        requests.append({'replaceAllText': {'replaceText': "", "containsText": {"text": "&Tuleva siivousvuoro&", "matchCase": False}}},)

    # Asking further questions
    if inhibited == True:
        requests.append({'replaceAllText': {'replaceText': f"Estyneisyydestä ovat ilmoittaneet …", "containsText": {"text": "&Estyneisyys&", "matchCase": False}}})
    else:
        requests.append({'replaceAllText': {'replaceText': f"Estyneisyydestä ei tarvinnut ilmoittaa.", "containsText": {"text": "&Estyneisyys&", "matchCase": False}}})

    proceedings_text = ""
    if check_proceedings == True:
        proceedings_text += f"Tarkastetaan kokouksen {earlier_meeting_number}/2025 pöytäkirja.\n"

    if check_later == True:
        proceedings_text += f"Tämä pöytäkirja tarkastetaan kokouksessa {next_meeting_number}/2025."
    else:
        proceedings_text += f"Valitaan kokoukselle pöytäkirjan tarkastajat."
    
    requests.append({'replaceAllText': {'replaceText': proceedings_text, "containsText": {"text": "&Pöytäkirjan tarkastus&", "matchCase": False}}})

    # Execute the replace command for the agenda
    document = (
        doc_service.documents()
        .batchUpdate(documentId=agenda_id, body={'requests': requests})
        .execute()
    )

    # Create folder of attahcments
    folder_metadata = {
        "name": f"{meeting_number}/2025",
        "mimeType": "application/vnd.google-apps.folder",
        "parents": ["1Zs-Ivlk7ixXAiFl6E3AlkhDhc5cB6DaD"]
    }
    attachment_folder = (drive_service.files().create(body=folder_metadata, fields="id", supportsAllDrives=True).execute())
    folder_id = attachment_folder.get('id')

    # Ask confirmation before publishing
    check_input = input("Esityslista on nyt valmis. Julkaistaanko esityslista tiedotukseen? (Y/N):\n")
    if check_input == 'Y' or check_input == 'y':
        pass
    else:
        return None

    # Publish the agenda
    print("Julkaistaan pöytäkirja tiedotukseen...")
    document_link = f"https://docs.google.com/document/d/{document['documentId']}"
    asyncio.run(publish_empty_agenda(document_link, folder_id, meeting_number, meeting_date, meeting_time, inhibited))
    
    print("Valmista!")
    return 1


def publish_agenda(drive_service, sheet_service, doc_service) -> int:
    # Get meeting information from spreadsheet
    current_week = datetime.date.today().isocalendar().week
    meeting_week = current_week
    info_sheet = sheet_service.spreadsheets()
    sheet_result = (
        info_sheet.values()
        .get(spreadsheetId="1nS0NjD0YIfj1OxszkGOaOwHLgRsnkcc3FOExrKlQK5I", range="A:H")
        .execute()
    )
    meeting_info = sheet_result["values"]

    # Get meeting information corresponding to week number
    next_meeting_data = next(data_list for data_list in meeting_info if data_list[0] == str(meeting_week))
    meeting_number = str(int(next_meeting_data[1])).zfill(2)
    
    # Find id of agenda relating to meeting number
    agenda_search = drive_service.files().list(fields="files(id, name)", supportsAllDrives=True, corpora="allDrives", includeItemsFromAllDrives=True, q=f"name = 'Esityslista {meeting_number}/2025' and trashed = false").execute()
    result = agenda_search.get("files", [])
    
    if len(result) > 1:
        ValueError("Multiple agenda files found!")

    agenda_id = result[0]['id']

    # Get agenda file
    agenda_doc = doc_service.documents().get(documentId=agenda_id).execute()
    
    # Read file contents
    agenda_contents = agenda_doc["body"]["content"]

    distribute_agenda_penalties(agenda_contents, sheet_service)

    inhibited_list = check_present(sheet_service, meeting_number)

    if len(inhibited_list) > 1:
        names = ", ".join(inhibited_list[:-1]) + " and " + inhibited_list[-1]
    elif inhibited_list:  # Handle a single name
        names = inhibited_list[0]
    else:  # Handle an empty list
        names = ""

    requests = [{'replaceAllText': {'replaceText': names, "containsText": {"text": "Estyneisyydestä ovat ilmoittaneet …", "matchCase": False}}}]

    # Execute the replace command for the agenda
    document = (
        doc_service.documents()
        .batchUpdate(documentId=agenda_id, body={'requests': requests})
        .execute()
    )


def check_proceedings(drive_service, sheet_service, doc_service):
    meeting_number = str(input("\nAnna kokouksen numero: "))
    meeting_number = meeting_number.zfill(2)

    # Find id of agenda relating to meeting number
    proceedings_search = drive_service.files().list(fields="files(id, name)", supportsAllDrives=True, corpora="allDrives", includeItemsFromAllDrives=True, q=f"name = 'Kokous {meeting_number}/2025' and trashed = false").execute()
    result = proceedings_search.get("files", [])
    
    if len(result) > 1:
        ValueError("Multiple proceedings files found!")

    proceedings_id = result[0]['id']

    # Get agenda file
    proceedings_doc = doc_service.documents().get(documentId=proceedings_id).execute()
    
    # Get inhibited persons
    proceedings_contents = proceedings_doc["body"]["content"]
    num_lines = len(proceedings_contents)
    for i in range(num_lines):
        try:
            if "Estyneisyysilmoitukset" in proceedings_contents[i]["paragraph"]["elements"][0]["textRun"]["content"]:
                inhibited_persons = proceedings_contents[i+1]["paragraph"]["elements"][0]["textRun"]["content"]
        except:
            continue
    
    # Get persons in attendance
    in_attendance = []
    for i in range(7,100):
        try:
            if "killan toimihenkilöä" in proceedings_contents[i]["paragraph"]["elements"][0]["textRun"]["content"]: # Reached end
                break
            else:
                paragraph_txt = proceedings_contents[i]["paragraph"]["elements"][0]["textRun"]["content"]
                paragraph_list = paragraph_txt.split("\x0b")
                for line in paragraph_list:
                    if line[0] == "\t":
                        in_attendance.append(line.split("\t")[1].strip())
        except:
            continue
    
    distribute_attendance_penalties(in_attendance, inhibited_persons, sheet_service)