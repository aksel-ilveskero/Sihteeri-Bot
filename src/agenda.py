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
    
    print(penalty_list) # TO-DO: Lisää näille henkilöille sakot


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
        proceedings_text += f"Tarkastetaan kokouksen {meeting_number-1}/2025 pöytäkirja.\n"

    if check_later == True:
        proceedings_text += f"Tämä pöytäkirja tarkastetaan kokouksessa {meeting_number+1}/2025."
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
    meeting_week = current_week + 1
    info_sheet = sheet_service.spreadsheets()
    sheet_result = (
        info_sheet.values()
        .get(spreadsheetId="1nS0NjD0YIfj1OxszkGOaOwHLgRsnkcc3FOExrKlQK5I", range="A:G")
        .execute()
    )
    meeting_info = sheet_result["values"]

    # Get meeting information corresponding to week number
    next_meeting_data = next(data_list for data_list in meeting_info if data_list[0] == str(meeting_week))
    meeting_number = int(next_meeting_data[1])
    
    # Find id of agenda relating to meeting number
    agenda_search = drive_service.files().list(fields="files(id, name)", q=f"name = 'Esityslista {meeting_number}/2025' and trashed = false").execute()
    result = agenda_search.get("files", [])
    
    if len(result) > 1:
        ValueError("Multiple agenda files found!")

    agenda_id = result[0]['id']

    # Get agenda file
    agenda_doc = doc_service.documents().get(documentId=agenda_id).execute()
    
    # Read file contents
    agenda_contents = agenda_doc["body"]["content"]

    distribute_agenda_penalties(agenda_contents, sheet_service)