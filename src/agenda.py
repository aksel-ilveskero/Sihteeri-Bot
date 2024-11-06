import datetime
import time
import asyncio

from tg_bot import publish_empty_agenda

def cleaning_date(week):
    """Returns start and end date of week in brackets

    Args:
        week (_int_): Week number

    Returns:
        str: Date range
    """
    
    first_day = datetime.datetime.strptime(f'2025-W{int(week)-1}-2', "%Y-W%W-%w").date()
    last_day = first_day + datetime.timedelta(days=6.9)

    return f"({first_day.strftime("%d.%m.")} - {last_day.strftime("%d.%m.")})"
    

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
    
    # Get meeting information from spreadsheet
    info_sheet = sheet_service.spreadsheets()
    sheet_result = (
        info_sheet.values()
        .get(spreadsheetId="1nS0NjD0YIfj1OxszkGOaOwHLgRsnkcc3FOExrKlQK5I", range="A:G")
        .execute()
    )
    meeting_info = sheet_result["values"]

    # Get meeting information corresponding to week number
    next_meeting_data = next(data_list for data_list in meeting_info if data_list[0] == str(meeting_week))
    next_week_cleaner_data = next(data_list for data_list in meeting_info if data_list[0] == str(meeting_week+1))
    
    meeting_number = int(next_meeting_data[1])
    meeting_date = next_meeting_data[2]
    meeting_time = next_meeting_data[3]
    meeting_location = next_meeting_data[4]

    current_cleaner = [next_meeting_data[5], next_meeting_data[6]]
    current_week_dates = cleaning_date(meeting_week)

    next_week_cleaner = [next_week_cleaner_data[5], next_week_cleaner_data[6]]
    next_week_dates = cleaning_date(meeting_week+1)

    # New agenda filename:
    #filename = f"Esityslista {meeting_number}/2025"
    filename = f"Esityslista {time.time()}/2025"
    
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
    print(f"Tulevat siivoajat: {current_cleaner[0]} ja {current_cleaner[1]}")
    print(f"Sitä seuraavat siivoajat: {next_week_cleaner[0]} ja {next_week_cleaner[1]}")
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

    # Fill in meeting information:
    print("Täytetään pohja annetuilla tiedoilla...")
    requests = [
        {'replaceAllText': {'replaceText': f"{meeting_number}/2025", "containsText": {"text": "&Kokousnumero&", "matchCase": False}}},
        {'replaceAllText': {'replaceText': f"ti {meeting_date} klo {meeting_time}", "containsText": {"text": "&Aika&", "matchCase": False}}},
        {'replaceAllText': {'replaceText': f"{meeting_location}", "containsText": {"text": "&Paikka&", "matchCase": False}}},
        {'replaceAllText': {'replaceText': f"{meeting_week} {current_week_dates} ovat {current_cleaner[0]} ja {current_cleaner[1]}.", "containsText": {"text": "&Siivousvuoro&", "matchCase": False}}},
        {'replaceAllText': {'replaceText': f"{meeting_week+1} {next_week_dates} ovat {next_week_cleaner[0]} ja {next_week_cleaner[1]}.", "containsText": {"text": "&Tuleva siivousvuoro&", "matchCase": False}}},
    ]

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

    # Publish the agenda
    print("Julkaistaan pöytäkirja tiedotukseen...")
    document_link = f"https://docs.google.com/document/d/{document["documentId"]}"
    asyncio.run(publish_empty_agenda(document_link, None, meeting_number, meeting_date, meeting_time))
    
    print("Valmista!")
    return 1