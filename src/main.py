from drive_utils import login
from agenda import create_agenda, publish_agenda, check_proceedings

def start() -> None:
        
    drive_service, sheet_service, doc_service = login()
    
    print("Tervetuloa KIK:n sihteeribottiin.")
    print("1. Julkaise tyhjä esityslista.\n2. Julkaise valmis esityslista.\n3. Käy pöytäkirja läpi\n4. Siivoa kuulumiskierros.\n")
    
    while True:
        chosen_function = input("Valitse toiminto: ")
        if chosen_function == "1" or chosen_function == "2" or chosen_function == "3":
            break
        else:
            print("Tapahtui virhe, kokeile uudestaan:")

    if chosen_function == "1":
        create_agenda(drive_service, sheet_service, doc_service)
        
    elif chosen_function == "2":
        publish_agenda(drive_service, sheet_service, doc_service)

    elif chosen_function == "3":
        check_proceedings(drive_service, sheet_service, doc_service)

if __name__ == "__main__":
    start()