def start() -> None:
    print("Tervetuloa KIK:n sihteeribottiin. Valitse toiminto:")
    print("1. Julkaise tyhjä esityslista.\n2. Julkaise täytetty esityslista.\n3. Siivoa kuulumiskierros.")
    
    while True:
        chosen_function = input()
        if chosen_function == "1" or chosen_function == "2" or chosen_function == "3":
            break
        else:
            print("Tapahtui virhe, kokeile uudestaan:")
    
    if chosen_function == "1":
        ...
        #create_agenda()
    elif chosen_function == "2":
        ...
        #publish_agenda()
    elif chosen_function == "3":
        ...
        #clean_proceedings()

if __name__ == "__main__":
    start()