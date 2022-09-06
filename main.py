from app.generadorPython import generate_numbers
import os
from app.constants import path

def main():
    check = os.path.exists(path)
    if(check == True):
        option = int(input("\n1. Empezar uno nuevo.\n2. Continuar desde el ultimo guardado.\n-> Opción: "))
        if(option == 1):
            generate_numbers()
        if(option == 2):
            from app.methods import checkDataChange
            checkDataChange()
    else:
        generate_numbers()

if __name__ == "__main__":
    main()