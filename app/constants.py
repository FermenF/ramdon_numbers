import os

path = "./app/docs/num_random.xlsx";
path_open_route = "app/docs/num_random.xlsx";

def open_file(path):
    option = str(input("¿Desea abrir el documeto automaticamente? S/N: "))
    if(option == "S" or option == "s"):
        try:
            print("\nAbriendo archivo...")
            os.system("start EXCEL.EXE "+ path)
        except Exception as E:
            print(E)