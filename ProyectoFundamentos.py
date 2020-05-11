import xlrd

path = "ProyectoFundamentos.xlsx"
inputArchivo = xlrd.open_workbook(path)
inputPrincipal = inputArchivo.sheet_by_index(0)
numFilas = inputPrincipal.nrows


def menu():
    global opc

    print("1. Listado de Materias\n2. Secciones de una Materia\n3. Listado de Profesores\n4. Materias de un Profesor\n5. Salir\n")
    opc = input("Selecciona una opcion: ")


def listadoMaterias():
    for i in range(numFilas):
        while inputPrincipal.cell_value(i, 0) != "":
            print(inputPrincipal.cell_value(i, 0).ljust(15, " "), "|", inputPrincipal.cell_value(i, 1))
            i += 1
    print("\n")


def seccionesMaterias():
    id_Materia = input("Introduce el ID de Materia: ")

    for i in range(numFilas):
        if inputPrincipal.cell_value(i, 0) == id_Materia:
            print(inputPrincipal.cell_value(0, 2).ljust(15, " "), "|", inputPrincipal.cell_value(0, 4).ljust(20, " "),
            "|", inputPrincipal.cell_value(0, 5))
            while inputPrincipal.cell_value(i, 0) == id_Materia or inputPrincipal.cell_value(i, 0) == "":
                print(inputPrincipal.cell_value(i, 2).ljust(15, " "), "|", inputPrincipal.cell_value(i, 4).ljust(20, " "),
                "|", inputPrincipal.cell_value(i, 5))
                i += 1
                if i >= numFilas:
                    break
        i += 1
    print("\n")


def listadoProfesores():
    for i in range(numFilas):
        if inputPrincipal.cell_value(i, 7) != "":
            print(inputPrincipal.cell_value(i, 7).ljust(15, " "), "|", inputPrincipal.cell_value(i, 8))
            i += 1
    print("\n")


def seccionesProfesores():
    validacion = 0
    id_Profesor = input("Introduce el ID de Profesor: ")

    for i in range(numFilas):
        if inputPrincipal.cell_value(i, 3) == id_Profesor:
            if validacion == 0:
                print(inputPrincipal.cell_value(0, 2).ljust(15, " "), "|", inputPrincipal.cell_value(0, 5))
                validacion = 1
            print(inputPrincipal.cell_value(i, 2).ljust(15, " "), "|", inputPrincipal.cell_value(i, 5))
        i += 1
    print("\n")


menu()

while opc != 1 or opc != 2 or opc != 3 or opc != 4:
    if opc == "1":
        listadoMaterias()

    elif opc == "2":
        seccionesMaterias()

    elif opc == "3":
        listadoProfesores()

    elif opc == "4":
        seccionesProfesores()

    elif opc == "5":
        break

    else:
        print("Opcion no valida...\n")

    menu()
