from ScriptSharePoint import getAccessToken,get_sites, listSiteFolders, listRootfolders, listPathcontent, downloadFile, transformExcel
import os

accessToken = getAccessToken()
sites = get_sites(accessToken)
libraries = listSiteFolders(accessToken)
root = listRootfolders(accessToken)
newFile = "nuevoExcel.xlsx"


valor = int(input("Elige un método para ejecutar: "))



def F1():
    print('Los sites disponibles en SharePoint son: ')
    for sitios in sites.get("value",[]):
        print(f"-{sitios['name']} -- {sitios['displayName']} -- id: {sitios['id']}")

def F2():
    print(f'Las bibliotecas disponibles en el site son: ')
    for librerias in libraries.get("value",[]):
        print(f"-{librerias['name']} id: {librerias['id']}")

def F3():
    print(f'Las carpetas de raíz son: ')
    for librerias in root.get("value",[]):
        print(f"-{librerias['name']} id: {librerias['id']}")

def F4():
    path = input("Introduce la ruta del directorio el cual quieres que liste: ")
    directory = listPathcontent(accessToken, path)
    print(f'El contenido del directorio {path} es: ')
    for directorio in directory.get("value",[]):
        print(f"-Elemento: {directorio['name']}, id: {directorio['id']}")

def F5():
    downloadFile(accessToken,newFile)
    #Se comprueba si se ha creado el archivo
    if os.path.exists(newFile) and os.path.getsize(newFile) > 0:
        print(f"Archivo Excel creado con éxito: {newFile}")
    else:
        print(f"Advertencia: no se pudo crear o el archivo está vacío: {newFile}")

def F6():
    transformExcel(accessToken,newFile)
    if os.path.exists(newFile) and os.path.getsize(newFile) > 0:
        print(f"Archivo Excel creado con éxito: {newFile}")
    else:
        print(f"Advertencia: no se pudo crear o el archivo está vacío: {newFile}")


if valor == 1:
    F1()
elif valor == 2:
    F2()
elif valor == 3:
    F3()
elif valor == 4:
    F4()
elif valor == 5:
    F5()
elif valor == 6:
    F6()
else:
    print("No hay ningun metodo asociado al numero")