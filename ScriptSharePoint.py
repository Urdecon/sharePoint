import os,requests,io
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

TENANT_ID = os.getenv('DB_TENANT_ID')
CLIENT_ID = os.getenv('DB_CLIENT_ID')
CLIENT_SECRET = os.getenv('DB_CLIENT_SECRET')
SITE_ID = os.getenv('INF_SITE_ID')
DRIVE_ID = os.getenv('INF_DRIVE_ID')
ITEM_ID = os.getenv('ITEM_ID')
SCOPE = 'https://graph.microsoft.com/.default'

##Mismo metodo que para Business Central##

def getAccessToken():
    url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'

    headers = {
        'Content-type' :'application/x-www-form-urlencoded'
    }

    data = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': SCOPE
    }

    response = requests.post(url, headers=headers, data=data)
    response.raise_for_status()
    return response.json()['access_token']

# Metodo para obtener los sites, Sharepoint es de urdecon.es por lo que necesitamos los datos de la app creada para DynBuilder
# Lo que queremos de aqui es el site_id, compuesto por id: urdecon.sharepoint.com, xx-xx-xx-xx-xx,xx-xx-xx-xx-xx

def get_sites (access_token):
    url = f'https://graph.microsoft.com/v1.0/sites?search=*'
    headers = {
        'Authorization' : f'Bearer {access_token}',
        'Accept':'application/json'
    }
    response = requests.get(url,headers = headers)
    response.raise_for_status()
    return response.json()

# Con este site id ya podemos ir listando bibliotecas, por debajo del site  principal,
# tenemos subcarpetas, donde ya podremos navegar por los archivos, Estudios/Documentos/General/Licitaciones...... Informatica/Documentos/General/IA....
# Deberemos comprobar primero donde queremos acceder ya que una vez obtenido el site_id, obtendremos el drive_id (identificador de la carpeta raiz) (id de la biblioteca principal suele ser Documentos)
# Y con este drive_id realizamos las consultas a los directorios que están por debajo del raiz
#Por lo general debe ser Estudios/Documentos/General/   Informatica/Documentos/General.... he probado con los que mi cuenta tiene acceso para así comprobar.

#El drive_id corresponde a toda la biblioteca, no es específico de una subcarpeta, estudios tendrá uno, informática otro.. así en cada root del site
#Metodo para ver la raiz de la biblioteca

def listSiteFolders(access_token):
    url = f'https://graph.microsoft.com/v1.0/sites/urdecon.sharepoint.com,{SITE_ID}/drives'
    headers = {
        'Authorization' : f'Bearer {access_token}',
        'Accept' : 'application/json'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

#Una vez tenemos el drive_id, cambian los endpoints, ahora apuntamos a drives/{drive_id}
#Aquí vamos a obtener los directorios debajo del raiz, como siempre es el mismo este metodo solo lista debajo del raiz
#Con el siguiente metodo podemos navegar por todas las bibliotecas
#pero lo dejo por ilustrar un poco

def listRootfolders(access_token):
    url = f'https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root/children'
    headers = {
        'Authorization' : f'Bearer {access_token}',
        'Accept' : 'application/json'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

##Como sabemos que todos los sharePoint están estructurados por Nombre/Documentos/General/, es debajo del directorio general donde empezamos a encontrar "chicha"
#por lo que la idea sería obtener el site_id y drive_id (id de la raiz (Documentos))
#la idea con este metodo es introducir rutas donde encontremos archivos, podría pedirse por consola la ruta (depende del uso que queramos darle)...
#una vez estamos dentro de Documentos/General/  los endpoints son https://graph.microsoft.com/v1.0/drives/{drive-id}/root:/{folder-path}:/children
# lo que se busca es el item_id, necesario para descargar ese archivo.

#Para poder ir probando cosas, le vamos a pasar la ruta por consola.

#Este metodo lista todos los elementos de un directorio
#para descargar un archivo necesitamos el item_id de este, por lo que al obtener el id en el main lo guardaremos en una variable

def listPathcontent(access_token,path):
    url = f'https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{path}:/children'
    headers = {
        'Authorization' : f'Bearer {access_token}',
        'Accept' : 'application/json'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

#metodo para descargar un archivo, podemos descargar en formato Excel o leer el contenido
#En este metodo, el contenido de la respuesta no está en formato JSON
#Se crea un nuevo Excel en la ruta que le pasemos, incluido el nombre,
# si solo le pasamos el nombre, lo crea en la carpeta donde se ejecuta el script

def downloadFile(access_token, fileName):
    url = f'https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/content'
    headers = {'Authorization' : f'Bearer {access_token}'
               }
    response = requests.get(url, headers=headers)
    response.raise_for_status()

    with open(fileName, "wb") as f:
        f.write(response.content)

## A diferencia del anterior metodo, en este lo que buscamos es, tomar el contenido de un Excel
## y hacerle "transformaciones" con la librería pandas

def transformExcel (access_token, fileName):
    url = f'https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/content'
    headers = {'Authorization': f'Bearer {access_token}'
               }
    response = requests.get(url, headers=headers)
    response.raise_for_status()

    file_bytes = io.BytesIO(response.content)
    df = pd.read_excel(file_bytes)

#En este punto ya tenemos almacenado en la variable 'df' el Excel, podemos realizar cualquier transformación a partir de este punto.
#a modo de ejemplo voy a crear una nueva columna con datos e introducir datos en las existentes

    df.loc[:4, "horasExtra"] = [2, 4, 6, 10, 25]
    df["email JO"] = df["email JO"].astype("object")
    df.loc[0, "email JO"] = "jparraga@urdecon.es"
    df.to_excel(fileName, index=False)

    print(f"Archivo Excel transformado y guardado en: {fileName}")

    #Los Excel se crearán en la carpeta del script.