import gspread
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


def get_user_info_from_sheets(email):

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly"
    ]
    
    #Cargamos los credenciales del API de google sheets
    creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    client = gspread.authorize(creds)

    #ID del spreadsheet para acceder y encontrar los correos
    sheet_id = "1kPe7V3tWMvpXx7FtDj-Y_83ve3MSwkhTdI8eC4kdEMM"
    sheet = client.open_by_key(sheet_id).sheet1
    user_info = []
    
    try:
        row = sheet.find(email).row
        user_info = sheet.row_values(row)
    except gspread.exceptions.CellNotFound:
        pass
    
    return user_info

def update_gmail_signature(email,  signature, live=False): 
    
    scopes = [
        "https://www.googleapis.com/auth/gmail.settings.basic"
        "https://www.googleapis.com/auth/gmail.settings.sharing"
    ]
    
    
    #Cargamos los credenciales del API de gmail
    creds = Credentials.from_service_account_file("gmail_credentials.json", scopes=scopes)
    if live:
        creds_delegated = creds.with_subject(email)
        
        #Creamos la instancia de la API de gmail
        gmail_service = build('gmail', 'v1', credentials=creds_delegated)
        addresses = gmail_service.users().settings().sendAs().list(userId="me", fields="seendAs(isPrimary, sendAsEmail)").execute().get("sendAs")
        
        address = None
        for address in addresses:
            if address.get('isPrimary'):
                break
        if address:
            rsp = gmail_service.users().settings().sendAs().patch(userId='me', sendAsEmail=address['sendAsEmail'], body={'signature': signature}).execute()
            print(f'Signature changed for: {email}')
        else:
            print(f'Could not find primary address for: {email}')

    

#Cargamos el contenido de la plantilla HTML
with open('template.html', 'r') as file:
    template_html = file.read()
    

#Funcion para obtener un listado de los emails de los sheets
def get_emails_from_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly"
    ]
    
    #Cargamos los credenciales del API de google sheets
    creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    client = gspread.authorize(creds)

    #ID del spreadsheet para acceder y encontrar los correos
    sheet_id = "1kPe7V3tWMvpXx7FtDj-Y_83ve3MSwkhTdI8eC4kdEMM"
    sheet = client.open_by_key(sheet_id).sheet1
    emails_from_sheets = sheet.col_values(4)[1:]
    return emails_from_sheets


#Construimos la lista de usuarios
usuarios = [{'email': email} for email in get_emails_from_sheets()]  

#Pasamos por cada usuario de la lista para actualizar su firma
for user in usuarios: 
    email = user['email']
    
    #Obtenemos la info del usuario desde google sheets
    user_info = get_user_info_from_sheets(email)
    if user_info:
        departamento, cargo, nombre = user_info[1:4]
        
        #Reemplazamos los marcadores del html por las variables
        firma_personalizada = template_html.replace('{{NOMBRE}}', nombre).replace('{{CARGO}}', cargo).replace('{{DEPARTAMENTO}}', departamento)
        
        #Actualizamos la firma para el usuario en gmail
        update_gmail_signature(email, firma_personalizada)
        print(f'Firma actualizada para {email}')
    else:
        print(f'No se encontro información para {email} en google sheets')

print('Proceso completo.')
    
    