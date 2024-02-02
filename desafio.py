## From sheets, api and drive from google.

import os.path
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# modify the scopes to manipulate.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1t4wJJ_pIqqGkXQNCQZU322HKIJNfTycPWXev8iCVAeg"
SAMPLE_RANGE_NAME = "engenharia_de_software!A4:H27"


def main():
    """Shows basic usage of the Sheets API. Prints values from a sample spreadsheet."""
    creds = None
    # The file token.json stores the user's access and refresh tokens and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()

        # Check if 'values' key exists in the result dictionary

        # using datas from 'values' and assume as valors
        if 'values' in result:
            valors = result['values']
          # creating variables and assuming 'lines and columns'
            add_situacao = []
            add_naf = []

            for i, line in enumerate(valors):
                absences = int(line[2])
                r_absences = 15
                p1 = int(line[3])
                p2 = int(line[4])
                p3 = int(line[5])
                m = (p1 + p2 + p3) / 3

                if absences >= r_absences: # calculating the situacion of absence
                    situacion = "Reprovado por falta"
                elif m < 50:
                    situacion = "Reprovado por nota" # making condicions for 'exame final and aprove
                elif 50 <= m < 70:
                    situacion = "Exame Final"
                else:
                    situacion = "Aprovado"

                add_situacao.append([situacion])

                if situacion == "Exame Final": # creating the if for 'Exame Final"
                    naf = int(m - 10)
                else:
                    naf = 0
                add_naf.append([naf])

                # Update "G" cell dynamically based on the current student's position
                result = sheet.values().update(
                    spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"G{i+4}",
                    valueInputOption="USER_ENTERED",
                    body={'values': [[situacion]]}
                ).execute()

                 #Update "H" cell dynamically based on the current student's position
                result = sheet.values().update(
                    spreadsheetId=SAMPLE_SPREADSHEET_ID,
                    range=f"H{i+4}",
                    valueInputOption="USER_ENTERED",
                    body={'values': [[naf]]}
                ).execute()

    except HttpError as err: # show me in case apear an error
        print(err)


if __name__ == "__main__":
    main()
