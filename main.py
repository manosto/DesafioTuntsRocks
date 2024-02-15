import os.path
import math
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Google Sheets API credentials and settings
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1ZgbB8apEHlN-Vl1hSvJxJKDA4mW7KrDvGirlk96xiqk'
SAMPLE_RANGE_NAME = 'engenharia_de_software!A1:H27'
TOTAL_CLASSES_SEMESTER = 60

class Student:
    def __init__(self, data):
        self.id = int(data[0])
        self.name = data[1]
        self.absences = int(data[2])
        self.scores = list(map(int, data[3:6]))
        self.situation = ""
        self.final_exam_score = 0

    def calculate_average(self):
        return sum(self.scores) / 3


    def calculate_situation(self, total_classes):
        average = self.calculate_average()
        if self.absences > total_classes * 0.25:
            self.situation = "Reprovado por Falta"
        elif average < 50:
            self.situation = "Reprovado por Nota"
        elif 50 <= average < 70:
            self.situation = "Exame Final"
            self.calculate_final_exam_score()
        else:
            self.situation = "Aprovado"

    def calculate_final_exam_score(self):
        self.final_exam_score = math.ceil(2 * 50 -self.calculate_average())


def process_student_data(data):
    students = []
    total_classes = TOTAL_CLASSES_SEMESTER
    for row in data[3:]:
        student = Student(row)
        student.calculate_situation(total_classes)
        students.append(student)
    return students


def main():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    service = build("sheets", "v4", credentials=creds)

    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])
    
    if values:
        students = process_student_data(values)
        for student in students:
            sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="engenharia_de_software!G"+str(student.id + 3)+"", valueInputOption = "USER_ENTERED", body={"values": [[student.situation]]} ).execute()
            if student.situation == "Exame Final":
                sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="engenharia_de_software!H"+str(student.id + 3)+"", valueInputOption = "USER_ENTERED", body={"values": [[student.final_exam_score]]} ).execute()

if __name__ == "__main__":
    main()
