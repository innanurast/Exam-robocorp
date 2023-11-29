from robocorp.tasks import task
from RPA.Excel.Files import Files
from RPA.Tables import Tables
import pandas as pd
from datetime import datetime, timedelta
import emailer
from email.mime.text import MIMEText

TRAINING_CSV_PATH = "input/training.csv"
TRAINING_EXCEL_PATH = "input/training_excel.xlsx"
STUDENT_EXCEL_PATH = "input/Students.xlsx"

tables = Tables()

# def csv_to_excel():
#     """ Read CSV file into a DataFrame """
#     df = pd.read_csv(TRAINING_CSV_PATH)
#     df.to_excel(TRAINING_EXCEL_PATH, index=False, sheet_name="data")

def read_excel_as_table(path):
    """read excel file as a table"""
    excel = Files()
    excel.open_workbook(path)
    try:
        return excel.read_worksheet_as_table(header=True)
    finally:
        excel.close_workbook()

def get_active_students(excel_path):
    """read and filter data by status and category"""
    students = read_excel_as_table(excel_path)
    tables.filter_table_by_column(students, "Status", "==", "Active")
    return students

def get_table_training(excel_path):
    trainings = read_excel_as_table(excel_path)
    return trainings

def get_matching_training_data(student_id, training_table):
    return tables.filter_table_by_column(
        training_table, "Person ID", "==", student_id
    )

def reminder_training():
    active_students = get_active_students(STUDENT_EXCEL_PATH)
    tbl_training = get_table_training(TRAINING_EXCEL_PATH)

    # Print active students for debugging
    print("Active Students:")
    print(active_students)

    print("training:")
    print(tbl_training)

    # Get the current date
    current_date = datetime.now()
    week_from_now = current_date + timedelta(days=2)
    print(week_from_now)

    # Set to track notified training names for each student
    notified_trainings = {}

    # Iterate through active students
    for student in active_students.get_table():
        student_id = student["Person ID"]
        print(f"student id : {student_id}")

        # Find matching training data for the student ID
        matching_trainings = tables.filter_table_by_column(
            tbl_training, "Person ID", "==", student_id
        )
        print(f"match : {matching_trainings}")

        # # Get the set of training names for this student
        # student_trainings = set(tables.get_table_column(matching_trainings, "Training name"))

        # # Check for previously notified training names
        # student_notified_trainings = notified_trainings.get(student_id, set())

        # # Filter out already notified trainings
        # pending_trainings = student_trainings - student_notified_trainings

        # # Iterate through pending trainings
        # for training in pending_trainings:
        #     # Get training details
        #     training_details = tables.filter_table_by_column(
        #         matching_trainings, "Training name", "==", training
        #     )

        #     training_date = datetime.fromisoformat(training_details.get_table()[0]["Date training"])

        #     if training_date <= week_from_now:
        #         # Send a reminder email for the upcoming training
        #         name = f"{student['First Name']} {student['Last Name']}"
        #         recipient = student["Email"]
        #         subject = "Remember to complete your training!"

        #         # Create a table of upcoming trainings
        #         trainings_table = pd.DataFrame({
        #             "Training name": [training],
        #             "Date training": [training_date]
        #         })

        #         # Convert the DataFrame to HTML with CSS styles
        #         html = trainings_table.to_html(style='<style>' +
        #                                             'table { border-collapse: collapse; width: 100%; }' +
        #                                             'td, th { border: 1px solid #dddddd; text-align: left; padding: 8px; }' +
        #                                             '</style>')

        #         # Create the body of the email with the formatted HTML table
        #         body = MIMEText(f"Hi, {name}! \n\nDon't forget your upcoming trainings: \n\n" + html, 'html')
        #         body.set_subject(subject)

        #         # Send the email
        #         emailer.sendmail(recipient, body.as_string())

        #         # Update notified trainings for this student
        #         student_notified_trainings.add(training)
        #         notified_trainings[student_id] = student_notified_trainings

@task
def minimal_task():
    # csv_to_excel()
    get_active_students(STUDENT_EXCEL_PATH)
    reminder_training()
