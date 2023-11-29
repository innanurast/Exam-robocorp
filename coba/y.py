from robocorp.tasks import task
from RPA.Excel.Files import Files
from RPA.Tables import Tables
import pandas as pd
from datetime import datetime, timedelta
import emailer
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import matplotlib.pyplot as plt
import os
import smtplib

TRAINING_CSV_PATH = "../input/training.csv"
TRAINING_EXCEL_PATH = "../output/training.xlsx"
PERSON_EXCEL_PATH = "../input/Person.xlsx"

tables = Tables()

def csv_to_excel():
    """ Read CSV file into a DataFrame """
    df = pd.read_csv(TRAINING_CSV_PATH)
    df.to_excel(TRAINING_EXCEL_PATH, index=False, sheet_name="data")

def read_excel_as_table(path):
    """read excel file as a table"""
    excel = Files()
    excel.open_workbook(path)
    try:
        return excel.read_worksheet_as_table(header=True)
    finally:
        excel.close_workbook()

def get_active_persons(excel_path):
    """read and filter data by status and category"""
    persons = read_excel_as_table(excel_path)
    tables.filter_table_by_column(persons, "Status", "==", "Active")
    return persons

def get_table_training(excel_path):
    trainings = read_excel_as_table(excel_path)
    return trainings

def table_to_dataframe(table):
    """Convert RPA.Tables.Table to pandas DataFrame"""
    data = []
    for row in table:
        row_data = {}
        for col_name, cell in row.items():
            row_data[col_name] = cell
        data.append(row_data)
    return pd.DataFrame(data)

def send_email(recipient, subject, body, attachment_path, html_table=None, close_message=None):
    sender_email = "innanur16@gmail.com"
    sender_password = "tino hnxh rjry sukd"

    # Create the email message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient
    message['Subject'] = subject

    # Attach body text
    message.attach(MIMEText(body, 'plain'))

    # Attach HTML table
    if html_table:
        html_part = MIMEText(html_table, 'html')
        message.attach(html_part)
    
    # Attach closing message
    if close_message:
        message.attach(MIMEText(close_message, 'plain'))

    # Attach image
    with open(attachment_path, 'rb') as image_file:
        image_attachment = MIMEImage(image_file.read(), name=os.path.basename(attachment_path))

    message.attach(image_attachment)

    # Connect to the SMTP server and send the email
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(message)

def reminder_training():
    active_persons = get_active_persons(PERSON_EXCEL_PATH)
    tbl_training = get_table_training(TRAINING_EXCEL_PATH)

    # Convert RPA.Tables.Table to pandas DataFrame
    active_persons_df = table_to_dataframe(active_persons)
    tbl_training_df = table_to_dataframe(tbl_training)

    # Merge DataFrames based on "Person ID"
    merged_df = pd.merge(active_persons_df, tbl_training_df, on="Person ID", how="inner")

    # Print the merged DataFrame for debugging
    print("Merged DataFrame:")
    print(merged_df)

    # Get the current date
    current_date = datetime.now()
    week_from_now = current_date + timedelta(days=2)
    print(week_from_now)

    # Keep track of unique recipients to avoid duplicate emails
    unique_recipients = set()

    # Iterate through merged data
    for _, row in merged_df.iterrows():
        # Get details from the merged row
        person_id = row["Person ID"]
        training_date = datetime.fromisoformat(row["Date training"])

       # Check if training date is within a week from now
        if current_date <= training_date <= week_from_now:
            # Send a reminder email for the upcoming training
            person_name = f"{row['First Name']} {row['Last Name']}"
            recipient = row["Email"]
            subject = "Remember to complete your training!"

            if recipient not in unique_recipients:
                # Add the recipient to the set to mark as sent
                unique_recipients.add(recipient)
                # Filter training data for the person
                person_trainings = merged_df.loc[merged_df["Person ID"] == person_id, "Training name"].tolist()

                # Ensure the 'images' directory exists to save the images
                image_directory = 'output/images'
                os.makedirs(image_directory, exist_ok=True)

                # Create a table of upcoming trainings
                trainings_table = pd.DataFrame({
                    "Training name": person_trainings,
                    "Date training": [week_from_now] * len(person_trainings)
                })

                # Convert DataFrame to HTML table
                html_table = trainings_table.to_html(index=False)

                # Create a figure and plot the table
                fig, ax = plt.subplots(figsize=(8, 4))
                ax.axis('tight')
                ax.axis('off')
                ax.table(cellText=trainings_table.values, colLabels=trainings_table.columns, cellLoc='center', loc='center')

                # Save the figure as a single image for the person
                image_filename = os.path.join(image_directory, f'{person_name}.png')
                plt.savefig(image_filename, bbox_inches='tight')

                print(f"{person_name} Figure saved as {image_filename}")

                # Create the email body
                # Create the email body
                # body = f"""Hi, {person_name}!\n\n
                #         Don't forget your upcoming training:\n\n
                #         I hope this email finds you well. As part of our ongoing commitment to your professional development, 
                #         we would like to remind you of your upcoming training session.
                #         Training Details:
                        

                #         Your participation is crucial for the success of the program, and we believe this training will contribute significantly to your skill set.

                #         If you have any questions or concerns, please don't hesitate to reach out to us.

                #         Thank you, and we look forward to seeing you at the training!
                #         """


                close_message = f"Best Regards,\n\n Ina Nur Astuti!"

                # Send the email with the attached image and HTML table
                send_email(recipient, subject, f"Hi, {person_name}!\n\n Don't forget your upcoming training!\n\n I hope this email finds you well. As part of our ongoing commitment to your professional development, we would like to remind you of your upcoming training session.\n\n Training Details:", image_filename, html_table, close_message)

@task
def minimal_task():
    csv_to_excel()
    get_active_persons(PERSON_EXCEL_PATH)
    reminder_training()
