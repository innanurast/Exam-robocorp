from robocorp.tasks import task
from RPA.Excel.Files import Files
from RPA.Tables import Tables
import pandas as pd
from datetime import datetime, timedelta
import emailer
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from bs4 import BeautifulSoup
from datetime import datetime
import matplotlib.pyplot as plt
import os




TRAINING_CSV_PATH = "input/training.csv"
TRAINING_EXCEL_PATH = "output/training.xlsx"
PERSON_EXCEL_PATH = "input/Person.xlsx"

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

            # Filter training data for the person
            person_trainings = merged_df.loc[merged_df["Person ID"] == person_id, "Training name"].tolist()

            """convert dataframe to images"""
            # Ensure the 'images' directory exists to save the images
            image_directory = 'output/images'
            os.makedirs(image_directory, exist_ok=True)
            
            # Example: Creating and saving multiple figures
            for i in range(1, len(person_trainings)):
                #Create a table of upcoming trainings
                trainings_table = pd.DataFrame({
                    "Training name": person_trainings,
                    "Date training": [training_date] * len (person_trainings)
                })
                print(trainings_table)

                # Create a figure and plot the table
                fig, ax = plt.subplots(figsize=(8, 4))  # You can adjust the figure size
                ax.axis('tight')
                ax.axis('off')
                ax.table(cellText=trainings_table.values, colLabels=trainings_table.columns, cellLoc='center', loc='center')

                # Save the figure as an image
                image_filename = os.path.join(image_directory, f'{person_name}_{i}.png')
                plt.savefig(image_filename, bbox_inches='tight')

                print(f"{person_name} Figure {i} saved as {image_filename}")

                # Close the figure to free up resources
                plt.close()

                #Convert the DataFrame to HTML 
                html_table = trainings_table.to_html(index=False, border=1)
                print(html_table)
                #Create the body of the email with the formatted HTML table
                
                body = MIMEMultipart()
                body.attach(MIMEText(f"Hi, {person_name}!\n\nDon't forget your upcoming training:\n\n"))
                body.attach(MIMEText(BeautifulSoup(html_table, 'html.parser').get_text(), 'plain'))
                body.attach(MIMEText
                            (f"Best Regards, 
                             {emailer.account}!")
                )
                # Set the subject
                body["Subject"] = subject

                #Send the email
                emailer.send_email(recipient, subject, body.as_string())

                print("Successfully email sending reminder for training!")

            # # else :
            # #     print(f"No reminder training date for data person {person_id}")

@task
def minimal_task():
    csv_to_excel()
    get_active_persons(PERSON_EXCEL_PATH)
    reminder_training()
