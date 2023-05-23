import csv
import os
import time

import pandas as pd
import urllib.parse
import re
import shutil

folder_path = 'EmailExel'
csv_file = 'EmailCSV/Emails.csv'
new_csv_file = 'ParseEmailList/email_list.csv'
unique_emails = 'UniqueEmailList/unique_emails.csv'
gmail_file = 'GmailOtherEmailList/gmail_emails.csv'
other_file = 'GmailOtherEmailList/other_emails.csv'

# Convert Excel file to CSV
def convert_excel_to_csv(excel_file, csv_file):
    file_extension = excel_file.split('.')[-1].lower()
    if file_extension == 'csv':
        # File is already in CSV format
        shutil.move(excel_file, csv_file)
    elif file_extension == 'xlsx':
        # Excel file
        df = pd.read_excel(excel_file, engine='openpyxl')
        df.to_csv(csv_file, index=False)
    elif file_extension == 'json':
        # JSON file
        df = pd.read_json(excel_file)
        df.to_csv(csv_file, index=False)
    # Add more conditions for other file formats
    else:
        print(f"Unsupported file format: {file_extension}")
    print("File "+ excel_file + " converted to CSV successfully!")
    if os.path.exists(excel_file):
        os.remove(excel_file)
    print("File "+ excel_file + " deleted successfully!")

# URL decode a string
def url_decode(string):
    return urllib.parse.unquote(string)


# Parse email addresses using regex
def parse_email_addresses(line):
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'
    email_addresses = re.findall(pattern, line)
    return email_addresses

# Write text to a new CSV file
def write_text_to_new_csv(email, csv_file):
    with open(csv_file, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([email])



def extract_emails():
    with open(csv_file, 'r') as file:
        for line in file:
            # URL decode the line
            decoded_line = url_decode(line)

            # Parse email addresses using regex
            email_addresses = parse_email_addresses(decoded_line)
            for email in email_addresses:
                Email = email
                # Write the email to the new CSV file
                write_text_to_new_csv(Email, new_csv_file)
    print(excel_file + " Emails extracted successfully!")
    if os.path.exists(csv_file):
        os.remove(csv_file)
    print("File "+ csv_file + " deleted successfully!")
def extract_unique_emails(new_csv_file, output_csv):
    unique_emails = set()

    with open(new_csv_file, 'r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            if len(row) > 0:
                email = row[0].strip()
                if email not in unique_emails:
                    unique_emails.add(email)

    with open(output_csv, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for email in unique_emails:
            writer.writerow([email])
    print(excel_file+" Unique Emails extracted successfully!")
    if os.path.exists(new_csv_file):
        os.remove(new_csv_file)
    print("File "+ new_csv_file + " deleted successfully!")

def extract_emails_gmail_other(input_csv, gmail_csv, other_csv):
    gmail_emails = set()
    other_emails = set()

    # Read the input CSV file and extract Gmail and other emails
    with open(input_csv, 'r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            if len(row) > 0:
                email = row[0].strip()
                if email.endswith('@gmail.com'):
                    gmail_emails.add(email)
                else:
                    other_emails.add(email)

    # Write Gmail emails to the Gmail CSV file
    with open(gmail_csv, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for email in gmail_emails:
            writer.writerow([email])

    # Write other emails to the other CSV file
    with open(other_csv, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for email in other_emails:
            writer.writerow([email])
    print(excel_file+" Gmail and Other Emails extracted successfully!")
    if os.path.exists(input_csv):
        os.remove(input_csv)
    print("File "+ input_csv + " deleted successfully!")


def count_emails(csv_file):


    # Open the CSV file
    with open(csv_file, 'r') as file:
        # Create a CSV reader object
        reader = csv.reader(file)

        # Count the number of lines in the CSV file
        line_count = sum(1 for _ in reader)

    # Print the line count
    return line_count


def get_excel_file():
    folder_path = 'EmailExel'
    files = os.listdir(folder_path)


    if files:
        excel_file = os.path.join(folder_path, files[0])
        return excel_file


    else:
        print("No  files found in the folder.")
        excel_file = None
        return excel_file




while get_excel_file() != None:
    excel_file = get_excel_file()
    convert_excel_to_csv(excel_file, csv_file)
    extract_emails()
extract_unique_emails(new_csv_file, unique_emails)
extract_emails_gmail_other(unique_emails, gmail_file, other_file)
print("Gmail Emails count:" + str(count_emails(gmail_file)))
print("Other Emails count: " + str(count_emails(other_file)))
print("Total Emails count: " + str(count_emails(gmail_file) + count_emails(other_file)))
