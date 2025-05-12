import argparse
import logging
import os
import re

import win32com.client
from bs4 import BeautifulSoup


def extract_case_ids_from_email(subject, folder_path, output_file_name, output_path):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    #Default folder is the inbox, you can change it to any other folder you want to search in

    if folder_path is not None:
    # Split the folder path and navigate to the specific folder
        folders = folder_path.split("\\")
        folders = [folder for folder in folders if folder]
        folder = outlook.Folders.Item(2)
        
        for subfolder in folders:
            
            folder = folder.Folders.Item(subfolder)
        messages = folder.Items

    else:
        # If no folder is specified, use the default inbox
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items


    #replace args.subject with the subject you want to filter by
    messages = messages.Restrict(f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject}%'") # Replace with your desired keyword


    if messages.Count > 0:
        message = messages.GetLast()
        # Get the last message in the filtered list
        # Extract the HTML body content
        html_body = message.HTMLBody

        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(html_body, 'html.parser')
    
        case_ids = []

        for a_tag in soup.find_all('a'):
            # Extract case IDs using regular expression
            ids = re.findall(r'\b\d{16}\b', a_tag.text)
            # Add the found IDs to the case_ids list
            case_ids.extend(ids)

  
        #save file to args.outputPath with the name args.outputFileName with txt extension
        with open(f"{output_path}/{output_file_name}.txt", "w") as file:
            for case_id in case_ids:
                file.write(f"{case_id},\n")

    logging.info("Case IDs extracted successfully.")
    logging.info(f"Number of case IDs extracted: {len(case_ids)}")
    logging.info(f"Case IDs saved to {output_path}/{output_file_name}")


if __name__ == "__main__" : 
    parser = argparse.ArgumentParser(description="Extract case IDs from Outlook emails.")
    parser.add_argument("--subject", type=str, default="Cx Story Max A", help="Subject of the email to filter.")
    parser.add_argument("--folderPath", type=str, default="Inbox", help="Name of the Outlook folder to search.Default is Inbox. otherwise Specify the folder path.")
    parser.add_argument("--outputFileName", type=str, default="case_ids.txt", help="Output file name.")
    parser.add_argument("--outputPath", type=str, default=f"{os.getcwd()}", help="Default is current working directory,Output file path.")
    logging.basicConfig(level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')
    args = parser.parse_args()

    logging.info("Starting the script...")
    logging.info(f"Subject: {args.subject}")
    logging.info(f"Folder Path: {args.folderPath}")
    logging.info(f"Output File Name: {args.outputFileName}")
    logging.info(f"Output Path: {args.outputPath}")

    # Call the function to extract case IDs from the email
    extract_case_ids_from_email(args.subject, args.folderPath, args.outputFileName, args.outputPath)

    logging.info("Script completed.")
    logging.info("Exiting...")  

    