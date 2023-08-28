# SimpleDocumentRegisterCheck

# Document Checker App

The **Document Checker App** is a PyQt5-based desktop application that allows you to check if a list of documents specified in an Excel file are present within a specified folder and its subfolders. The application provides a user-friendly interface to browse for the Excel file, folder location, and set the output file for the results.

## Features

- Browse and select an Excel file containing a list of document titles.
- Choose a folder location to search for the specified documents.
- Set the output file to save the results of the document checking process.
- Check if each document title in the Excel file is present in the specified folder and its subfolders.
- Generate a new Excel file with an additional column indicating whether each document is found or not.

## Prerequisites

- Python 3.x
- PyQt5 library
- openpyxl library

## Installation

1. Make sure you have Python 3.x installed on your system.
2. Install the required libraries using the following command:
   

## Usage

1. Run the application by executing the script `document_checker.py`.
2. The application window will appear with fields to input the Excel file, folder location, and output file.
3. Click the "Browse" buttons to select the respective files/folders.
4. After filling in all the fields, click the "Check Documents" button to initiate the document checking process.
5. The results will be saved to the specified output Excel file, including a new column indicating whether each document is found ("Y") or not ("N").

## Styling and Customization

- You can customize the appearance of the application by modifying the style attributes in the script. Feel free to adjust colors, fonts, and layout to match your preferences.

## Notes

- The application currently supports Excel files in the `.xlsx` format.
- The specified folder location is searched recursively, including all subfolders.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

