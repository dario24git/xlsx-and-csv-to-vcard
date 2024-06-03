# xlsx and csv to vcard

CreaContatti is a user-friendly application that allows you to easily convert CSV or XLSX files into vCard format. With its intuitive interface and drag-and-drop functionality, creating vCard files has never been simpler.

## Features
- Convert CSV and XLSX files to vCard format
- Supports drag-and-drop for quick file selection
- Automatically generates a unique output file name based on the current date and time

## Requirements

Windows operating system
Microsoft Excel (if using XLSX files)

## Installation

Download the VCardGenerator.exe file from the releases page.
Double-click on the downloaded file to run the application.

## Usage
Launch the VCard Generator application.
Select a CSV or XLSX file to convert:
Click the "Select File" button and choose the desired file from the file dialog.
Alternatively, drag and drop the file onto the designated area in the application window.
The application will process the selected file and generate a vCard file with a unique name based on the current date and time.
A message box will appear to confirm the successful conversion of the file.
The generated vCard file will be saved in the same directory as the VCard Generator application.

## CSV and XLSX File Format

To ensure proper conversion, follow these guidelines for your CSV and XLSX files:

- The file should contain two columns: the first column for the full name and the second column for the phone number.
- Do not include column headers in the file.
- For CSV files, use a comma (,) or semicolon (;) as the delimiter.
- For XLSX files, create a new worksheet and enter the data starting from the first row, without any headers.

### Example CSV file format:

John Doe,1234567890
Jane Smith,9876543210

### Example XLSX file format:

|John Doe|1234567890|
|Jane Smith|9876543210|

## Troubleshooting
If the application fails to generate the vCard file, ensure that the input file is in the correct format and follows the guidelines mentioned above.
If you encounter any issues or have questions, please open an issue on the GitHub repository.

## License
This project is licensed under the MIT License.

## Acknowledgements

VCard Generator was developed using the following libraries:

- PyQt5
- pandas
- vobject

We would like to thank the developers and contributors of these libraries for their valuable work.
