
# Management System Software

This Management System Software is a Python application built using the Tkinter library for creating a GUI, and it utilizes the openpyxl library for handling Excel files. The software provides a user-friendly interface for managing student information, allowing users to input details such as name, email ID, contact number, and more, and saves this data to an Excel spreadsheet.

## Features

- **User-friendly Interface**: The software provides a simple and intuitive interface for entering and managing student information.
- **Data Validation**: It includes validation checks for ensuring the correctness of entered data, such as validating email IDs and contact numbers.
- **Dynamic Comboboxes and Radio Buttons**: The software utilizes comboboxes and radio buttons for selecting options such as state, semester, training session, etc.
- **Error Handling**: It incorporates error messages to alert users in case of invalid data entry or file access issues.

## Prerequisites

- Python 3.x
- Tkinter library
- openpyxl library

## Installation

1. Clone the repository to your local machine:

```
git clone https://github.com/yourusername/management-system-software.git
```

2. Navigate to the project directory:

```
cd management-system-software
```

3. Install the required dependencies:

```
pip install -r requirements.txt
```

## Usage

1. Run the application:

```
python main.py
```

2. Fill in the necessary student information in the provided fields.
3. Click the "SAVE" button to save the entered data to an Excel spreadsheet.
4. Click the "CLEAR" button to reset the form fields.
