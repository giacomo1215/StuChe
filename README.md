# StuChe (Student Checker)

A web-based tool for analyzing student appointment data from Excel spreadsheets.

## Features

- Upload Excel (.xlsx) files containing student appointment data
- Automatically identifies students requiring follow-up based on the following criteria:
  - Missing first appointment
  - Second appointment marked with "-"
  - Third appointment marked with "-"
- Displays results in an easy-to-read table format

## Setup

1. Clone this repository
2. Open the project in a web browser
   - Use Live Server in VS Code, or
   - Set up any HTTP server pointing to the project directory

## Usage

1. Click the file upload button
2. Select an Excel file containing student data with the following columns:
   - SID (Student ID)
   - First Name
   - Last Name
   - 1 Appt
   - 2 Appt
   - 3 Appt
3. The results will automatically display showing students requiring follow-up

## Technologies Used

- HTML5
- Bootstrap 5.3
- JavaScript
- SheetJS (XLSX)

## Project Structure

```
StuChe/
├── assets/
│   ├── css/
│   │   └── styles.css
│   └── js/
│       └── scripts.js
├── index.html
└── README.md
```
