# INN Data Parser

A Python-based data parsing tool designed to extract contact information (emails, phone numbers, and websites) from a list of Taxpayer Identification Numbers (INNs) using the **List-org** website. The tool features a user-friendly **Tkinter GUI** for easy interaction and saves the results in an Excel file.

## Key Features

- **GUI Interface**: Simple and intuitive graphical interface for selecting input files, saving results, and choosing data sources.
- **Multi-Step Workflow**: Guides users through file selection, saving options, and source selection.
- **Selenium Automation**: Automates web scraping to extract contact information from the List-org website.
- **Error Handling**: Includes validation for input files and user-friendly error messages.
- **Excel Output**: Saves parsed data (INN, email, website, phone numbers) in a structured Excel file.

## How It Works

1. **Input**: Provide an Excel file containing a list of INNs.
2. **Source Selection**: Choose the data source (currently supports List-org).
3. **Parsing**: The tool automatically navigates the website, extracts contact information, and handles CAPTCHA prompts.
4. **Output**: Results are saved in an Excel file at the specified location.

## Technologies Used

- **Python**: Core programming language.
- **Tkinter**: For building the graphical user interface.
- **Selenium**: For web scraping and automation.
- **Pandas**: For handling Excel files and data processing.

## Future Improvements

- Support for additional data sources.
- Enhanced CAPTCHA handling.
- Multi-threading for faster processing.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Gerrgor/inn_parser.git
   ```bash
   pip install -r requirements.txt
   ```bash
   python main.py
