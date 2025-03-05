INN Data Parser
A Python-based data parsing tool designed to extract contact information (emails, phone numbers, and websites) from a list of Taxpayer Identification Numbers (INNs) using the List-org website. The tool features a user-friendly Tkinter GUI for easy interaction and saves the results in an Excel file.

Key Features:
GUI Interface: Simple and intuitive graphical interface for selecting input files, saving results, and choosing data sources.

Multi-Step Workflow: Guides users through file selection, saving options, and source selection.

Selenium Automation: Automates web scraping to extract contact information from the List-org website.

Error Handling: Includes validation for input files and user-friendly error messages.

Excel Output: Saves parsed data (INN, email, website, phone numbers) in a structured Excel file.

How It Works:
Input: Provide an Excel file containing a list of INNs.

Source Selection: Choose the data source (currently supports List-org).

Parsing: The tool automatically navigates the website, extracts contact information, and handles CAPTCHA prompts.

Output: Results are saved in an Excel file at the specified location.

Technologies Used:
Python: Core programming language.

Tkinter: For building the graphical user interface.

Selenium: For web scraping and automation.

Pandas: For handling Excel files and data processing.

Future Improvements:
Support for additional data sources.

Enhanced CAPTCHA handling.

Multi-threading for faster processing.

Installation:
Clone the repository:

bash
Copy
git clone https://github.com/your-username/inn-data-parser.git
Install dependencies:

bash
Copy
pip install -r requirements.txt
Run the application:

bash
Copy
python main.py
Usage:
Select an Excel file containing INNs.

Choose a save location for the output file.

Select the data source (List-org).

Click "Finish" to start parsing.

The results will be saved in the specified Excel file.

Contributing:
Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

License:
This project is licensed under the MIT License. See the LICENSE file for details.

