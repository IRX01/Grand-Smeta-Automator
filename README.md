# Grand Smeta Automator

Automation tool for processing construction estimate documents generated from Grand-Smeta.

This project automates the routine workflow of processing Excel documents such as KS-2, KS-3 and related reports.  
The script significantly reduces manual work by automatically cleaning, editing and synchronizing data between multiple documents.

## 🚀 Features

The bot performs the following operations automatically:

### 1. Excel document processing
- Opens Excel documents generated from Grand-Smeta
- Extracts estimate numbers and dates from file names
- Automatically calculates:
  - first day of month
  - last day of month
- Updates document metadata

### 2. KS-2 processing
- Finds the line containing **contract estimate cost**
- Extracts the total amount
- Cleans unnecessary sections of the document
- Keeps only required structural blocks:
  - estimate cost
  - main table
  - final totals
  - signature section

### 3. KS-3 processing
- Automatically finds the KS-3 document in the project folder
- Updates:
  - construction address
  - document number
  - report period
  - final contract amount

### 4. File automation
- Automatically saves processed Excel files
- Generates correct file names based on estimate metadata
- Locates project folders dynamically

### 5. UI Launcher
A simple desktop interface allows the user to:
- select project root folder
- run the automation pipeline with one click

## 🧠 Technologies Used

- Python
- Excel COM Automation
- pywin32
- pyautogui
- openpyxl
- pyperclip

## 🏗 Architecture

The automation pipeline works as follows:

Grand-Smeta 
  
↓  

1. Export Excel documents  
2. Process **Base Estimate**  
3. Save **Base Estimate**  
4. Process **Defective act**  
5. Edit **Defective act**  
6. Save **Defective act**  
7. Process **KS-2**  
8. Edit **KS-2**  
9. Extract total amount  
10. Save **KS-2**  
11. Process **KS-3**  
12. Edit **KS-3**  
13. Save **KS-3**

↓  

End

## 📈 Benefits

The tool reduces manual work involved in estimate document processing.

Typical manual workflow:
- open multiple Excel files
- copy data between documents
- update dates and numbers
- clean document structure

The automation script performs these steps automatically.

Estimated time saved:
**5-15 minutes per estimate package.**

## 💡 Motivation

This project was created to automate repetitive document processing tasks in construction cost estimation workflows.

The goal was to build a practical automation tool using Python and Excel COM automation.

## 📂 Project Status

Active development.

Future improvements:
- full GUI application
- improved document detection
- logging system
- packaged executable version

## 👨‍💻 Author

Python automation project for workflow optimization.
