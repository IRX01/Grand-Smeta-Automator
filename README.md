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
Export Excel documents  
↓  
Process **Base Estimate**  
↓
Save **Base Estimate**
↓
Process **Defective act**  
↓  
Edit **Defective act**  
↓
Save **Defective act**
↓
Process **KS-2**  
↓ 
Edit a **KS-2**  
↓
Extract total amount  
↓ 
Save **KS-2**
↓
Process **KS-3**
↓
Edit a **KS-3**  
↓
Save **KS-3**
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
