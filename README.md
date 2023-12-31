# Automated-Certificate-Generation-in-Python-Tkinter #Mail Sender
# Intro:

👋 I'm excited to share with you a Python program I recently developed to streamline the process of generating certificates in bulk. This program is designed to automate the creation of certificates based on information provided in an Excel sheet and a PowerPoint template.

## Features:
- Bulk Certificate Generation: Efficiently generate certificates in bulk.
- Customizable Templates: Design certificates in PowerPoint with dynamic placeholders (<<>>).
- User-Friendly: Simple and Easy-to-use GUI with tkinter for file selections.

## How to Use:
# Certificate Generator
- 📂 Select Excel File: Choose an Excel sheet containing participant information.
- 🖼️ Select PPTX Template: Design your certificate in PowerPoint with dynamic placeholders (<< text>>).
- 📁 Select Destination Folder: Choose where the generated certificates will be saved.
- 🚀 Start Merging: Run the program, and watch as personalized certificates are created!
  
# ✉️ Mail Dispatcher
- Additionally Enter Gmail id.
- Create App password and enter it.
- Confirm and click Start Merging Button.
  
## Limitation
- The names in the excel and the pptx must match.
For example the word **rollnumber** (column name) in excel must be written as << rollnumber >> in pptx file.
Can only send from a Gmailaccount.

## Requirements:
- Python (3.x recommended)
- tkinter
- pandas
- python-pptx
  
## Mail Sender Packages:
- smtplib
- EmailMessage

## Result/Test files:
To view the test files/ results click the [dropbox link](https://www.dropbox.com/scl/fo/msij5afvlkwwhzd5mkkyr/h?rlkey=amo6dz30e6rtfsh47uzka45kn&dl=0)

Let's make certificate generation hassle-free! 💼✨[CODE_LINK](https://github.com/vvs7722/Automated-Certificate-Generation-in-Python/blob/main/code/Certificate_Generator.py)
Certificate generation and Email Sender! 💼 ✨ [Code_Link](code/GeneratorAndMailSender.py)
