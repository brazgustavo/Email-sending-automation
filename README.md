Pandas Email Alert System

The goal of this project is to send emails to suppliers of items that are impacting the stock of NOK items. This code reads a set of Excel files, filters the items of interest, and sends personalized emails to each supplier asking them to show the destination of the parts, otherwise we will act according to the company's policy.

Requirements

To execute this project, the following libraries must be installed:
• Pandas
• win32com

How to use

Before running this code, you must:

Prepare the materials lists, database, and supplier contacts in Excel files.
Make sure the list of suppliers contains the following columns:
• "Supplier"
• "Supplier_Name"
• "Email"
Put the Excel files in the same directory as this code.
Icons

This project uses icons to indicate the action that will be taken in each part of the code:
• 📥: Importing databases.
• 🗃️: Treating databases.
• 🤝: Joining databases.
• 📉: Selecting desired columns.
• 🗑️: Removing rows with empty or null values.
• 📤: Sending emails.

How it works

The databases are imported from Excel files using the Pandas library.
The databases are treated using Pandas functions to select only the information of interest.
The databases are joined using the Pandas merge() function.
The columns of interest are selected using the Pandas drop() function.
Rows with empty or null values are removed using the Pandas dropna() function.
A connection is created with Outlook to send emails.
For each material of interest, the corresponding supplier in the contact list is found.
A personalized email is generated for each supplier.
The necessary attachments are added to the email.
The email is sent using the send() function of the MailItem object.
Confidentiality note

Please note that we will not be sharing the databases due to the confidentiality of the company.
