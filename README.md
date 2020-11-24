# VBAInvoiceGenRepository
Licensed under the [MIT License](LICENSE).

This Invoice Generator Tool is a part of an Excel VBA course on Udemy.com
'Unlock Excel VBA and Excel Macros' 

Purpose of the Invoice Generator Tool:
The aim of this tool is to automate the invoice generation procedure and keep an overview of the status of all invoices in one place. 

This tool makes creating invoices a pleasant admin task. It no longer takes 20-30 minutes per invoice but instead just a few clicks that takes 2-3 minutes.
Having the information for all the invoices in one place, allows you to keep an eye on jobs completed,
open invoices and money made.

1. Update Master Data

Input your original customer master data in the master data tab. 
Add new customer from the Dashboard tab, by clicking on Add Customer button.
Edit customer information by clicking on Edit Master Data button.

2. Update the Folder paths

Make sure you create two folders to keep the PDF versions as well as Excel Versions for each invoice created.
The folders should be created in the directory as wherever this tool is saved. VBA uses thisworkbook.path to find 
the directory and then looks for the folders you have specified on the dashboard page inside this directory.

3. Adding Invoice Information

As projects come in, add the information to the invoice tab. Start typing the name of the customer in column B 
then press Enter and select the correct customer from the list. Type in the remaining information.
Once your job is complete, click on the Create Invoice button.
You'll get a list of invoices that have a blank status.

Select the invoice you'd like to create and click on Create Invoice.
If you'd like to create the email with the customer's email address and PDF attached, then click on Create email.
Note: Create email, does not send the email, it creates a draft email which you can customize further if you'd like
prior to sending (This works if you have Outlook)

4. Calculations for Dashboard

The "Calculations" tab contains the data preparation table for the chart on the "Dashboard" tab. 




