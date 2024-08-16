import openpyxl
import pandas as pd
import tkinter as tk
from PIL import ImageTk, Image
from tkinter import filedialog, messagebox,ttk,font
from tkinterdnd2 import TkinterDnD
from tkinter import Tk, Button, Frame
import tkinter.dnd as dnd
supplier_consolidated = []
supplier_columns = ['Batch ID','Import Action', 'Supplier Name', 'Supplier Name New', 'Supplier Number', 'Alternate Name', 'Tax Organization Type', 'Supplier Type', 'Inactive Date', 'Business Relationship', 'Parent Supplier', 'Alias', 'D-U-N-S Number', 'One-time supplier', 'Customer Number', 'SIC', 'National Insurance Number', 'Corporate Web Site', 'Chief Executive Title', 'Chief Executive Name', 'Business Classifications Not Applicable', 'Taxpayer Country', 'Taxpayer ID', 'Federal reportable', 'Federal Income Tax Type', 'State reportable', 'Tax Reporting Name', 'Name Control', 'Tax Verification Date', 'Use withholding tax', 'Withholding Tax Group', 'Vat Code', 'Tax Registration Number', 'Auto Tax Calc Override', 'Payment Method', 'Delivery Channel', 'Bank Instruction 1', 'Bank Instruction 2', 'Bank Instruction', 'Settlement Priority', 'Payment Text Message 1', 'Payment Text Message 2', 'Payment Text Message 3', 'Bank Charge Bearer', 'Payment Reason', 'Payment Reason Comments', 'Payment Format', 'ATTRIBUTE_CATEGORY', 'ATTRIBUTE1', 'ATTRIBUTE2', 'ATTRIBUTE3', 'ATTRIBUTE4', 'ATTRIBUTE5', 'ATTRIBUTE6', 'ATTRIBUTE7', 'ATTRIBUTE8', 'ATTRIBUTE9', 'ATTRIBUTE10', 'ATTRIBUTE11', 'ATTRIBUTE12', 'ATTRIBUTE13', 'ATTRIBUTE14', 'ATTRIBUTE15', 'ATTRIBUTE16', 'ATTRIBUTE17', 'ATTRIBUTE18', 'ATTRIBUTE19', 'ATTRIBUTE20', 'ATTRIBUTE_DATE1', 'ATTRIBUTE_DATE2', 'ATTRIBUTE_DATE3', 'ATTRIBUTE_DATE4', 'ATTRIBUTE_DATE5', 'ATTRIBUTE_DATE6', 'ATTRIBUTE_DATE7', 'ATTRIBUTE_DATE8', 'ATTRIBUTE_DATE9', 'ATTRIBUTE_DATE10', 'ATTRIBUTE_TIMESTAMP1', 'ATTRIBUTE_TIMESTAMP2', 'ATTRIBUTE_TIMESTAMP3', 'ATTRIBUTE_TIMESTAMP4', 'ATTRIBUTE_TIMESTAMP5', 'ATTRIBUTE_TIMESTAMP6', 'ATTRIBUTE_TIMESTAMP7', 'ATTRIBUTE_TIMESTAMP8', 'ATTRIBUTE_TIMESTAMP9', 'ATTRIBUTE_TIMESTAMP10', 'ATTRIBUTE_NUMBER1', 'ATTRIBUTE_NUMBER2', 'ATTRIBUTE_NUMBER3', 'ATTRIBUTE_NUMBER4', 'ATTRIBUTE_NUMBER5', 'ATTRIBUTE_NUMBER6', 'ATTRIBUTE_NUMBER7', 'ATTRIBUTE_NUMBER8', 'ATTRIBUTE_NUMBER9', 'ATTRIBUTE_NUMBER10', 'GLOBAL_ATTRIBUTE_CATEGORY', 'GLOBAL_ATTRIBUTE1', 'GLOBAL_ATTRIBUTE2', 'GLOBAL_ATTRIBUTE3', 'GLOBAL_ATTRIBUTE4', 'GLOBAL_ATTRIBUTE5', 'GLOBAL_ATTRIBUTE6', 'GLOBAL_ATTRIBUTE7', 'GLOBAL_ATTRIBUTE8', 'GLOBAL_ATTRIBUTE9', 'GLOBAL_ATTRIBUTE10', 'GLOBAL_ATTRIBUTE11', 'GLOBAL_ATTRIBUTE12', 'GLOBAL_ATTRIBUTE13', 'GLOBAL_ATTRIBUTE14', 'GLOBAL_ATTRIBUTE15', 'GLOBAL_ATTRIBUTE16', 'GLOBAL_ATTRIBUTE17', 'GLOBAL_ATTRIBUTE18', 'GLOBAL_ATTRIBUTE19', 'GLOBAL_ATTRIBUTE20', 'GLOBAL_ATTRIBUTE_DATE1', 'GLOBAL_ATTRIBUTE_DATE2', 'GLOBAL_ATTRIBUTE_DATE3', 'GLOBAL_ATTRIBUTE_DATE4', 'GLOBAL_ATTRIBUTE_DATE5', 'GLOBAL_ATTRIBUTE_DATE6', 'GLOBAL_ATTRIBUTE_DATE7', 'GLOBAL_ATTRIBUTE_DATE8', 'GLOBAL_ATTRIBUTE_DATE9', 'GLOBAL_ATTRIBUTE_DATE10', 'GLOBAL_ATTRIBUTE_TIMESTAMP1', 'GLOBAL_ATTRIBUTE_TIMESTAMP2', 'GLOBAL_ATTRIBUTE_TIMESTAMP3', 'GLOBAL_ATTRIBUTE_TIMESTAMP4', 'GLOBAL_ATTRIBUTE_TIMESTAMP5', 'GLOBAL_ATTRIBUTE_TIMESTAMP6', 'GLOBAL_ATTRIBUTE_TIMESTAMP7', 'GLOBAL_ATTRIBUTE_TIMESTAMP8', 'GLOBAL_ATTRIBUTE_TIMESTAMP9', 'GLOBAL_ATTRIBUTE_TIMESTAMP10', 'GLOBAL_ATTRIBUTE_NUMBER1', 'GLOBAL_ATTRIBUTE_NUMBER2', 'GLOBAL_ATTRIBUTE_NUMBER3', 'GLOBAL_ATTRIBUTE_NUMBER4', 'GLOBAL_ATTRIBUTE_NUMBER5', 'GLOBAL_ATTRIBUTE_NUMBER6', 'GLOBAL_ATTRIBUTE_NUMBER7', 'GLOBAL_ATTRIBUTE_NUMBER8', 'GLOBAL_ATTRIBUTE_NUMBER9', 'GLOBAL_ATTRIBUTE_NUMBER10', 'Registry ID', 'Payee Service Level', 'Pay Each Document Alone', 'Delivery Method', 'Remittance E-mail', 'Remittance Fax', 'DataFox ID']
supplier_address= ['Batch ID', 'Import Action', 'Supplier Name', 'Address Name ', 'Address Name New', 'Country', 'Address Line 1', 'Address Line 2', 'Address Line 3', 'Address Line 4', 'Phonetic Address Line', 'Address Element Attribute 1', 'Address Element Attribute 2', 'Address Element Attribute 3', 'Address Element Attribute 4', 'Address Element Attribute 5', 'Building', 'Floor Number', 'City', 'State', 'Province', 'County', 'Postal code', 'Postal Plus 4 code', 'Addressee', 'Global Location Number', 'Language', 'Inactive Date', 'Phone Country Code', 'Phone Area Code', 'Phone', 'Phone Extension', 'Fax Country Code', 'Fax Area Code', 'Fax', 'RFQ Or Bidding', 'Ordering', 'Pay', 'ATTRIBUTE_CATEGORY', 'ATTRIBUTE1', 'ATTRIBUTE2', 'ATTRIBUTE3', 'ATTRIBUTE4', 'ATTRIBUTE5', 'ATTRIBUTE6', 'ATTRIBUTE7', 'ATTRIBUTE8', 'ATTRIBUTE9', 'ATTRIBUTE10', 'ATTRIBUTE11', 'ATTRIBUTE12', 'ATTRIBUTE13', 'ATTRIBUTE14', 'ATTRIBUTE15', 'ATTRIBUTE16', 'ATTRIBUTE17', 'ATTRIBUTE18', 'ATTRIBUTE19', 'ATTRIBUTE20', 'ATTRIBUTE21', 'ATTRIBUTE22', 'ATTRIBUTE23', 'ATTRIBUTE24', 'ATTRIBUTE25', 'ATTRIBUTE26', 'ATTRIBUTE27', 'ATTRIBUTE28', 'ATTRIBUTE29', 'ATTRIBUTE30', 'ATTRIBUTE_NUMBER1', 'ATTRIBUTE_NUMBER2', 'ATTRIBUTE_NUMBER3', 'ATTRIBUTE_NUMBER4', 'ATTRIBUTE_NUMBER5', 'ATTRIBUTE_NUMBER6', 'ATTRIBUTE_NUMBER7', 'ATTRIBUTE_NUMBER8', 'ATTRIBUTE_NUMBER9', 'ATTRIBUTE_NUMBER10', 'ATTRIBUTE_NUMBER11', 'ATTRIBUTE_NUMBER12', 'ATTRIBUTE_DATE1', 'ATTRIBUTE_DATE2', 'ATTRIBUTE_DATE3', 'ATTRIBUTE_DATE4', 'ATTRIBUTE_DATE5', 'ATTRIBUTE_DATE6', 'ATTRIBUTE_DATE7', 'ATTRIBUTE_DATE8', 'ATTRIBUTE_DATE9', 'ATTRIBUTE_DATE10', 'ATTRIBUTE_DATE11', 'ATTRIBUTE_DATE12', 'E_Mail', 'Delivery Channel', 'Bank Instruction 1', 'Bank Instruction 2', 'Bank Instruction', 'Settlement Priority', 'Payment Text Message 1', 'Payment Text Message 2', 'Payment Text Message 3', 'Payee Service Level', 'Pay Each Document Alone', 'Bank Charge Bearer', 'Payment Reason', 'Payment Reason Comments', 'Delivery Method', 'Remittance E_Mail', 'Remittance Fax']
supplier_site_data = ['Batch ID', 'Import Action', 'Supplier Name', 'Procurement BU', 'Address Name ', 'Supplier Site', 'Supplier Site New', 'Inactive Date', 'Sourcing only', 'Purchasing', 'Procurement card', 'Pay', 'Primary Pay', 'Income tax reporting site', 'Alternate Site Name', 'Customer Number', 'B2B Communication Method', '\r\nB2B Supplier Site Code', 'Communication Method', 'E-Mail', 'Fax Country Code', 'Fax Area Code', 'Fax', 'Hold all new purchasing documents', 'Purchasing Hold Reason', 'Carrier', 'Mode of Transport', 'Service Level', 'Freight Terms', 'Pay on receipt', 'FOB', 'Country of Origin', 'Buyer Managed Transportation', 'Pay on use', 'Aging Onset Point', 'Aging Period Days', 'Consumption Advice Frequency', 'Consumption Advice Summary', 'Alternate Pay Site', 'Invoice Summary Level', 'Gapless invoice numbering', 'Selling Company Identifier', 'Create debit memo from return', 'Ship-to Exception Action ', 'Receipt Routing', 'Over-receipt Tolerance', 'Over-receipt Action', 'Early Receipt Tolerance in Days', 'Late Receipt Tolerance in Days', 'Allow Substitute Receipts', 'Allow unordered receipts', 'Receipt Date Exception', 'Invoice Currency', 'Invoice Amount Limit', 'Invoice Match Option', 'Match Approval Level', 'Payment Currency', 'Payment Priority', 'Pay Group', 'Quantity Tolerances', 'Amount Tolerance', 'Hold All Invoices', 'Hold Unmatched Invoices', 'Hold Unvalidated Invoices', 'Payment Hold By', 'Payment Hold Date', 'Payment Hold Reason', 'Payment Terms', 'Terms Date Basis', 'Pay Date Basis', 'Bank Charge Deduction Type', 'Always Take Discount', 'Exclude Freight From Discount', 'Exclude Tax From Discount', 'Create Interest Invoices', 'Vat Code-Obsoleted', 'Tax Registration Number-Obsoleted', 'Payment Method', 'Delivery Channel', 'Bank Instruction 1', 'Bank Instruction 2', 'Bank Instruction', 'Settlement Priority', 'Payment Text Message 1', 'Payment Text Message 2', 'Payment Text Message 3', 'Bank Charge Bearer', 'Payment Reason', 'Payment Reason Comments', 'Delivery Method', 'Remittance E-Mail', 'Remittance Fax', 'ATTRIBUTE_CATEGORY', 'ATTRIBUTE1', 'ATTRIBUTE2', 'ATTRIBUTE3', 'ATTRIBUTE4', 'ATTRIBUTE5', 'ATTRIBUTE6', 'ATTRIBUTE7', 'ATTRIBUTE8', 'ATTRIBUTE9', 'ATTRIBUTE10', 'ATTRIBUTE11', 'ATTRIBUTE12', 'ATTRIBUTE13', 'ATTRIBUTE14', 'ATTRIBUTE15', 'ATTRIBUTE16', 'ATTRIBUTE17', 'ATTRIBUTE18', 'ATTRIBUTE19', 'ATTRIBUTE20', 'ATTRIBUTE_DATE1', 'ATTRIBUTE_DATE2', 'ATTRIBUTE_DATE3', 'ATTRIBUTE_DATE4', 'ATTRIBUTE_DATE5', 'ATTRIBUTE_DATE6', 'ATTRIBUTE_DATE7', 'ATTRIBUTE_DATE8', 'ATTRIBUTE_DATE9', 'ATTRIBUTE_DATE10', 'ATTRIBUTE_TIMESTAMP1', 'ATTRIBUTE_TIMESTAMP2', 'ATTRIBUTE_TIMESTAMP3', 'ATTRIBUTE_TIMESTAMP4', 'ATTRIBUTE_TIMESTAMP5', 'ATTRIBUTE_TIMESTAMP6', 'ATTRIBUTE_TIMESTAMP7', 'ATTRIBUTE_TIMESTAMP8', 'ATTRIBUTE_TIMESTAMP9', 'ATTRIBUTE_TIMESTAMP10', 'ATTRIBUTE_NUMBER1', 'ATTRIBUTE_NUMBER2', 'ATTRIBUTE_NUMBER3', 'ATTRIBUTE_NUMBER4', 'ATTRIBUTE_NUMBER5', 'ATTRIBUTE_NUMBER6', 'ATTRIBUTE_NUMBER7', 'ATTRIBUTE_NUMBER8', 'ATTRIBUTE_NUMBER9', 'ATTRIBUTE_NUMBER10', 'GLOBAL_ATTRIBUTE_CATEGORY', 'GLOBAL_ATTRIBUTE1', 'GLOBAL_ATTRIBUTE2', 'GLOBAL_ATTRIBUTE3', 'GLOBAL_ATTRIBUTE4', 'GLOBAL_ATTRIBUTE5', 'GLOBAL_ATTRIBUTE6', 'GLOBAL_ATTRIBUTE7', 'GLOBAL_ATTRIBUTE8', 'GLOBAL_ATTRIBUTE9', 'GLOBAL_ATTRIBUTE10', 'GLOBAL_ATTRIBUTE11', 'GLOBAL_ATTRIBUTE12', 'GLOBAL_ATTRIBUTE13', 'GLOBAL_ATTRIBUTE14', 'GLOBAL_ATTRIBUTE15', 'GLOBAL_ATTRIBUTE16', 'GLOBAL_ATTRIBUTE17', 'GLOBAL_ATTRIBUTE18', 'GLOBAL_ATTRIBUTE19', 'GLOBAL_ATTRIBUTE20', 'GLOBAL_ATTRIBUTE_DATE1', 'GLOBAL_ATTRIBUTE_DATE2', 'GLOBAL_ATTRIBUTE_DATE3', 'GLOBAL_ATTRIBUTE_DATE4', 'GLOBAL_ATTRIBUTE_DATE5', 'GLOBAL_ATTRIBUTE_DATE6', 'GLOBAL_ATTRIBUTE_DATE7', 'GLOBAL_ATTRIBUTE_DATE8', 'GLOBAL_ATTRIBUTE_DATE9', 'GLOBAL_ATTRIBUTE_DATE10', 'GLOBAL_ATTRIBUTE_TIMESTAMP1', 'GLOBAL_ATTRIBUTE_TIMESTAMP2', 'GLOBAL_ATTRIBUTE_TIMESTAMP3', 'GLOBAL_ATTRIBUTE_TIMESTAMP4', 'GLOBAL_ATTRIBUTE_TIMESTAMP5', 'GLOBAL_ATTRIBUTE_TIMESTAMP6', 'GLOBAL_ATTRIBUTE_TIMESTAMP7', 'GLOBAL_ATTRIBUTE_TIMESTAMP8', 'GLOBAL_ATTRIBUTE_TIMESTAMP9', 'GLOBAL_ATTRIBUTE_TIMESTAMP10', 'GLOBAL_ATTRIBUTE_NUMBER1', 'GLOBAL_ATTRIBUTE_NUMBER2', 'GLOBAL_ATTRIBUTE_NUMBER3', 'GLOBAL_ATTRIBUTE_NUMBER4', 'GLOBAL_ATTRIBUTE_NUMBER5', 'GLOBAL_ATTRIBUTE_NUMBER6', 'GLOBAL_ATTRIBUTE_NUMBER7', 'GLOBAL_ATTRIBUTE_NUMBER8', 'GLOBAL_ATTRIBUTE_NUMBER9', 'GLOBAL_ATTRIBUTE_NUMBER10', 'Required Acknowledgement', 'Acknowledge Within Days', 'Invoice Channel', 'Payee Service Level', 'Pay Each Document Alone']
supplier_third_party_relationship =['Batch ID', 'Import Action', 'Supplier Name', 'Supplier Site', 'Procurement BU', 'Default', 'Remit-to Supplier', 'Address Name', 'From Date', 'To Date', 'Description']
supplier_site_assignment = ['Batch ID', 'Import Action', 'Supplier Name', 'Supplier Site', 'Procurement BU', 'Client BU', 'Bill-to BU', 'Ship-to Location', 'Bill-to Location', 'Use Withholding Tax', 'Withholding Tax Group', 'Liability Distribution', 'Prepayment Distribution', 'Bills Payable Distribution', 'Distribution Set', 'Inactive Date']
supplier_contact=['Batch ID', 'Import Action', 'Supplier Name', 'Prefix', 'First Name', 'First Name New', 'Middle Name', 'Last Name', 'Last Name New', 'Job Title', 'Administrative Contact', 'E-Mail', 'E-Mail New', 'Phone Country Code', 'Phone Area Code', 'Phone', 'Phone Extension', 'Fax Country Code', 'Fax Area Code', 'Fax', 'Mobile Country Code', 'Mobile Area Code', 'Mobile', 'Inactive Date', 'ATTRIBUTE_CATEGORY', 'ATTRIBUTE1', 'ATTRIBUTE2', 'ATTRIBUTE3', 'ATTRIBUTE4', 'ATTRIBUTE5', 'ATTRIBUTE6', 'ATTRIBUTE7', 'ATTRIBUTE8', 'ATTRIBUTE9', 'ATTRIBUTE10', 'ATTRIBUTE11', 'ATTRIBUTE12', 'ATTRIBUTE13', 'ATTRIBUTE14', 'ATTRIBUTE15', 'ATTRIBUTE16', 'ATTRIBUTE17', 'ATTRIBUTE18', 'ATTRIBUTE19', 'ATTRIBUTE20', 'ATTRIBUTE21', 'ATTRIBUTE22', 'ATTRIBUTE23', 'ATTRIBUTE24', 'ATTRIBUTE25', 'ATTRIBUTE26', 'ATTRIBUTE27', 'ATTRIBUTE28', 'ATTRIBUTE29', 'ATTRIBUTE30', 'ATTRIBUTE_NUMBER1', 'ATTRIBUTE_NUMBER2', 'ATTRIBUTE_NUMBER3', 'ATTRIBUTE_NUMBER4', 'ATTRIBUTE_NUMBER5', 'ATTRIBUTE_NUMBER6', 'ATTRIBUTE_NUMBER7', 'ATTRIBUTE_NUMBER8', 'ATTRIBUTE_NUMBER9', 'ATTRIBUTE_NUMBER10', 'ATTRIBUTE_NUMBER11', 'ATTRIBUTE_NUMBER12', 'ATTRIBUTE_DATE1', 'ATTRIBUTE_DATE2', 'ATTRIBUTE_DATE3', 'ATTRIBUTE_DATE4', 'ATTRIBUTE_DATE5', 'ATTRIBUTE_DATE6', 'ATTRIBUTE_DATE7', 'ATTRIBUTE_DATE8', 'ATTRIBUTE_DATE9', 'ATTRIBUTE_DATE10', 'ATTRIBUTE_DATE11', 'ATTRIBUTE_DATE12', 'User Account Action', 'Role 1', 'Role 2', 'Role 3', 'Role 4', 'Role 5']
supplier_contact_address =  ['Batch ID', 'Import Action', 'Supplier Name', 'Address Name', 'First Name', 'Last Name', 'E-Mail']
supplier_profile_attachment =  ['Batch ID', 'Import Action', 'Supplier Name', 'Category', 'Type', 'File/Text/URL', 'File Attachments .ZIP', 'Title', 'Description']
supplier_site_attachment =  ['Batch ID', 'Import Action', 'Supplier Name', 'Procurement BU', 'Supplier Site', 'Category', 'Type', 'File/Text/URL', 'File Attachments .ZIP', 'Title', 'Description']
supplier_business_class_attachment =  ['Batch ID', 'Import Action', 'Supplier Name', 'Classification', 'Subclassification', 'Certifying Agency', 'Certificate Number', 'Category', 'Type', 'File/Text/URL', 'File Attachments .ZIP', 'Title', 'Description']
supplier_business_classification =  ['Batch ID', 'Import Action', 'Supplier Name', 'Classification', 'Classification New', 'Subclassification', 'Certifying Agency', 'Certifying Agency New', 'Create Certifying Agency', 'Certificate Number', 'Certificate Number New', 'Start Date', 'Expiration Date', 'Notes', 'Provided By First Name', 'Provided By Last Name', 'Provided By E-Mail', 'Confirmed On']
supplier_product_and_service_category =  ['Batch ID', 'Import Action', 'Supplier Name', 'Category Type', 'Category Name']
supplier_payee =  ['Import Batch Identifier', 'Payee Identifier', 'Business Unit Name', 'Supplier Number', 'Supplier Site', 'Pay Each Document Alone', 'Payment Method Code', 'Delivery Channel Code', 'Settlement Priority', 'Remit Delivery Method', 'Remit Advice Email', 'Remit Advice Fax', 'Bank Instructions 1', 'Bank Instructions 2', 'Bank Instruction Details', 'Payment Reason Code', 'Payment Reason Comments', 'Payment Message1', 'Payment Message2', 'Payment Message3', 'Bank Charge Bearer Code']
supplier_bank_accounts =  ['Import Batch Identifier', 'Payee Identifier', 'Payee Bank Account Identifier', 'Bank Name', 'Branch Name', 'Account Country Code', 'Account Name', 'Account Number', 'Account Currency Code', 'Allow International Payments', 'Account Start Date', 'Account End Date', 'IBAN', 'Check Digits', 'Account Alternate Name', 'Account Type Code', 'Account Suffix', 'Account Description', 'Agency Location Code', 'Exchange Rate Agreement Number', 'Exchange Rate Agreement Type', 'Exchange Rate', 'Secondary Account Reference', 'Attribute Category', 'Attribute 1', 'Attribute 2', 'Attribute 3', 'Attribute 4', 'Attribute 5', 'Attribute 6', 'Attribute 7', 'Attribute 8', 'Attribute 9', 'Attribute 10', 'Attribute 11', 'Attribute 12', 'Attribute 13', 'Attribute 14', 'Attribute 15']
supplier_bank_account_assignment =  ['Import Batch Identifier', 'Payee Identifier', 'Payee Bank Account Identifier', 'Payee Bank Account Assignment Identifier', 'Primary Flag', 'Account Assignment Start Date', 'Account Assignment End Date']
project_details = ['Project Number', 'Project Name', 'Project ORG']
project_task_detail = ['Project Number', 'Project Name', 'Task name']
data_list =[]