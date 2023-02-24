# GST-Matching
A python script that matches the GST data from the portal and the company's own invoice data, and creates a sheet of matches, close matches, and no matches

### How it works?
- You need two CSV files: one for the data from the GST portal, and one for the company's own invoice data
- Enter the paths to those files as input
- The script will match your firm's invoice numbers to the invoice numbers on the portal
  - If they are an exact match, it matches the invoice amount and creates another column for the difference between portal amount and company's invoice amount
  - If they don't match, then it checks whether the date of the invoice is same and the amount is within some fixed buffer ('Close Match'), i.e., it tries to avoid any human error which might be there in creating the invoice. 
- As output, it generates an MS Excel file with two sheets: one for the invoices that are exact or close matches, and the other for the invoices that have no exact match and neither are they close matches.
