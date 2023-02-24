import pandas as pd
import re
import os
import sys
from datetime import datetime
import time

def get_path(prompt):
    """
    Keeps prompting until the user enters a valid path. 
    Returns: The path that the user entered
    """
    while True:
        path = input(prompt)
        path.strip()
        if os.path.exists(path):
            break
        else:
            print("Please print a valid path.")
            continue
    return path

def get_paths():
    """Gets all the required paths to files"""
    prompts = ["Company Data Path: ", "Portal Data Path: ", "Output Folder Path: "]
    paths = []
    for prompt in prompts:
        paths.append(get_path(prompt))
    return paths

def get_buffer_size():
    while True:
        ans = input("Enter the buffer size for the invoice amount:")
        ans = ans.strip()
        try:
            ans = (int)(ans)
            break
        except ValueError:
            print("Please enter a valid integer")
            continue
    return ans

def build_output_path(folder, filename):
    """Creates a path to the output file from the directory and the filename"""
    suffix = '.xlsx'
    output_path = os.path.join(folder, filename + suffix)
    return output_path

def read_data(company_data_path, portal_data_path):
    """Loads the data from the path into a dataframe and returns the important columns"""

    company_dataframe = pd.DataFrame(pd.read_excel(company_data_path))
    portal_dataframe = pd.DataFrame(pd.read_excel(portal_data_path))

    # Get main columns, else show an error message
    try:
        company_dataframe = company_dataframe[['GSTIN of supplier', 'Party Name', 'Accounting Document No', 'Invoice No', 'Invoice Date', 'CGST Amount', "SGST Amount", "IGST Amount"]]
        company_dataframe["Invoice Date"] = pd.to_datetime(company_dataframe["Invoice Date"], format=r"%d-%m-%Y")

        portal_dataframe = portal_dataframe[["GSTIN of supplier", 'Invoice number', 'Invoice Date', 'Central Tax(₹)', 'State/UT Tax(₹)', "Integrated Tax(₹)"]]
        portal_dataframe["Invoice Date"] = pd.to_datetime(portal_dataframe["Invoice Date"], format=r"%d/%m/%Y")
    
    except KeyError as e:
        text = f"""
        Something went wrong with this column: {e.args[0]}. Please check the following and try again:
            1. The column headers are all in the first row, and there are no merged, multi-row headers
            2. Column names are as follows:
                a) GST Column in both sheets- GSTIN of supplier
                b) Invoice Number- "Invoice No" in Firm's sheet, and "Invoice number" in Portal sheet
                c) CGST- "CGST Amount" in Firm's sheet, and "Central Tax(₹)" in portal sheet
                d) SGST- "SGST Amount" in the Firm's sheet, and "State/UT Tax(₹)" in the portal sheet
                e) IGSt- "IGST Amount" in Firm's sheet, and "Integrated Tax(₹)" in portal sheet
                f) Name- "Party Name" in Firm's sheet
                g) Date- "Invoice Date" in both sheets
            """
        print(text)
        sys.exit(1)
    return company_dataframe, portal_dataframe

def init_data_dicts():
    '''Returns the dicts that are used to write the output file'''
    matched_D = {
        "GSTIN" : [],
        'Accounting Document No': [],
        'Party Name':[],
        "Invoice No" : [],
        "Invoice Date":[],
        "Firm Total" : [], 
        "Portal Total": [], 
        "Difference":[], 
        "Match Status":[], 
        "Portal Match":[]
    }

    unmatched_D = {
        "GSTIN" : [],
        "Party Name":[],
        "Invoice No" : [],
        "Invoice Date":[],
        "Firm Total":[]
    }
    return matched_D, unmatched_D

def filter_by_gstin(gstin, firm_df, portal_df):
    """Returns only the rows which have the same GSTIN"""
    firm_df = firm_df[["Invoice No", "Invoice Date", "Party Name", 'Accounting Document No', 'CGST Amount', 'SGST Amount', 'IGST Amount']].loc[firm_df['GSTIN of supplier'] == gstin]
    portal_df = portal_df[["Invoice number", "Invoice Date", "Central Tax(₹)", 'State/UT Tax(₹)', 'Integrated Tax(₹)']].loc[portal_df['GSTIN of supplier'] == gstin]
    return firm_df, portal_df

def clean_invoice(invoice):
    invoice = re.sub('[^0-9a-zA-Z]+', '', invoice)
    return invoice

def add_row_to_matched_dict(dict, gstin, party_name, accounting_doc_no, firm_invoice_no, firm_invoice_date, firm_total, portal_total, match_status, portal_match):
    dict["Invoice No"].append(firm_invoice_no)
    dict["Firm Total"].append(firm_total)
    dict["GSTIN"].append(gstin)
    dict["Portal Total"].append(portal_total)
    dict["Accounting Document No"].append(accounting_doc_no)
    dict["Party Name"].append(party_name)
    dict["Invoice Date"].append(firm_invoice_date)
    dict["Match Status"].append(match_status)
    dict["Portal Match"].append(portal_match)

    difference = float("{:.2f}".format(portal_total - firm_total))
    dict["Difference"].append(difference)
    return dict

def add_row_to_unmatched_dict(dict, gstin, party_name, firm_invoice_no, firm_invoice_date, firm_total):
    dict["Invoice No"].append(firm_invoice_no)
    dict["Party Name"].append(party_name)
    dict["Invoice Date"].append(firm_invoice_date)
    dict["Firm Total"].append(firm_total)
    dict["GSTIN"].append(gstin)
    return dict

def save_output_file(matched_dict, unmatched_dict, output_file_path):
    """Saves the matched and unmatched records to an output file"""
    matched_df = pd.DataFrame(matched_dict, columns=matched_dict.keys())
    unmatched_df = pd.DataFrame(unmatched_dict, columns=unmatched_dict.keys())
    try:
        writer = pd.ExcelWriter(output_file_path)
        matched_df.to_excel(writer, "Matched", index=False)
        unmatched_df.to_excel(writer, "Unmatched", index=False)
        writer.save()
    except PermissionError as e:
        print(f"There was an error: {e.args[0]}. The output file might be open. Please close it and try again.")
        sys.exit(1)

    return matched_df.shape[0], unmatched_df.shape[0]

def main():
    # Get/ Build paths to the files
    FIRM_DATA_PATH, PORTAL_DATA_PATH, output_folder = get_paths()
    filename = input("Filename: ")
    OUTPUT_PATH = build_output_path(output_folder, filename)

    buffer_size = get_buffer_size()
    # Read the data from the paths
    firm_df, PORTAL_DF = read_data(FIRM_DATA_PATH, PORTAL_DATA_PATH)
    
    # Init the output dicts
    matched_dict, unmatched_dict = init_data_dicts()
    
    # Get the unique GSTIN from Metaforge data
    unique_gstins_in_firm_sheet = firm_df['GSTIN of supplier'].unique()

    # Loop over all the unique GSTINs
    for gstin in unique_gstins_in_firm_sheet:

        # Get all the rows for the particular GSTIN
        NEW_firm_df, NEW_PORTAL_DF = filter_by_gstin(gstin, firm_df, PORTAL_DF)
        
        # Loop over the Metaforge Data
        for firm_index, firm_row in NEW_firm_df.iterrows():
            firm_invoice_date = firm_row["Invoice Date"]
            firm_invoice_date = firm_invoice_date.strftime(r"%d-%m-%Y")
            firm_invoice, firm_total = clean_invoice(str(firm_row["Invoice No"])), firm_row["CGST Amount"] + firm_row["SGST Amount"] + firm_row["IGST Amount"]
            is_matched = False

            # Match the invoice with the Portal Data
            for portal_index, portal_row in NEW_PORTAL_DF.iterrows():
                portal_invoice, portal_total = clean_invoice(str(portal_row["Invoice number"])), portal_row["Central Tax(₹)"] + portal_row['State/UT Tax(₹)'] + portal_row["Integrated Tax(₹)"]

                # If invoice number matched, add it to the matched data
                if firm_invoice == portal_invoice:
                    is_matched = True
                    matched_dict = add_row_to_matched_dict(matched_dict, gstin, firm_row["Party Name"] , firm_row["Accounting Document No"], firm_row["Invoice No"], firm_invoice_date, firm_total, portal_total, "Exact", portal_row["Invoice number"])
                    break
            
            # Find a close match, if no exact match
            is_close_match = False
            if not is_matched:
                for portal_index, portal_row in NEW_PORTAL_DF.iterrows():
                    portal_invoice_date, portal_total = portal_row["Invoice Date"].to_pydatetime(), portal_row["Central Tax(₹)"] + portal_row['State/UT Tax(₹)'] + portal_row["Integrated Tax(₹)"]
                    portal_invoice_date = portal_invoice_date.strftime(r"%d-%m-%Y")

                    # Check if the row is a close match based on the buffer size
                    if (firm_total - buffer_size <= portal_total <= firm_total + buffer_size) and (firm_invoice_date == portal_invoice_date):
                        is_close_match = True
                        matched_dict = add_row_to_matched_dict(matched_dict, gstin, firm_row["Party Name"] , firm_row["Accounting Document No"], firm_row["Invoice No"], firm_invoice_date, firm_total, portal_total, "Close", portal_row["Invoice number"])   

            # Add the unmatched data
            if not is_matched and not is_close_match:
                unmatched_dict = add_row_to_unmatched_dict(unmatched_dict, gstin, firm_row["Party Name"], firm_row["Invoice No"], firm_invoice_date, firm_total)
    
    # Save the data to the output file
    matched_rows, unmatched_rows = save_output_file(matched_dict, unmatched_dict, OUTPUT_PATH)
    print(f"""Result: Out of {firm_df.shape[0] - 1} rows,
            {matched_rows} rows matched/ close matched. {unmatched_rows} not matched""")
    
    # Give user time to read the text
    time.sleep(2)
    print("Successfully executed")
    time.sleep(2)

    sys.exit(0)

if __name__ =='__main__':
    main()         
