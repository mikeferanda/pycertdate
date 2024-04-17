# Required: pip install cryptography dnspython

import openpyxl
import ssl
from datetime import datetime, timezone
from cryptography import x509
from cryptography.hazmat.backends import default_backend
import dns.resolver
import socket

def has_ns_record(domain):
    try:
        dns.resolver.resolve(domain, 'NS')
        return True
    except dns.resolver.NoAnswer:
        return False
    except Exception as e:
        print(f"Error checking NS record for {domain}: {e}")
        return False

def get_certificate_expiration_date(url):
    try:
        cert = ssl.get_server_certificate((url, 443))
        cert_obj = x509.load_pem_x509_certificate(cert.encode(), default_backend())
        expiration_date = cert_obj.not_valid_after_utc
        expiration_date = expiration_date.replace(tzinfo=timezone.utc).astimezone(tz=None)  # Convert to timezone-aware datetime object
        return expiration_date
    except Exception as e:
        print(f"Error fetching certificate expiration date for {url}: {e}")
        return str(e)  # Return the error message as a string

def update_excel_with_certificate_expiration(file_path, url_column, py_date_column):
    try:
        wb = openpyxl.load_workbook(file_path)
    except PermissionError as pe:
        print(f"PermissionError: Unable to open the file {file_path}: {pe}")
        return

    sheet = wb.active

    for row in sheet.iter_rows(min_row=2, min_col=url_column, max_col=url_column):
        url = row[0].value
        if not url:  # Skip processing if the URL is blank
            continue
        try:
            expiration_date = get_certificate_expiration_date(url)
            if isinstance(expiration_date, datetime):
                # Convert datetime object to string in "mm/dd/yyyy" format for Excel compatibility
                expiration_date_str = expiration_date.strftime("%m/%d/%Y")
                sheet.cell(row=row[0].row, column=py_date_column).value = expiration_date_str
                print(f"Updated date for URL: {url}")
            else:
                sheet.cell(row=row[0].row, column=py_date_column).value = expiration_date
        except socket.gaierror:
            sheet.cell(row=row[0].row, column=py_date_column).value = "Contact Owner - Invalid"
        except Exception as e:
            print(f"Error processing {url}: {e}")
            sheet.cell(row=row[0].row, column=py_date_column).value = str(e)  # Put the error text in the column

    wb.save(file_path)
    wb.close()  # Close the workbook to release the file handle

if __name__ == "__main__":
    file_path = "pycertdate.xlsx"  # Change this to the path of your Excel file
    url_column = 1  # Column containing URLs (1-based index)
    py_date_column = 3  # Column where you want to update the expiration dates (1-based index)
    update_excel_with_certificate_expiration(file_path, url_column, py_date_column)




