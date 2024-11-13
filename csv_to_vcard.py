import csv
import vobject
import os

def csv_to_vcard(csv_file, vcf_file='contacts.vcf'):
    try:
        with open(csv_file, mode='r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            
            # Open the vCard file for writing
            with open(vcf_file, 'w', encoding='utf-8') as vcf:
                for row in csv_reader:
                    # Create a combined Full Name field from First and Last Name
                    full_name = f"{row.get('First Name', '').strip()} {row.get('Last Name', '').strip()}".strip()
                    
                    # Skip if full name is missing
                    if not full_name:
                        print(f"Skipping entry: Missing full name for row {row}")
                        continue  # Skip entries without a full name
                    
                    # Create a new vCard
                    vcard = vobject.vCard()
                    
                    # Add full name
                    vcard.add('fn')
                    vcard.fn.value = full_name
                    
                    # Add email
                    if 'Email' in row and row['Email'].strip():
                        email = vcard.add('email')
                        email.value = row['Email']
                        email.type_param = 'INTERNET'
                    
                    # Add work phone
                    if 'Work Phone' in row and row['Work Phone'].strip():
                        work_phone = vcard.add('tel')
                        work_phone.value = row['Work Phone']
                        work_phone.type_param = 'WORK'
                    
                    # Add cell phone
                    if 'Cell Phone' in row and row['Cell Phone'].strip():
                        cell_phone = vcard.add('tel')
                        cell_phone.value = row['Cell Phone']
                        cell_phone.type_param = 'CELL'
                    
                    # Add company
                    if 'Company' in row and row['Company'].strip():
                        vcard.add('org')
                        vcard.org.value = [row['Company']]
                    
                    # Add address fields (City, State, Country)
                    address = vcard.add('adr')
                    address.value = vobject.vcard.Address(
                        city=row.get('City', '').strip(),
                        region=row.get('State', '').strip(),
                        country=row.get('Country', '').strip()
                    )
                    address.type_param = 'WORK'
                    
                    # Write the vCard to the file
                    vcf.write(vcard.serialize())
        
        print(f"Contacts from {csv_file} have been converted to {vcf_file}.")

    except FileNotFoundError:
        print("Error: CSV file not found. Please check the file name and try again.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Ask for the file name
csv_file = input("Enter the name of the CSV file (e.g., 'contacts.csv'): ")
if not os.path.exists(csv_file):
    print(f"The file {csv_file} does not exist. Please make sure it is in the same folder as this script.")
else:
    csv_to_vcard(csv_file)
