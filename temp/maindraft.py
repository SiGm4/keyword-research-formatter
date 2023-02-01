import glob
import sheetslib
from readcsvs import configure_csv_files
import csv
from modules.metadescriptions import meta_descriptions

# The ID of the main spreadsheet
spreadsheet_id = '1JRf0cA0--JJwvClNp9aVaQEilksYen6V8smOF9ncyjo'

csv_files = configure_csv_files()

if __name__ == '__main__':
    creds, service = sheetslib.authorize()

    meta_descriptions(csv_files, service, spreadsheet_id)    






## REFERENCE CODE

##    values = sheetslib.read_values(spreadsheet_id, "A1:C3", service)
##    
##    print('Data Stream starts:')
##    for row in values:
##        # Print columns A and E, which correspond to indices 0 and 4.
##        print(row)
##        print('%s, %s' % (row[0], row[4]))
##        
##    sheetslib.add_new_tab(spreadsheet_id, "Test", service)
##    sheetslib.update_values(spreadsheet_id, "A1:C2", "USER_ENTERED", [['A', 'B', 'L'],['E', 'D']], service)

##    with open(csv_files["Meta Descriptions Too Long"], 'r', encoding = "utf-8") as csv_file:
##        csv_reader = csv.reader(csv_file, delimiter=',')
##        print(csv_reader)
##        for row in csv_reader:
##            print(row)
