import csv
import sheetslib
import sheetsformatlib

def meta_description_file_reader(csv_files, final_values, key, issue):
    with open(csv_files[key], 'r', encoding = "utf-8") as csv_file:
        csv_reader = csv.DictReader(csv_file)
        row_count = 0
        for row in csv_reader:
            final_values.append([row["URL"], issue, False, row["Meta description"], "", False])
            row_count += 1
    return final_values, row_count

def meta_descriptions(csv_files, service, spreadsheet_id):

    sheet_id = sheetslib.create_sheet_or_get_sheet_id(spreadsheet_id, "Meta Description Issues", service)

    final_values = [["URL","Issue","Fixed","Current Meta Description","Proposed Meta Description","Approved"]]
    row_counts = []

    final_values, row_count = meta_description_file_reader(csv_files, final_values, "Meta Descriptions Too Long", "Too Long") #TODO Turn these keys into config file variables
    final_values.append([])
    row_counts.append(row_count)
    
    final_values, row_count = meta_description_file_reader(csv_files, final_values, "Meta Descriptions Too Short", "Too Short")
    final_values.append([])
    row_counts.append(row_count)
    
    final_values, row_count = meta_description_file_reader(csv_files, final_values, "Meta Descriptions Tag Missing", "Missing")
    row_counts.append(row_count)
    
    writable_range = "Meta Description Issues!A1:" + chr(64+len(final_values[0])) + str(len(final_values))
    sheetslib.update_values(spreadsheet_id, writable_range, "USER_ENTERED", final_values, service)

    meta_descriptions_formatting(service, spreadsheet_id, sheet_id, row_counts)

def meta_descriptions_formatting(service, spreadsheet_id, sheet_id, row_counts):
    for i in range (len(row_counts)):
        sheetsformatlib.add_checkboxes_with_color(spreadsheet_id, sheet_id, service, {
                            'startRowIndex': (i+1)+sum(row_counts[:i]),
                            'endRowIndex': (i+1)+sum(row_counts[:i+1]),
                            'startColumnIndex': 2,
                            'endColumnIndex': 3})
        sheetsformatlib.add_checkboxes_with_color(spreadsheet_id, sheet_id, service, {
                            'startRowIndex': (i+1)+sum(row_counts[:i]),
                            'endRowIndex': (i+1)+sum(row_counts[:i+1]),
                            'startColumnIndex': 5,
                            'endColumnIndex': 6})
    for i in range (len(row_counts)-1):
        sheetsformatlib.color_row(spreadsheet_id, sheet_id, service, (i+2)+sum(row_counts[:i+1]))

    sheetsformatlib.align_vertical_middle_all(spreadsheet_id, sheet_id, service)
    sheetsformatlib.bold_center_header_row(spreadsheet_id, sheet_id, service)
    sheetsformatlib.center_columns(spreadsheet_id, sheet_id, service, 1, 2)
    sheetsformatlib.color_header_row(spreadsheet_id, sheet_id, service, 6)
    sheetsformatlib.auto_resize_column(spreadsheet_id, sheet_id, service, 1)
    sheetsformatlib.resize_columns(spreadsheet_id, sheet_id, service, 400, 3, 5)
    sheetsformatlib.set_columns_wrap_strategy(spreadsheet_id, sheet_id, service, "WRAP", 3, 5)
