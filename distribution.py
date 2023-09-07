import gspread

gc = gspread.oauth()

distribution_sheet = gc.open_by_key('1nFSEMI7KSdQ7z4vVSrc-tyIBuUE8bdNo-hsOBEy-qB4')
application_sheet = gc.open_by_key('1RgTqIQNlmTzLSNYf65uGF3iUaFgwCPPG9owc2GSU_TQ') 

distribution_worksheet = distribution_sheet.get_worksheet(0)
application_worksheet = application_sheet.get_worksheet(0)

dist_rows = distribution_worksheet.get_all_values()
app_rows = application_worksheet.get_all_values()
row = 0

for dist_row in dist_rows:
    row += 1
    if (dist_row[7] == 'Pending'):
        email = dist_row[1]
        temp = application_worksheet.find(email, in_row=None, in_column=None, case_sensitive=False)
        app_row = app_rows[temp.row-1]
        collected = True
        for item in app_row:
            if 'Pending' in item:
                collected = False
        if collected:
            print(dist_row[1])
            distribution_worksheet.update_cell(row, 8, 'Picked Up')

print('Script is done running!')
