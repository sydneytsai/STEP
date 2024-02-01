import gspread

gc = gspread.oauth()

distribution_sheet = gc.open_by_key('1nFSEMI7KSdQ7z4vVSrc-tyIBuUE8bdNo-hsOBEy-qB4')
application_sheet = gc.open_by_key('1RgTqIQNlmTzLSNYf65uGF3iUaFgwCPPG9owc2GSU_TQ') 

distribution_worksheet = distribution_sheet.worksheet("Spring 2024")
application_worksheet = application_sheet.worksheet("DATABASE")

dist_rows = distribution_worksheet.get_all_values()
app_rows = application_worksheet.get_all_values()
row = 0

for dist_row in dist_rows:
    row += 1
    if (dist_row[5] == 'Pending'):
        email = dist_row[0]
        temp = application_worksheet.find(email, in_row=None, in_column=None, case_sensitive=False)
        if (temp):
            app_row = app_rows[temp.row-1]
            collected = False
            for item in app_row:
                if 'picked up' in item.lower():
                    collected = True
                if 'pending' in item.lower():
                    collected = False
                    break
                if 'does not need' in item.lower() or 'no longer needed' in item.lower():
                    collected = False
                    break
                    print(f"Student {email} does not need an item")
            if collected:
                print(email)
                distribution_worksheet.update_cell(row, 6, 'Picked Up') # 1 + row number from line 17
        else:
            print(f"Error {email} not found in application sheet.")

print('Script is done running!')
