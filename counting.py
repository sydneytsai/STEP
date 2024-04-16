import gspread
from datetime import datetime

gc = gspread.oauth()

application_sheet = gc.open_by_key('1RgTqIQNlmTzLSNYf65uGF3iUaFgwCPPG9owc2GSU_TQ') 

application_worksheet = application_sheet.worksheet("DATABASE")

app_rows = application_worksheet.get_all_values()

students_serving = 0

date = datetime(2023, 7, 1, 0, 0, 0)

for i in range(2, len(app_rows)):
    print(app_rows[i])
    datetime_str = app_rows[i][0]
    datetime_object = datetime.strptime(datetime_str, '%m/%d/%Y %H:%M:%S')
    if (datetime_object < date):
        for item in app_rows[i]:
            if 'collected' in item.lower() or 'mailed' in item.lower():
                if ('chromebook' not in item.lower() and 'dell' not in item.lower()):
                    students_serving += 1
                break
    print(students_serving)



