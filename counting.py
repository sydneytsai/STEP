import gspread
from datetime import datetime

gc = gspread.oauth()

application_sheet = gc.open_by_key('1RgTqIQNlmTzLSNYf65uGF3iUaFgwCPPG9owc2GSU_TQ') 

application_worksheet = application_sheet.worksheet("DATABASE")

app_rows = application_worksheet.get_all_values()

students_serving = 0
apple_adapter = 0
belkin_hub = 0
drawing_tablet = 0
headphones = 0
hotspots = 0
iclicker = 0
keyboard = 0
laptops = 0
microphone = 0
mouse = 0
webcam = 0
extender = 0
yubikey = 0


date = datetime(2023, 7, 1, 0, 0, 0)

for i in range(2, len(app_rows)):
    if (('picked up' in app_rows[i][55].lower() or 'mailed' in app_rows[i][55].lower() or 'collected' in app_rows[i][55].lower()) and 'dell' not in app_rows[i][55].lower() and 'chromebook' not in app_rows[i][55].lower()):
        laptops += 1
    if ('picked up' in app_rows[i][60].lower() or 'mailed' in app_rows[i][60].lower()):
        hotspots += 1
    if ('picked up' in app_rows[i][64].lower() or 'mailed' in app_rows[i][64].lower()):
        drawing_tablet += 1
    if ('picked up' in app_rows[i][67].lower() or 'mailed' in app_rows[i][67].lower()):
        headphones += 1
    if ('picked up' in app_rows[i][68].lower() or 'mailed' in app_rows[i][68].lower()):
        microphone += 1
    if ('picked up' in app_rows[i][69].lower() or 'mailed' in app_rows[i][69].lower()):
        webcam += 1
    if ('picked up' in app_rows[i][70].lower() or 'mailed' in app_rows[i][70].lower()):
        iclicker += 1
    if ('picked up' in app_rows[i][72].lower() or 'mailed' in app_rows[i][72].lower()):
        belkin_hub += 1
    if ('picked up' in app_rows[i][73].lower() or 'mailed' in app_rows[i][73].lower()):
        apple_adapter += 1
    if ('picked up' in app_rows[i][74].lower() or 'mailed' in app_rows[i][74].lower()):
        mouse += 1
    if ('picked up' in app_rows[i][75].lower() or 'mailed' in app_rows[i][75].lower()):
        keyboard += 1
    if ('picked up' in app_rows[i][76].lower() or 'mailed' in app_rows[i][76].lower()):
        extender += 1
    if ('picked up' in app_rows[i][77].lower() or 'mailed' in app_rows[i][77].lower()):
        yubikey += 1
    datetime_str = app_rows[i][0]
    datetime_object = datetime.strptime(datetime_str, '%m/%d/%Y %H:%M:%S')
    if (datetime_object < date):
        for item in app_rows[i]:
            if (('collected' in item.lower() or 'mailed' in item.lower() or 'picked up' in item.lower()) and ('pending' not in item.lower()) and ('chromebook' not in item.lower() and 'dell' not in item.lower())):
                students_serving += 1
                break

print(f"students serving pre 7/1/23 {students_serving}")
print(f"keyboard {keyboard}")
print(f"apple_adapter {apple_adapter}")
print(f"belkin_hub {belkin_hub}")
print(f"drawing_tablet {drawing_tablet}")
print(f"headphones {headphones}")
print(f"hotspots {hotspots}")
print(f"iclicker {iclicker}")
print(f"laptops {laptops}")
print(f"microphone {microphone}")
print(f"mouse {mouse}")
print(f"webcam {webcam}")
print(f"extender {extender}")
print(f"yubikey {yubikey}")



