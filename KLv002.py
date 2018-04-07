import pandas as pd
import openpyxl
from os import listdir
from functools import reduce

print('Initializing libs...')

#find excels in folder
def find_xl_filenames( path_to_dir, suffix=".xlsx" ):
    filenames = listdir(path_to_dir)
    return [ filename for filename in filenames if filename.endswith( suffix ) ]

excel_names = find_xl_filenames('.', suffix = 'xlsx')
print('Найдено ТКП: {}'.format (len(excel_names)))
# read them in
excel_names = [pd.ExcelFile(name) for name in excel_names]

# turn them into dataframes
frames = [x.parse(x.sheet_names[0]) for x in excel_names]
#print(frames)
# delete the first row for all frames except the first
#frames[1:] = [df[1:] for df in frames[1:]]
#frames = list(frames)

# concatenate them
#combined = pd.concat(frames)
print('Формируем КЛ')
result = pd.concat(frames, axis=1, join='inner')
#result = frames.merge(frames, on='Защищенные поля')
output = result = reduce(lambda left,right: pd.merge(left,right,on='Защищенные поля'), frames)
#result.drop(columns='A', axis = 1)

writer = pd.ExcelWriter('Output.xlsx')

output.to_excel(writer,'КЛ')
writer.save()
newFile = "output.xlsx"

wb = openpyxl.load_workbook(filename = newFile)
worksheet = wb.active
for col in worksheet.columns:
    max_length = 0
    column = col[0].column # Get the column name
    for cell in col:
        if cell.coordinate in worksheet.merged_cells: # not check merge_cells
            continue
        try: # Necessary to avoid error on empty cells
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2) * 1.2
    worksheet.column_dimensions[column].width = adjusted_width

wb.save('КЛ.xlsx')
wb.close()


#sending the email

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

print('Отправляем письмо')

fromaddr = "paul.golubev@gmail.com"
toaddr = "paul.golubev@gmail.com"

msg = MIMEMultipart()

msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "Testing"

body = "I am a robot"

msg.attach(MIMEText(body, 'plain'))

filename = "sformatted_output.xlsx"
attachment = open("КЛ.xlsx", "rb")

part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

msg.attach(part)

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(fromaddr, "password")
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
server.quit()

print('Письмо отпарвлено!')