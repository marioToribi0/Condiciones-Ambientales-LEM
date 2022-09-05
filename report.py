import pandas as pd
import openpyxl

import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
 
def generar_reporte(csv, number: str):
    data = pd.read_csv(csv, sep=";", header=3)
    def cleaned_date(x):
        values = x.split("/")
        return f"{int(values[0]):02d}/{int(values[1]):02d}/{int(values[2])}"

    def cleaned_time(x):
        values = x.split(":")
        return f"{int(values[0]):02d}:{int(values[1]):02d}:{int(values[2]):02d}"

    data.DATE = data.DATE.apply(lambda x: cleaned_date(x))
    data.TIME = data.TIME.apply(lambda x: cleaned_time(x))
    data.TIME = pd.to_datetime(data.DATE + ' ' +data.TIME, format='%d/%m/%Y %H:%M:%S')

    years = data.TIME.dt.year.value_counts()
    years = years.index.to_list()

    data = data.drop('DATE', axis=1)

    hour_9_to_1 = data[(data.TIME.dt.hour>=9) & (data.TIME.dt.hour<13)]
    hour_2_to_6 = data[(data.TIME.dt.hour>=14) & (data.TIME.dt.hour<18)]
    hour_7_to_10 = data[(data.TIME.dt.hour>=19) & (data.TIME.dt.hour<22)]

    # Monday to Sunday
    hour_2_to_6 = hour_2_to_6[hour_2_to_6.TIME.dt.dayofweek!=6]
    hour_9_to_1 = hour_9_to_1[hour_9_to_1.TIME.dt.dayofweek!=6]
    hour_7_to_10 = hour_7_to_10[hour_7_to_10.TIME.dt.dayofweek!=6]

    max_month = hour_2_to_6.groupby([hour_2_to_6.TIME.dt.month]).count()['TIME']
    max_month = max_month[max_month.values==max_month.max()].index[0]

    hour_9_to_1 = hour_9_to_1[hour_9_to_1.TIME.dt.month==max_month]
    hour_2_to_6 = hour_2_to_6[hour_2_to_6.TIME.dt.month==max_month]
    hour_7_to_10 = hour_7_to_10[hour_7_to_10.TIME.dt.month==max_month]

    hour_9_to_1 = hour_9_to_1.groupby([hour_9_to_1.TIME.dt.date]).apply(lambda x: x.sample(1))
    hour_2_to_6 = hour_2_to_6.groupby([hour_2_to_6.TIME.dt.date]).apply(lambda x: x.sample(1))
    hour_7_to_10 = hour_7_to_10.groupby([hour_7_to_10.TIME.dt.date]).apply(lambda x: x.sample(1))

    hour_7_to_10["TIME"] = list(map(lambda x: x[0], hour_7_to_10.index.to_list()))
    hour_2_to_6["TIME"] = list(map(lambda x: x[0], hour_2_to_6.index.to_list()))
    hour_9_to_1["TIME"] = list(map(lambda x: x[0], hour_9_to_1.index.to_list()))

    data = [hour_9_to_1, hour_7_to_10, hour_2_to_6]
    for el in data:
        el.drop(["DP","SN"], axis=1, inplace=True)
        
    days = []
    months_name = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    year = years[0]
    
    # Open base excel
    path = f"./static/static"
    wb_obj = openpyxl.load_workbook(f"{path}/base.xlsx") 
    month = months_name[max_month-1]

    sheet_obj = wb_obj['Mes']

    sheet_obj.cell(row=7, column=2).value = month
    
    bad_temperature = 0
    bad_humedity = 0
    
    labor_days = len(hour_2_to_6)

    for b in range(1, 4):
        df = data[b-1]
        day = list(filter(lambda x: x.month==max_month, map(lambda x: x[0], df.index.to_list())))
        # First update days
        days = list(map(lambda x: x.day, day))
        
        temperature = df[pd.to_datetime(df["TIME"]).dt.month==max_month]["oC"].to_list()
        humedity = df[pd.to_datetime(df["TIME"]).dt.month==max_month]["%RH"].to_list()
        
        bad_temperature += sum([1 if t<10 or t>=35 else 0 for t in temperature])
        bad_humedity += sum([1 if t<10 or t>=80 else 0 for t in humedity])
        
        # Add data to excel
        for i, day in enumerate(days):
            # Temperature
            cell_obj = sheet_obj.cell(row = day+10, column = b+1)
            cell_obj.value =  temperature[i]
            
            
            # Humedity
            cell_obj = sheet_obj.cell(row = day+10, column = b+21)
            cell_obj.value =  humedity[i]

    excel = f"./static/outputs/LEM-F-6.3-01-04 Informe de Control de Condiciones Ambientales v.8 {number}.xlsx"
    wb_obj.save(excel)
    
    return excel, labor_days, year, month, bad_temperature, bad_humedity

def send_mail(send_from,send_to,subject,text,excel,server,port,username='',password='',isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(excel, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{excel.split("/")[-1]}"')
    msg.attach(part)

    #SSL connection only working on Python 3+
    smtp = smtplib.SMTP(server, port)
    if isTls:
        smtp.starttls()

    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()
