from flask import Flask, render_template, jsonify, request, flash
from forms import Form
from report import generar_reporte, send_mail

#ENVIROMENT VARIABLES
from os import environ
from dotenv import load_dotenv, find_dotenv
import pandas as pd

# Error empty file
from pandas.errors import EmptyDataError


load_dotenv(find_dotenv())

secret_key = environ["SECRET_KEY"]
USER = environ['EMAIL']
PASSWORD = environ['PASSWORD']

app = Flask(__name__)
app.config['SECRET_KEY'] = secret_key

# Load emails
# emails
with open("lista_de_difusión.txt") as file:
    EMAILS = file.read().split()
EMAILS = ",".join(EMAILS)
print(EMAILS)
@app.route("/", methods=["POST", "GET"])
def index():
    form = Form()
    if request.method=="POST":
        name = form.name.data
        correlative = form.correlative.data
        data = request.files['filename']

        
        ## ESTO NO SE DEBE HACER PARA MULTIPLES PETICIONES
        try:
            excel, labor_days, year, month, message_1 = generar_reporte(data, number=correlative, name=name)
            
            subject = f"Informe de control de condiciones ambientales - {year} {month}"
            message = f"Se identificó las siguientes características:\n\t-Días laborables --> {labor_days}\n{message_1}\nEste informe fue realizado por: {name}\nCorrelativo: {correlative}"
            
            send_mail('reportes.lem@outlook.com',EMAILS,subject,message,excel,'smtp-mail.outlook.com', 587,username=USER,password=PASSWORD,isTls=True)
            flash("Reporte enviado con éxito", "sucess")
            #print("Ready :)")
        except EmptyDataError:
            flash("No has cargado el csv!", "danger")
        except Exception as e:
            print(e)
            flash("Ha ocurrido un error al enviar el correo. Contacta al administrador", 'danger')
        
    return render_template("index.html", form=form)

if __name__=="__main__":
    app.run(debug=True)