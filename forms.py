from flask_wtf import FlaskForm
from wtforms import StringField, EmailField, FileField, SubmitField, BooleanField
from wtforms.validators import DataRequired, Email, Length

class Form(FlaskForm):
    name = StringField("Name", validators=[DataRequired()])
    correlative = StringField("Correlativo", validators=[DataRequired()])
    submit = SubmitField("Enviar informe")