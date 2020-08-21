from flask_wtf import FlaskForm
from wtforms.fields import FileField, StringField, SubmitField
from wtforms.validators import DataRequired

class XlsxForm(FlaskForm):
    file = FileField('Archivo a arreglar', validators=[DataRequired()])
    starting_cell = StringField('Celda donde empezar', validators=[DataRequired()])
    ending_cell = StringField('Celda donde acabar', validators=[DataRequired()])
    wb = StringField('Hoja a revisar', validators=[DataRequired()])
    submit = SubmitField('Enviar!')

    