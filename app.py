import os
import glob
from flask import render_template, url_for, Response, redirect, request, send_from_directory, abort, flash
from werkzeug.utils import secure_filename

from app_assets import create_app
from app_assets.forms import XlsxForm
from fixing_scripts import fix_xlsx

#App setup
app = create_app()
app.secret_key = os.getenv('SECRET_KEY', 'for dev') 

#File handling
UPLOAD_FOLDER = 'user_uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
absolute_upload_folder = os.path.abspath(UPLOAD_FOLDER)
print(absolute_upload_folder)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.errorhandler(500)
def not_found(error):
    return render_template('500.html', error=error)

@app.errorhandler(404)
def not_found(error):
    return render_template('404.html', error=error)


#Routes
@app.route('/', methods=['GET', 'POST'])
def index():
    files = glob.glob(f"{app.config['UPLOAD_FOLDER']}/*")
    for f in files:
        os.remove(f)

    xlsx_form = XlsxForm()
    context = {
        'xlsx_form' : xlsx_form
    }

    if request.method == 'POST':
        file = xlsx_form.file.data
        starting_cell = xlsx_form.starting_cell.data
        ending_cell = xlsx_form.ending_cell.data
        wb = xlsx_form.wb.data 

        if file.filename == '':
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename =  secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            result = fix_xlsx(UPLOAD_FOLDER + '/' + filename,starting_cell,ending_cell,wb)
            if result:
                flash(result, 'alert alert-danger')
                return redirect(url_for('index'))

            return redirect(url_for('uploaded_file',
                                    filename=filename))

    return render_template('index.html', **context)

@app.route('/uploads/<path:filename>')
def uploaded_file(filename):
    try:
        return send_from_directory(absolute_upload_folder, filename=filename, as_attachment=True)

    except FileNotFoundError:
        abort(404)

