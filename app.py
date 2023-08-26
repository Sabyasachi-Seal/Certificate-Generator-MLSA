# importing the required libraries
import os
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from main_certificate import main

# initialising the flask app
app = Flask(__name__)

# Creating the upload folder
upload_folder = "Data/"
if not os.path.exists(upload_folder):
    os.mkdir(upload_folder)

# Max size of the file
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024

# Configuring the upload folder
app.config['UPLOAD_FOLDER'] = upload_folder

# configuring the allowed extensions
allowed_extensions = ['csv']


def check_file_extension(filename):
    return filename.split('.')[-1] in allowed_extensions

# The path for uploading the file


@app.route('/')
def upload_file():
    return render_template('upload.html')


@app.route('/upload', methods=['GET', 'POST'])
def uploadfile():
    if request.method == 'POST':  # check if the method is post
        f = request.files['Participant_List']  # get the file from the files object
        event = request.form['Event_Name']
        name = request.form['Ambassador_Name']
        # Saving the file in the required destination
        if check_file_extension(f.filename):
            newfilename = secure_filename(f.filename)
            abs_path = os.path.join(app.config['UPLOAD_FOLDER'], newfilename)
            # this will secure the file
            f.save(abs_path)
            main(newfilename, event, name)
            os.remove(abs_path)
            return 'file uploaded successfully'  # Display thsi message after uploading
        else:
            return 'The file extension is not allowed'


if __name__ == '__main__':
    app.run()  # running the flask app
