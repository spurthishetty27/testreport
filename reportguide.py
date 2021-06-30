from flask import Flask, json, request, jsonify
import os
import urllib.request
from werkzeug.utils import secure_filename
import pythoncom
import win32com.client as win32 
 
app = Flask(__name__)
 
#app.secret_key = "caircocoders-ednalan"
 
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
 
ALLOWED_EXTENSIONS = set(['docm','json'])
 
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
 
@app.route('/')
def main():
    return 'Homepage'
 
@app.route('/upload', methods=['POST'])
def upload_file():
    # check if the post request has the file part
    if 'files[]' not in request.files:
        resp = jsonify({'message' : 'No file part in the request'})
        resp.status_code = 400
        return resp
 
    files = request.files.getlist('files[]')
     
    errors = {}
    success = False
    word_file_received=0
    json_file_received=0
    for file in files:      
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            splitfilename=filename.split(".")
            if splitfilename[1]=="docm":
                cwd=os.getcwd()
                word_file_path=cwd+"\\static\\uploads\\"+filename
                wordfilename=filename
                word_file_received=1
            else:
                json_file_path=cwd+"\\static\\uploads\\"+filename
                json_file_received=1
        
            success = True
             
        else:
            errors[file.filename] = 'File type is not allowed'

    if (word_file_received==1 and json_file_received==1):
        #print(word_file_path)
        pythoncom.CoInitialize()
        wordApp = win32.gencache.EnsureDispatch('Word.Application') #create a word application object
        wordApp.Visible = False # hide the word application
        cwd=os.getcwd()
        doc = wordApp.Documents.Open(word_file_path)
        with open(json_file_path) as file:
            # Load its content and make a new dictionary
                data = json.load(file)
                for i in range(len(data["guide"])):
                    if "Check1" in data['guide'][i]:
                        descCheck1val=data['guide'][i]['Check1']
                        doc.FormFields("Check1").CheckBox.Value = descCheck1val
                        #print("y1")

                    if "Check2" in data['guide'][i]:
                        descCheck2val=data['guide'][i]['Check2']
                        doc.FormFields("Check2").CheckBox.Value = descCheck2val
                        #print("y2")

                    if "DclientName" in data['guide'][i]:
                        descClientName=data['guide'][i]['DclientName']
                        rng=doc.Bookmarks("DclientName").Range
                        rng.InsertAfter(descClientName)
                        #print("y3")

                    if "Engagement_Partner" in data['guide'][i]:
                        Eng_partnerName=data['guide'][i]['Engagement_Partner']
                        rng=doc.Bookmarks("Engagement_Partner").Range
                        rng.InsertAfter(Eng_partnerName)
                        #print("y4")

                    if "Engagment_draft_date" in data['guide'][i]:
                        Eng_draftDate=data['guide'][i]['Engagment_draft_date']
                        rng=doc.Bookmarks("Engagment_draft_date").Range
                        rng.InsertAfter(Eng_draftDate)
                        #print("y5")

                    if "Code" in data['guide'][i]:
                        Eng_Code=data['guide'][i]['Code']
                        rng=doc.Bookmarks("Code").Range
                        rng.InsertAfter(Eng_Code)
                        #print("y6")

                    if "Date" in data['guide'][i]:
                        Eng_date=data['guide'][i]['Date']
                        rng=doc.Bookmarks("Date").Range
                        rng.InsertAfter(Eng_date)
                        #print("y6")
                    
        
        updated_filename="updated_"+wordfilename
        path=cwd+"\\static\\uploads\\"+updated_filename

        isFile = os.path.isfile(path) 
        if isFile==False:
            #print("Enter")
            doc.SaveAs(path) 
            success=True

    if success and errors:
        errors['message'] = 'File(s) successfully uploaded'
        resp = jsonify(errors)
        resp.status_code = 500
        return resp
    if success:
        resp = jsonify({'message' : 'Report guide document successfully updated'})
        resp.status_code = 201
        return resp
    else:
        resp = jsonify(errors)
        resp.status_code = 500
        return resp


 
if __name__ == '__main__':
    app.run(debug=True)