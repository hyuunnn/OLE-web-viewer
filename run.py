from flask import Flask, render_template, request, redirect
from werkzeug import secure_filename
from io import BytesIO
import olefile
import zlib
import re
import os

ALLOWED_EXTENSIONS = set(['hwp', 'doc', 'xls'])
UPLOAD_FOLDER = '/root/my_project/hwp'

def allowed_file(filename):
    return '.' in filename and \
            filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.errorhandler(404)
def page_not_found(error):
    return render_template("page_not_found.html", error = "404 error!!"), 404

@app.route('/')
def main():
    return render_template("index.html")

@app.route('/upload',methods=['POST','GET'])
def upload():
    if request.method == "POST":
        file = request.files['file']
    else:
	file = '/root/my_project/hwp'
    title = request.args.get('title')
    print(title)
    filename = ""
    if file:
	if request.method == "POST":
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
	else:
	    filename = request.args.get("filename")
	ole = olefile.OleFileIO(UPLOAD_FOLDER+"/"+filename)
        ole_dir = ole.listdir()
	ole_dir_list = []
	encode_data = ""
	decode_data = ""
	title_split = ""
	for i in range(len(ole_dir)):
	    ole_dir_list.append("/".join(ole_dir[i]))
	    if "/" in ole_dir_list[i]:
	        if ole_dir_list[i].split("/")[len(ole_dir_list[i].split("/"))-1] == title:
	            title = "/".join(ole_dir[i])
		    title_split = ole_dir_list[i].split("/")[len(ole_dir_list[i].split("/"))-1]
	    elif ole_dir_list[i] == title:
		title = ole_dir[i]
		title_split = ole_dir_list[i]
	if title == None:
	    if "/" in ole_dir_list[0]:
	        title = ole_dir_list[0].split("/")[len(ole_dir_list[0].split("/"))-1]
	    else:
		title = ole_dir_list[0]
	stream = ole.openstream(title)
        encode_data = stream.read().encode("hex")
        #decode_data = BytesIO(zlib.decompress(stream.read(), -15))
	#print(decode_data)
	return render_template("result.html", title_split = title_split, filename = filename, ole_dir_list = ole_dir_list,encode_data = encode_data, decode_data = decode_data)
    else:
        return redirect("/")

if __name__ == '__main__':
    app.run(host='0.0.0.0',debug=True)
