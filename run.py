from flask import Flask, render_template, request, redirect
from werkzeug import secure_filename
from io import BytesIO
import olefile
import zlib
import re
import os
import string

ALLOWED_EXTENSIONS = set(['hwp', 'doc', 'xls'])
UPLOAD_FOLDER = 'C:\\Users\\hy00un\\Desktop\\aa\\hwp\\'

def hexdump(src, length=32):
    FILTER = ''.join([(len(repr(chr(x))) == 3) and chr(x) or '.' for x in range(256)])
    lines = []
    for c in xrange(0, len(src), length):
        chars = src[c:c+length]
        hex = ' '.join(["%02X" % ord(x) for x in chars])
        printable = ''.join(["%s" % ((ord(x) <= 127 and FILTER[ord(x)]) or '.') for x in chars])
        lines.append("%08X  %-*s  %s\n" % (c, length*3, hex, printable))
    return ''.join(lines)

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
        file = 'C:\\Users\\hy00un\\Desktop\\aa\\hwp\\'
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
        encode_data = stream.read()
        try:
            decode_data = hexdump(BytesIO(zlib.decompress(encode_data, -15)).read())
        except:
            decode_data = "zlib decode error!!"
        #print(decode_data)
        return render_template("result.html", title_split = title_split, filename = filename, ole_dir_list = ole_dir_list, encode_data = hexdump(encode_data), decode_data = decode_data)
    else:
        return redirect("/")

if __name__ == '__main__':
    app.run(host='0.0.0.0',debug=True)