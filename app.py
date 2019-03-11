# -*- coding: utf-8 -*-
"""
Created on 2019-03-09
成绩分析程序
@author: hujili
"""
import os, datetime, random
from werkzeug.utils import secure_filename
from util.socre import analysis
from flask import Flask, request, jsonify, render_template, send_from_directory, url_for
from urllib import parse

#文件上传存放的文件夹, 值为非绝对路径时，相对于项目根目录
IMAGE_FOLDER  = 'static/upload/'
#生成无重复随机数
gen_rnd_filename = lambda :"%s%s" %(datetime.datetime.now().strftime('%Y%m%d%H%M%S'), str(random.randrange(1000, 10000)))
#文件名合法性验证
allowed_file = lambda filename: '.' in filename and filename.rsplit('.', 1)[1] in set(['xls', 'xlsx'])

app = Flask(__name__)
app.config.update(
    SECRET_KEY = os.urandom(24),
    # 上传文件夹
    UPLOAD_FOLDER = os.path.join(app.root_path, IMAGE_FOLDER),
    # 最大上传大小，当前2MB
    MAX_CONTENT_LENGTH = 2 * 1024 * 1024
)

@app.route("/")
def index_view():
    """主页视图"""
    return render_template("index.html")

@app.route('/showdoc/<filename>')
def showimg_view(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/upload/', methods=['POST','OPTIONS'])
def upload_view():
    res = dict(code=-1, msg=None)
    f = request.files.get('file')
    if f and allowed_file(f.filename):
        filename = secure_filename(gen_rnd_filename() + "." + f.filename.split('.')[-1]) #随机命名
        # 自动创建上传文件夹
        if not os.path.exists(app.config['UPLOAD_FOLDER']):
            os.makedirs(app.config['UPLOAD_FOLDER'])
        # 保存文件
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        print("filename="+f.filename)
        # 开始分析
        res.update(analysis(f,os.path.join(app.config['UPLOAD_FOLDER'], '成绩分析表模板.docx'),os.path.join(app.config['UPLOAD_FOLDER'], filename)))
        if res.get("code") != -1:
            xlsUrl = url_for('showimg_view', filename=filename.split('.')[0]+ "." + "docx", _external=True)
            #调用微软在线预览office
            office = {"src": "http://47.92.25.35:8092/static/upload/"+filename.split('.')[0]+ "." + "docx"}
            print(parse.urlencode(office))
            res.update(code=0, data=dict(src=xlsUrl,docsrc=parse.urlencode(office)))
    else:
        res.update(msg="文件获取失败或文件格式错误！")
    return jsonify(res)

if __name__ == '__main__':
    #app.run(host='127.0.0.1', port=5000, debug=True)
    app.run(host='0.0.0.0', port=8096, debug=True)
    #app.run()