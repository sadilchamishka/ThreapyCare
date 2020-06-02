from flask import Flask,send_file,request,jsonify
from flask_cors import CORS
from mailmerge import MailMerge
import pandas as pd
import json

data_entries = [{
    'SupportCategory': 'Red Shoes',
    'SupportItem': '$10.00',
    'H': '2500',
    'D': '10',
    'S':'gsfagh',
    'Cost':'$200',
    'Description':'gafstdt',
    'Goals':'rtwtfhg'
}, {
    'SupportCategory': 'blue Shoes',
    'SupportItem': '$50.00',
    'H': '2500',
    'D': '10',
    'S':'mannnn',
    'Cost':'$400',
    'Description':'gafstdt',
    'Goals':'rtwtfhg'
}, {
    'SupportCategory': 'green Shoes',
    'SupportItem': '80.00',
    'H': '2500',
    'D': '10',
    'S':'gsfagh',
    'Cost':'$200',
    'Description':'sadil',
    'Goals':'ailaa'
}]

data = pd.read_excel('PB Support Catalogue 2019-20 Feb .xlsx')
result = {}
values = []

for i in data['Support Category Name'].values:
    if i not in values:
        values.append(i)
result['SupportCategoryName'] = values

Support_Category_Names = json.dumps(result)

app = Flask(__name__)
cors = CORS(app)

@app.route("/")
def home():
    document = MailMerge('Schedule of Services (SOS)  draft.docx')
    document.merge_rows('SupportCategory',data_entries)
    document.write('test-output.docx')
    return send_file('test-output.docx', as_attachment=True)

@app.route("/supportcategoryname")
def supportCategoryName():
    return Support_Category_Names

@app.route("/supportitemname")
def supportItemName():
    content = request.args
    supportcategoryname = content['supportcategoryname']
    
    item_list=data.loc[data['Support Category Name']==supportcategoryname]
    result = {}
    values = []
    for i in item_list[['Support Item Number','Support Item Name']].values:
        item = {}
        item["ItemNumber"] = i[0]
        item["ItemName"] = i[1]
        values.append(item)

    result['SupportItem'] = values
    json_data = json.dumps(result)
    return json_data

@app.route("/supportitemdetails")
def supportitemdetails():
    content = request.args
    supportitem = content['supportitem']
    item_details = data.loc[data['Support Item Name']==supportitem].values[0]
    price = item_details[6]
    if pd.isna(item_details[6]):
        price = 0
    return jsonify({"SupportCategoryName": item_details[0], "SupportItemNumber": item_details[1], "SupportItemName": item_details[2],"Price":price })

@app.route('/document', methods=['POST'])
def document():
    r=""
    content = request.json
    for i in content['data']:
        r=r+i
    return r

if __name__ == "__main__":
    app.run()