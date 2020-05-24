from flask import Flask,send_file,jsonify
from mailmerge import MailMerge
import pandas as pd

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
Support_Category_Names = str(set(data['Support Category Name']))

app = Flask(__name__)

@app.route("/")
def home():
    document = MailMerge('Schedule of Services (SOS)  draft.docx')
    document.merge_rows('SupportCategory',data_entries)
    document.write('test-output.docx')
    return send_file('test-output.docx', as_attachment=True)

@app.route("/supportcategoryname")
def SupportCategoryName():
    return jsonify(Support_Category_Names)
    
if __name__ == "__main__":
    app.run()