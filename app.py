from flask import Flask,send_file,request,jsonify
from flask_cors import CORS
from mailmerge import MailMerge
import pandas as pd
import json

# load the dataset and remove items with price is null or not provided. Support category names are lot of duplicates.
# Set operation get unique names and list of names are created.
data = pd.read_excel('Dataset.xlsx')    
data = data[data['Price'].notna()]
Support_Category_Name = list(set(data['Support Category Name'].values))

# load the goals data and create a list from services
goals = pd.read_excel('Goals.xlsx')
goals_list = [service for service in goals['Service'].values]

# Create Flask app and enable CORS
app = Flask(__name__)
cors = CORS(app)

# Return json array of goals
@app.route("/goals")
def goals():
    response = {}
    response['goals'] = goals_list
    return json.dumps(response)

# Retunr json array of support catogery names
@app.route("/supportcategoryname")
def supportCategoryName():
    response = {}
    response['SupportCategoryName'] = Support_Category_Name
    return json.dumps(response)

# Return json array of support item names and ids
@app.route("/supportitemname")
def supportItemName():
    content = request.args
    supportcategoryname = content['supportcategoryname']                     # get support category name from the request parameters
    item_list=data.loc[data['Support Category Name']==supportcategoryname]   # get the list of

    result = {}
    result['SupportItem'] = item_list['Support Item Name'].values
    print(result)
    json_data = json.dumps(result)
    print("*************************")
    
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
    content = request.json
    data_entries = []
    
    for i,j,k,l in zip(content['data'],content['hoursperweek'],content['duration'],content['goals']):
        x={}
        x['SupportCategory'] = i['SupportCategoryName']
        x['ItemName'] = i['SupportItemName']
        x['ItemId'] = i['SupportItemNumber']
        x['Cost'] = str(i['Price']*int(j)*int(k)*4)
        x['H'] = j
        x['M'] = k
        x['Description'] = 'Not yet implemented'
        goals = ""
        for goal in l:
            goals = goals + goal + "\n" + "\n"
        x['Goals'] = goals
        data_entries.append(x)

    document = MailMerge('Schedule of Services (SOS)  draft.docx')
    document.merge_rows('SupportCategory',data_entries)
    document.write('test-output.docx')
    return send_file('test-output.docx', as_attachment=True)

if __name__ == "__main__":
    app.run()