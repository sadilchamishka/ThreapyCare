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
Support_Category_Name.sort()  # This is not need.

# load the goals data and create a list from services
goals = pd.read_excel('Goals.xlsx')
goals_list = [service for service in goals['Service'].values]

# load the policy file and creata a list
policies = pd.read_excel('Policies.xlsx')
policy_list = [policy for policy in policies['Policy'].values]

# Create Flask app and enable CORS
app = Flask(__name__)
cors = CORS(app)

@app.route("/updatedata",methods = ['POST'])
def updateData():
    f = request.files['file'] 
    f.save('Dataset.xlsx')
    global data 
    data = pd.read_excel('Dataset.xlsx')
    data = data[data['Price'].notna()]
    global Support_Category_Name 
    Support_Category_Name = list(set(data['Support Category Name'].values))
    Support_Category_Name.sort()
    return "Success"

@app.route("/updategoals",methods = ['POST'])
def updateGoals():
    f = request.files['file'] 
    f.save('Goals.xlsx') 
    goals = pd.read_excel('Goals.xlsx')
    global goals_list 
    goals_list = [service for service in goals['Service'].values]
    return "Success"

# Return json array of goals
@app.route("/goals")
def goals():
    response = {}
    response['goals'] = goals_list
    return json.dumps(response)

# Return json array of policies
@app.route("/policy")
def policy():
    response = {}
    response['policy'] = policy_list
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
    print("***************************")
    supportcategoryname = content['supportcategoryname']                     # get support category name from the request parameters
    print(supportcategoryname)
    item_list=data.loc[data['Support Category Name']==supportcategoryname]   # get the array of items with requested support category name
    print(item_list)
    result = {}
    
    result['SupportItem'] = [item for item in item_list['Support Item Name'].values]   # create a list from array of items in order to retun easily
    json_data = json.dumps(result)    
    return json_data

# Return json object of the details of requested item
@app.route("/supportitemdetails")
def supportitemdetails():
    content = request.args
    supportitem = content['supportitem']
    item_details = data.loc[data['Support Item Name']==supportitem].values[0]   # get the first and only item from the array
    return jsonify({"SupportCategoryName": item_details[0], "SupportItemNumber": item_details[1], "SupportItemName": item_details[2],"Price": item_details[6]})

# Return the word document filled with data
@app.route('/document', methods=['POST'])
def document():
    content = request.json
    data_entries = []
    
    for i,j,l,m,n in zip(content['data'],content['hours'],content['goals'],content['description'],content['hoursFrequncy']):
        x={}
        x['SupportCategory'] = i['SupportCategoryName']
        x['ItemName'] = i['SupportItemName']
        x['ItemId'] = i['SupportItemNumber']
        x['Cost'] = str(i['Price']*int(j))
        if (n[-1]=="W"):
            x['H'] = "Hours per Week "+ n.split(',')[0] + "\n" + "Duration " + n.split(',')[1]
        elif (n[-1]=="M"):
            x['H'] = "Hours per Month "+ n.split(',')[0] + "\n" + "Duration " + n.split(',')[1]
        else:
            x['H'] = "Hours "+ n
        x['Description'] = str(m)
        goals = ""
        for goal in l:
            goals = goals + goal + "\n" + "\n"
        x['Goals'] = goals
        data_entries.append(x)

    document = MailMerge('WordTemplate.docx')
    document.merge(name=str(content['name']),ndis=str(content['ndis']),sos=str(content['sos']),duration=str(int(content['duration']/7))+" Weeks",start=content['start'],end=content['end'],today=content['today'],policy=content['policy'])
    document.merge_rows('SupportCategory',data_entries)
    document.write('test-output.docx')
    return send_file('test-output.docx', as_attachment=True)

if __name__ == "__main__":
    app.run()