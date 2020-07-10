from flask import Flask,send_file,request,jsonify
from flask_cors import CORS
from mailmerge import MailMerge
from money import Money
import pandas as pd
import json

# Create Flask app and enable CORS
app = Flask(__name__)
cors = CORS(app)

@app.route("/updatedata",methods = ['POST'])
def updateData():
    f = request.files['file'] 
    f.save('Dataset.xlsx')
    return "Success"

@app.route("/updategoals",methods = ['POST'])
def updateGoals():
    f = request.files['file'] 
    f.save('Goals.xlsx') 
    return "Success"

@app.route("/updatepolicy",methods = ['POST'])
def updatePolicy():
    f = request.files['file'] 
    f.save('Policies.xlsx')
    return "Success"

# Return json array of goals
@app.route("/goals")
def goals():
    response = {}
    # load the goals data and create a list from services
    goals = pd.read_excel('Goals.xlsx')
    goals = goals.fillna("")
    goals_list = []
    for i,j in zip(goals['Service'].values,goals['Goals'].values):
        goal = {}
        goal[i] = j
        goals_list.append(goal)
    response['goals'] = goals_list
    return json.dumps(response)

@app.route("/goaldescription")
def goaldescription():
    response = {}
    response['description'] = goals_descriptions
    return json.dumps(response)

# Return json array of policies
@app.route("/policy")
def policy():
    response = {}
    # load the policy file and creata a list
    policies = pd.read_excel('Policies.xlsx')
    policy_list = [policy for policy in policies['Policy'].values]
    response['policy'] = policy_list
    return json.dumps(response)

# Retunr json array of support catogery names
@app.route("/supportcategoryname")
def supportCategoryName():
    # load the dataset and remove items with price is null or not provided. Support category names are lot of duplicates.
    # Set operation get unique names and list of names are created.
    data = pd.read_excel('Dataset.xlsx')    
    data = data[data['Price'].notna()]
    Support_Category_Name = list(set(data['Support Category Name'].values))
    Support_Category_Name.sort()  # This is not need.
    response = {}
    response['SupportCategoryName'] = Support_Category_Name
    return json.dumps(response)

# Return json array of support item names and ids
@app.route("/supportitemname")
def supportItemName():
    content = request.args
    supportcategoryname = content['supportcategoryname']                     # get support category name from the request parameters
    data = pd.read_excel('Dataset.xlsx')    
    data = data[data['Price'].notna()]
    item_list=data.loc[data['Support Category Name']==supportcategoryname]   # get the array of items with requested support category name
    result = {}
    
    result['SupportItem'] = [item for item in item_list['Support Item Name'].values]   # create a list from array of items in order to retun easily
    json_data = json.dumps(result)    
    return json_data

# Return json object of the details of requested item
@app.route("/supportitemdetails")
def supportitemdetails():
    content = request.args
    supportcategoryname = content['supportcategoryname'] 
    supportitem = content['supportitem']
    data = pd.read_excel('Dataset.xlsx')    
    data = data[data['Price'].notna()]
    item_details = data.query('`Support Category Name`=={} & `Support Item Name`=={}'.format('"'+supportcategoryname+'"','"'+supportitem+'"'))
    return jsonify({"SupportCategoryName": item_details['Support Category Name'].values[0], "SupportItemNumber": item_details['Support Item Number'].values[0], "SupportItemName": item_details['Support Item Name'].values[0],"Price": item_details['Price'].values[0]})

# Return the word document filled with data
@app.route('/document', methods=['POST'])
def document():
    content = request.json
    data_entries = []
    total_cost = 0
    for i,j,l,m,n in zip(content['data'],content['hours'],content['goals'],content['description'],content['hoursFrequncy']):
        x={}
        x['SupportCategory'] = i['SupportCategoryName']
        x['ItemName'] = i['SupportItemName']
        x['ItemId'] = i['SupportItemNumber']
        
        multiplication = ""
        if (n[-1]=="W"):
            x['H'] = "Hours per Week: "+ n.split(',')[0] + "\n" + "Duration: " + n.split(',')[1] + " weeks"
            multiplication = n.split(',')[0] + "x" + n.split(',')[1] + "x"
        elif (n[-1]=="M"):
            x['H'] = "Hours per Month: "+ n.split(',')[0] + "\n" + "Duration: " + n.split(',')[1] + " months"
            multiplication = n.split(',')[0] + "x" + n.split(',')[1] + "x"
        else:
            x['H'] = "Hours per plan period: "+ n
            multiplication = n + "x"
        
        cost = Money(str(i['Price']*int(j)), 'USD')
        x['Cost'] = multiplication + "$" + Money(str(i['Price']),'USD').format('en_US') + "\n= " + cost.format('en_US') 
        total_cost += i['Price']*int(j)
        x['Description'] = str(m)
        goals = ""
        for goal in l:
            goals = goals + goal + "\n" + "\n"
        x['Goals'] = goals
        data_entries.append(x)

    document = MailMerge('WordTemplate.docx')
    total_cost = Money(str(total_cost), 'USD')
    document.merge(totalcost= total_cost.format('en_US'))
    document.merge(name=str(content['name']),ndis=str(content['ndis']),sos=str(content['sos']),duration=str(int(content['duration']/7))+" Weeks",start=content['start'],end=content['end'],today=content['today'],policy=content['policy'])
    document.merge_rows('SupportCategory',data_entries)
    document.write('test-output.docx')
    return send_file('test-output.docx', as_attachment=True)

if __name__ == "__main__":
    app.run()