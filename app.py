from flask import Flask,send_file,request,jsonify
from flask_cors import CORS
from mailmerge import MailMerge
from money import Money
import pandas as pd
from datetime import datetime
from datetime import timedelta
import json
import mysql.connector
import os
import jwt
import shutil
import hashlib

dbhost = os.environ.get('dbhost', None)
user = os.environ.get('user', None)
password = os.environ.get('password', None)
database = os.environ.get('database', None)
secret = os.environ.get('secret', None)

mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
mycursor = mydb.cursor()
sql = "SELECT * from files"
mycursor.execute(sql)
record = mycursor.fetchall()

with open("Dataset.xlsx", 'wb') as file:
    file.write(record[0][1])

with open("Goals.xlsx", 'wb') as file:
    file.write(record[1][1])

with open("Policies.xlsx", 'wb') as file:
    file.write(record[2][1])


## Create Flask app and enable CORS
app = Flask(__name__)
cors = CORS(app)

def encode_auth_token(role):
    """
    Generates the Auth Token
    :return: string
    """
    try:
        payload = {
            'exp': datetime.utcnow() + timedelta(days=1, seconds=0),
            'iat': datetime.utcnow(),
            'role': role
        }
        return jwt.encode(
            payload,
            secret,
            algorithm='HS256'
        )
    except Exception as e:
        return e

def decode_auth_token(auth_token):
    """
    Decodes the auth token
    :param auth_token:
    :return: integer|string
    """
    try:
        payload = jwt.decode(auth_token, secret)
        return payload['role']
    except jwt.ExpiredSignatureError:
        return 'Signature expired'
    except jwt.InvalidTokenError:
        return 'Invalid token'


@app.route("/login",methods = ['POST'])
def login():
    content = request.json
    mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
    mycursor = mydb.cursor()

    hash_object = hashlib.md5(content['password'].encode())
    hash = hash_object.hexdigest()
    sql = "SELECT * FROM users WHERE name = '"+content['name']+"' and password = '"+hash+"'"
    mycursor.execute(sql)
    myresult = mycursor.fetchall()
    if len(myresult)==1:
        token = encode_auth_token(myresult[0][3])
        return token
    else:
        return "Invalid"

@app.route("/registeruser",methods = ['POST'])
def register():
    content = request.json
    mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
    mycursor = mydb.cursor()

    hash_object = hashlib.md5(content['password'].encode())
    hash = hash_object.hexdigest()

    sql = "INSERT INTO users (email, name, password, role) VALUES (%s, %s, %s, %s)"
    val = (content['email'], content['name'], hash, content['role'])
    try:
        mycursor.execute(sql, val)
        mydb.commit()
        return "Success"
    except:
        return "User already Exist"

@app.route("/users")
def viewUsers():
    mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
    mycursor = mydb.cursor()
    sql = "SELECT email,name,role FROM users ORDER BY role ASC"
    mycursor.execute(sql)
    myresult = mycursor.fetchall()
    
    result = []

    for x in myresult:
        user1 = list(x)
        user_details = {}

        if user1[2]=='admin':
            user_details['name'] = "(admin) "+user1[1]
        else:
            user_details['name'] = user1[1]
            
        user_details['email'] = user1[0]
        user_details['password'] = ""
        result.append(user_details)

    return jsonify({'users':result})

@app.route("/updateuser",methods=['POST'])
def updateUser():
    content = request.json
    result = decode_auth_token(content['token'])
    if (result=='Signature expired' or result=='Invalid token'):
        return "Invalid token"
    elif (result=="admin"):
        mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
        mycursor = mydb.cursor()

        name = content['name']
        if name[1:6]=="admin":
            name = name[8:]
        
        if content['password']=="":
            sql = "UPDATE users SET email = %s WHERE name = %s"
            val = (content['email'],name)
            mycursor.execute(sql,val)
            mydb.commit()
            return "Success"
        else:
            hash_object = hashlib.md5(content['password'].encode())
            hash = hash_object.hexdigest()
            sql = "UPDATE users SET email = %s,password = %s WHERE name = %s"
            val = (content['email'],hash,name)
            mycursor.execute(sql,val)
            mydb.commit()
            return "Success"
    else:
        return "Unauthorized"

@app.route("/deleteuser",methods=['POST'])
def deleteUser():
    content = request.json
    result = decode_auth_token(content['token'])
    if (result=='Signature expired' or result=='Invalid token'):
        return "Invalid token"
    elif (result=="admin"):
        mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
        mycursor = mydb.cursor()
        sql = "DELETE FROM users WHERE name = "+"'"+content['name']+"'"
        mycursor.execute(sql)
        mydb.commit()
        return "Success"
    else:
        return "Unauthorized"


@app.route("/auth",methods = ['POST'])
def auth():
    content = request.json
    print(content['token'])
    return decode_auth_token(content['token'])

@app.route("/updatedata",methods = ['POST'])
def updateData():
    f = request.files['file'] 
    f.save('Dataset_temp.xlsx')
    data = pd.read_excel('Dataset_temp.xlsx')
    columns =  data.columns

    for i in ['Support Category Name','Support Item Number','Support Item Name','Price']:
        if i not in columns:
            return "Invalid"

    binaryData = ''
    with open("Dataset_temp.xlsx", 'rb') as file:
        binaryData = file.read()

    mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
    mycursor = mydb.cursor()

    sql = "UPDATE files SET file = %s WHERE name = %s"
    val = (binaryData,"Dataset")
    mycursor.execute(sql, val)
    mydb.commit()

    shutil.move('Dataset_temp.xlsx', 'Dataset.xlsx')
    return "Success"

@app.route("/updategoals",methods = ['POST'])
def updateGoals():
    f = request.files['file'] 
    f.save('Goals_temp.xlsx') 
    data = pd.read_excel('Goals_temp.xlsx')
    columns =  data.columns

    for i in ['Service','Goals']:
        if i not in columns:
            return "Invalid"

    binaryData = ''
    with open("Goals_temp.xlsx", 'rb') as file:
        binaryData = file.read()

    mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
    mycursor = mydb.cursor()

    sql = "UPDATE files SET file = %s WHERE name = %s"
    val = (binaryData,"Goals")
    mycursor.execute(sql, val)
    mydb.commit()

    shutil.move('Goals_temp.xlsx', 'Goals.xlsx')
    return "Success"

@app.route("/updatepolicy",methods = ['POST'])
def updatePolicy():
    f = request.files['file'] 
    f.save('Policies_temp.xlsx')
    data = pd.read_excel('Policies_temp.xlsx')
    columns =  data.columns

    for i in ['Policy']:
        if i not in columns:
            return "Invalid"

    binaryData = ''
    with open("Policies_temp.xlsx", 'rb') as file:
        binaryData = file.read()

    mydb = mysql.connector.connect(host=dbhost,user=user,password=password,database=database)
    mycursor = mydb.cursor()

    sql = "UPDATE files SET file = %s WHERE name = %s"
    val = (binaryData,"Policies")
    mycursor.execute(sql, val)
    mydb.commit()

    shutil.move('Policies_temp.xlsx', 'Policies.xlsx')
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
    support_category_map = {}
    for i,j,l,n in zip(content['data'],content['hours'],content['goals'],content['hoursFrequncy']):
        x={}
        SupportCategoryName = i['SupportCategoryName']
        x['SupportCategory'] = SupportCategoryName
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
            x['H'] = "Hours per plan period: "+ n + " hours"
            multiplication = n + "x"
        
        cost = Money(str(i['Price']*int(j)), 'USD')
        x['Cost'] = multiplication  + Money(str(i['Price']),'USD').format('en_US') + "\n= " + cost.format('en_US')

        if SupportCategoryName in support_category_map:
            support_category_map[SupportCategoryName] += i['Price']*int(j)
        else:
            support_category_map[SupportCategoryName] = i['Price']*int(j)
        
        goals = ""
        for goal in l:
            goals = goals + goal + "\n" + "\n"
        x['Goals'] = goals
        data_entries.append(x)
	
    totalcost = ""
    for key,value in support_category_map.items():
        totalcost = totalcost + key + " = " + Money(str(value),'USD').format('en_US') + "\n"

    document = MailMerge('WordTemplate.docx')
    total_cost = totalcost
    document.merge(totalcost= total_cost.format('en_US'))

    datetimeobject = datetime.strptime(content['start'],'%Y-%m-%d')
    startDate = datetimeobject.strftime('%d/%m/%Y')

    datetimeobject = datetime.strptime(content['end'],'%Y-%m-%d')
    endDate = datetimeobject.strftime('%d/%m/%Y')


    datetimeobject = datetime.strptime(content['today'],'%Y-%m-%d')
    today = datetimeobject.strftime('%d/%m/%Y')

    document.merge(name=str(content['name']),ndis=str(content['ndis']),sos=str(content['sos']),duration=str(int(content['duration']/7))+" weeks",start=startDate,end=endDate,today=today,policy=content['policy'])
    document.merge_rows('SupportCategory',data_entries)
    document.write('test-output.docx')
    return send_file('test-output.docx', as_attachment=True)

if __name__ == "__main__":
    app.run()