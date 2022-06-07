from asyncio import constants
from time import strftime
from flask import Flask, request, Response, jsonify
from pymongo import MongoClient
import mongoConnect
import json
import requests
import traceback
from bson import json_util
from flask_cors import CORS
import pandas as pd
from datetime import timedelta
import uuid
import os.path
from datetime import datetime
from bson.objectid import ObjectId
from openpyxl import Workbook
import xlsxwriter
import numpy as np

app = Flask(__name__)
CORS(app)

@app.route('/addData', methods=['POST'])
def addData():
    try:
        f = request.files['data']
        data_xls = pd.read_excel(f, sheet_name='Room Temp', skiprows=[0])
        data = data_xls.to_json(orient='records')
        jsonData = json.loads(data)
        collection = db['tempData']
        data = []
        for each in jsonData:
           each['date'] = each['dd.MM.yy-HHmmss']
           del each['dd.MM.yy-HHmmss']
           data.append(each)

        collection.insert_one({'roomTemperature': data})
        

        return {"data": jsonData}
    except Exception as ex:
        print("Exception : ", ex)
        traceback.print_exc()

@app.route('/getData', methods=['GET','POST'])
def getCompalData():
    try: 
        return { "data": "true"}
        collection = db['tempData']
        response = collection.find_one({}, {'roomTemperature': 1})
        json_docs = [json.dumps(doc, default=json_util.default)
                     for doc in response['roomTemperature']]
        data = []
        for each in json_docs:
            data.append(json.loads(each))

        book = Workbook()
        sheet = book.active
        count = 0
        for i in data:
            if count >= 1: break
            headers = i.keys()
            count = count + 1
            sheet.append(list(headers))
            
        for info in data:
            values = list(info.values())
            print(values)
            sheet.append(values)

        book.save("sample.xlsx")
        return {"data": data}

    except Exception as ex:
        print("Exception : ", ex)
        traceback.print_exc()


@app.route('/testData', methods=['GET','POST'])
def getTestData():
    try: 
        collection = db['compalRawData']
        response = collection.find_one({'_id': ObjectId('626d1a36882f05cd97fdf359')}, {'roomTemperature': 1})
        json_docs = [json.dumps(doc, default=json_util.default)
                     for doc in response['roomTemperature']]
        data = []
        for each in json_docs:
            data.append(json.loads(each))
        return {"data": data}

    except Exception as ex:
        print("Exception : ", ex)
        traceback.print_exc()


@app.route('/monthFilter', methods=['GET', 'POST'])
def monthFilter():
    try:
        collection = db['tempData']
        response = collection.aggregate([{
            "$match": {
            "_id": ObjectId('62826f7bd003892c6000d475')}},
            #  { '$unwind': "$roomTemperature" },
           {"$project": {"roomTemperature": {
            "$filter": {
               'input': "$roomTemperature",
               "as": "room",
               "cond": { "$eq": [{"$month": {'$toDate':  "$$room.date" }}, 9] },
            }}}}
            #  {
            # "$project": {
            # "roomTemperature": { 'month':  {'$month':  {'$toDate': "$roomTemperature.date" }}}, 
            #  }}
            ])
        print("response====", response)
        json_docs = [json.dumps(doc, default=json_util.default)
                     for doc in response]
          
        data = []
        for each in json_docs:
            data.append(json.loads(each))
        return {"data": data}


    except Exception as ex:
        print("Exception : ", ex)
        traceback.print_exc()

@app.route('/filterData', methods=['GET', 'POST'])
def getFilteredData():
    try:
        releaseNo = "44221"
        start = request.json['startDate']
        startDate =  datetime.strptime(start, '%Y/%m/%d %H:%M:%S')
        endDate = request.json['endDate']
        collection = db['compalRawData']
        response = collection.aggregate([{
            "$match": {
            "releaseNo": '44221'}},
        {"$project": {"roomTemperature": {
            "$filter": {
               'input': "$roomTemperature",
               "as": "room",
               "cond": { "$gte": ["$$room.datetime",start]}
            }
        }}}])
        
        json_docs = [json.dumps(doc, default=json_util.default)
                     for doc in response]
          
        data = []
        for each in json_docs:
            data.append(json.loads(each))
        return {"data": data}

    except Exception as ex:
        print("Exception : ", ex)
        traceback.print_exc()

@app.route('/updateData', methods=['GET', 'PUT', 'POST'])
def updateData():
    try:
        id = request.json['id']
        documentId = request.json['documentId']
        ExOvenAmbient= request.json['ExOvenAmbient']
        collection = db['compalRawData']
        collection.update_one({"_id": ObjectId(documentId), 'roomTemperature.id' : id},
         {"$set": {'roomTemperature.$.Ex oven ambient (C)' : ExOvenAmbient}})
        
        return {"message": 'Updated Successfully'}

    except Exception as ex:
        print("Exception : ", ex)
        traceback.print_exc()        
@app.route('/excel', methods=['GET'])
def create_workbook():
        collection = db['compalRawData']
        response = collection.find_one({'_id': ObjectId('626d1a36882f05cd97fdf359')}, {'roomTemperature': 1})
        json_docs = [json.dumps(doc, default=json_util.default)
                     for doc in response['roomTemperature']]
        data = []
        for each in json_docs:
            data.append(json.loads(each))
        workbook = Workbook()
        worksheet = workbook.active
        row =0
        for col, data in enumerate(data):
         worksheet.write_column(row, col, data)
         save_path = 'D:/Nayan_New_Project/Demo_Project/backend/TestFile.xlsx'
        # name_of_file = input("TestFile")
        # completeName = os.path.join(save_path, name_of_file + "xlsx") 
        # print("completeName", completeName)
        workbook.save('ggg.xlsx')
        return {"data": "true"}

if __name__ == '__main__':
    client = MongoClient('localhost', 27017)
    global db
    db = client.python_practice
    app.run(debug=True, port=4000, host='0.0.0.0')
