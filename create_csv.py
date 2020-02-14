import boto3
import json
import csv
from datetime import date
import os

def create_csv():
    #
    ACCESS_KEY = 'yours_access_key'
    SECRET_KEY = 'yours_secret_key'

    dynamodb = boto3.resource('dynamodb',
                              aws_access_key_id=ACCESS_KEY,
                              aws_secret_access_key=SECRET_KEY,
                              region_name='ap-northeast-2')

    table = dynamodb.Table('TEST')
    response = table.scan()
    items = response['Items']
    datenow = str(date.today())[2:4] + str(date.today())[5:7] + str(date.today())[8:10]
    output_file = 'createCSV/'+ datenow +'test.csv'
    #

    while 'LastEvaluatedKey' in response:
        response = table.scan(ExclusiveStartKey=response['LastEvaluatedKey'])
        items.extend(response['Items'])

    check_file = os.path.exists(output_file)
    if check_file == True:
        os.unlink(output_file)
        csvfile = open(output_file, "w", newline="")
        csvwirter = csv.writer(csvfile)
        for i in range(len(items)):
            csvwirter.writerow(json.dumps(items[i]).replace('"',"'").split(','))
        print(output_file, '생성되었습니다.')
        csvfile.close()
    else:
        csvfile = open(output_file, "w", newline="")
        csvwirter = csv.writer(csvfile)
        for i in range(len(items)):
            csvwirter.writerow(json.dumps(items[i]).replace('"',"'").split(','))
        print(output_file, '생성되었습니다.')
        csvfile.close()
