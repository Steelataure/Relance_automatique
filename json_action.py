import json

file = open('assets/login.json')
data = json.load(file)

email = data["Login"][0]["email"]
password = data["Login"][0]["password"]

mode = data["Login"][0]["mode"]
    


