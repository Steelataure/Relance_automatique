import json
from re import sub

file = open('assets/login.json')
data = json.load(file)

email = data["Login"][0]["email"]
password = data["Login"][0]["password"]
