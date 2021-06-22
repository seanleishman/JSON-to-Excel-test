import json, re

keywords = {'qui'}

with open('sample.json', 'r') as f:

    for line in f.readlines():
        ok_dict = json.loads(line)

        if any(keywords in ok_dict["title"].lower() for keyword in keywords):
            print(ok_dict["id"])
