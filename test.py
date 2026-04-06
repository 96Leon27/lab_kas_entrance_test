import requests

url = 'http://localhost:5000/public/report/export'
files = {'file': open('test.txt', 'rb')}
response = requests.post(url, files=files)

with open('report.xlsx', 'wb') as f:
    f.write(response.content)
