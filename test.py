import requests

BASE = "http://127.0.0.1:8000/api/v1"
ADMIN_TOKEN = ""

HEADERS = {
    "Authorization": ADMIN_TOKEN
}

# ----------------------------
# CREATE USER
# ----------------------------
r = requests.post(
    f"{BASE}/admin/users/create/",
    headers=HEADERS,
    json={"email": "user11@example.com"}
)

if r.status_code == 201:
    print("User created:", r.json())
elif r.status_code == 409:
    print("User already exists")
else:
    print("Error:", r.status_code, r.text)

# ----------------------------
# LIST USERS
# ----------------------------

headers = {"Authorization": ADMIN_TOKEN}

resp = requests.get(
    "http://127.0.0.1:8000/api/v1/admin/users/view/",
    headers=headers,
    params={"email": "user1@example.com"}
)

print(resp.status_code)
print(resp.text)

# ----------------------------
# ROTATE USER TOKEN
# ----------------------------
'''
r = requests.post(
    f"{BASE}/admin/users/change/",
    headers=HEADERS,
    json={"email": "user1@example.com"}
)

print("Token rotated:", r.json())
'''


import requests

API_URL = "http://127.0.0.1:8000/api/v1/upload/"
USER_TOKEN = "3e4065632306488cbe9f5cc63ca49e07"  # Токен пользователя, аутентификация

headers = {
    "Authorization": USER_TOKEN
}

files = {
    'file': ('example.docx', open('example.docx', 'rb'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
}

data = {
    "style": "formal",
    "format": "pdf",
    "dictionary": "custom_dict"
}

response = requests.post(API_URL, headers=headers, files=files, data=data)

print(response.status_code)
print(response.json())


import requests

# Your details
url = "http://127.0.0.1:8000/api/v1/status/30b1e1b0-6431-48e2-8be8-b8db17819160/"

# Create headers
headers = {
    "Authorization": USER_TOKEN
}

# Make the request
response = requests.get(url, headers=headers)

# Show what happened
print(f"Status: {response.status_code}")
print(f"Response: {response.text}")



import requests

# Replace these with your actual values
url = "http://127.0.0.1:8000/api/v1/get_work/"

headers = {"Authorization": ADMIN_TOKEN}

response = requests.get(url, headers=headers)

print(f"URL: {url}")
print(f"Status Code: {response.status_code}")
print(f"Response: {response.text}")




import requests

url = "http://127.0.0.1:8000/api/v1/save_work/"
token = ADMIN_TOKEN

headers = {
    "Authorization": token,
}

data = {
    "file_id": "e11b9c9d-653d-45d4-ac2b-0f83f35a588e",
    "processed": True
}

response = requests.post(url, json=data, headers=headers)

print(f"Status: {response.status_code}")
print(f"Response: {response.text}")