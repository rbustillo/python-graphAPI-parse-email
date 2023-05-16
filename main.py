import requests

# Replace with your actual credentials
client_id = "ENTER CLIENT ID HERE"
# Secret value not ID
client_secret = "ENTER CLIENT SECRET VALUE HERE"
tenant_id = "ENTER TENANT ID HERE"
email_address = "ENTER EMAIL ADDRESS HERE"
access_token = ""

# Authenticate and get access token
def get_access_token():
    global access_token

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }

    response = requests.post(url, data=data)
    response_data = response.json()
    access_token = response_data["access_token"]

# Get emails from the inbox
def get_emails():
    url = f"https://graph.microsoft.com/v1.0/users/{email_address}/messages"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    response_data = response.json()
    print(response_data)

    if "value" in response_data:
        return response_data["value"]
    else:
        return []

# Parse the body of the email
def parse_email_body(email):
    email_id = email["id"]
    url = f"https://graph.microsoft.com/v1.0/users/{email_address}/messages/{email_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    response_data = response.json()
    print(response_data)

    if "body" in response_data:
        body_type = response_data["body"]["contentType"]
        body_content = response_data["body"]["content"]

        if body_type == "text":
            print("Text body:")
            print(body_content)
        elif body_type == "html":
            print("HTML body:")
            print(body_content)
        elif body_type == "multipart":
            for part in response_data["body"]["content"]:
                if part["contentType"] == "text/plain":
                    print("Text body:")
                    print(part["content"])
                    break

# Main program
get_access_token()
emails = get_emails()

if emails:
    for email in emails:
        parse_email_body(email)
else:
    print("No emails found.")
