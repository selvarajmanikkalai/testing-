from requests import post
import json

CLIENT_SECRET_VALUE = 'bwo8Q~suPqebsMNDLiPbkG06i.46EjgckevtCa0H'
TENANT_ID = 'ca1b8d73-100a-46b1-b906-4cd263a54137'
CLIENT_ID = '797f4846-ba00-4fd7-ba43-dac1f8f63013'


LOGIN_URI = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'

headers = {
    'Host': 'login.microsoftonline.com',
    'Content-Type': 'application/x-www-form-urlencoded'
}


body = {
    'client_id': CLIENT_ID,
    'scope': 'https://graph.microsoft.com/.default',
    'client_secret': CLIENT_SECRET_VALUE,
    'grant_type': 'client_credentials',
    'tenant': TENANT_ID
}


response = post(url=LOGIN_URI, headers=headers, data=body)
response.raise_for_status()
response_body = response.json()

authorization_token = f"{response_body['token_type']} {response_body['access_token']}"

print(authorization_token)

email_header = {
    'Authorization': authorization_token,
    'Content-Type': 'application/json'
}

message = {
    'body': {
        'content': 'Outlook Mail Testing Demo',
        'contentType': 'Text'
    },
    'sender': {
        'emailAddress': {
            'address': '<selvaraj133ece@outlook.com>',
            'name': 'Name of Shared Mailbox'
        }
    },
    'subject': 'Testing email',
    'toRecipients': [
        {
            'emailAddress': {
                'address': '<smanikkalai@presidio.com>',
                'name': 'test-user'
            }
        }
    ]
}

email_body = {
    'message': message
}

email_send_response = post(url='https://graph.microsoft.com/v1.0/users/1ab4e76f-5f52-44b8-8a72-7d03c05e6ff4/sendMail', headers=email_header, data=json.dumps(email_body))
print(email_send_response)