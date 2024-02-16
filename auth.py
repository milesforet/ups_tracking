import requests

def get_client_id():
    file = open("./auth/client_id.txt", "r")
    cid = file.read()
    return cid

def get_client_secret():
    file = open("./auth/client_secret.txt", "r")
    secret = file.read()
    return secret

#GETS A NEW BEARER TOKEN FROM UPS. RETURNS TOKEN AND SAVES IT AS A FILE
def new_token():
    
    url = "https://wwwcie.ups.com/security/v1/oauth/token"

    payload = {
      "grant_type": "client_credentials"
    }

    headers = {
      "Content-Type": "application/x-www-form-urlencoded",
      "x-merchant-id": "string"
    }

    try:
        response = requests.post(url, data=payload, headers=headers, auth=(get_client_id(), get_client_secret()))
        data = response.json()
        bearer_file = open("./auth/bearer.txt", "w")
        bearer_file.write(data['access_token'])
        print("new token")
        return data['access_token']
      
    except Exception as err:
        print(f"error getting new token - {err}")

#GETS TOKEN FROM FILE
def get_bearer_tok():
    file = open("./auth/bearer.txt", "r")
    bear_tok = file.read()
    return bear_tok



#TEST IF BEARER TOKEN IS EXPIRED
def test_bearer_token(bear_tok):
    headers = {
        "transId": "absupstracking01",
        "transactionSrc": "tracking",
        "Authorization": f"Bearer {bear_tok}"
    }
    query = {
        "locale": "en_US",
        "returnSignature": "false",
        "returnMilestones": "false"
        }
    url = "https://wwwcie.ups.com/api/track/v1/details/1Z3124Y80339645436"
    try:
        response = requests.get(url, headers=headers, params=query)
        return response.status_code
    
    except Exception as err:
        print(f"error getting new token - {err}")

if __name__ == '__main__':
    print(get_bearer_tok())