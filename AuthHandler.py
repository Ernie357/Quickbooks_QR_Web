from dotenv import load_dotenv
import os
import requests
import base64
import urllib.parse
import requests

'''
    Purpose: Handles the retrieving of realm ID and access_token for API calls

    Note: "Callback URL" or "REDIRECT_URI" env is where the user is sent after being 
        authorized by Intuit. Currently is localhost:8080/callback. 
        When redirected here, the URL has search params with a code and a realmId.
        The code can be exchanged for access token via the intuit TOKEN_URL.
'''
class AuthHandler:
    def __init__(self, is_prod: bool):
        load_dotenv()
        self.client_id = os.getenv("PROD_CLIENT_ID" if is_prod else "DEV_CLIENT_ID")
        self.client_secret = os.getenv("PROD_CLIENT_SECRET" if is_prod  else "DEV_CLIENT_SECRET")
        self.redirect_uri = os.getenv("PROD_REDIRECT_URI" if is_prod else "DEV_REDIRECT_URI")
        self.scopes = os.getenv("SCOPES")
        if not self.scopes or not self.redirect_uri or not self.client_id or not self.client_secret:
            raise Exception("Missing 1 or more required .env variables.")
        discovery_document_url = "https://developer.api.intuit.com/.well-known/openid_configuration"
        discover_document_headers = { "Accept": "application/json" }
        response = requests.get(url=discovery_document_url, headers=discover_document_headers)
        data = response.json() if response.status_code == 200 else {}
        self.auth_base = data.get("authorization_endpoint", None)
        self.token_url = data.get("token_endpoint", None)
        if self.auth_base is None or self.token_url is None:
            raise Exception("Could not either/both auth or token URL from Intuit document discover.")

    '''
        Gets the url in which Intuit will ask for a sign in, built from env variables
    '''
    def get_auth_url(self):
        params = {
            "client_id": self.client_id,
            "response_type": "code",
            "scope": self.scopes,
            "redirect_uri": self.redirect_uri,
            "state": "random_string_for_csrf"
        }
        return f"{self.auth_base}?{urllib.parse.urlencode(params)}"
    
    '''
        Using the code and realmId params of the callback URL, which are
        passed in as arguments here, generates and returns json that includes
        refresh_token, access_token, realm_id, and expires_in
    '''
    def get_auth_tokens_from_code(self, code: str | None, realm_id: str | None) -> tuple[str, str]:
        if not code or not realm_id:
            raise Exception("One or more URL params missing.")
        credentials = base64.b64encode(f"{self.client_id}:{self.client_secret}".encode()).decode()
        response = requests.post(
            self.token_url if self.token_url is not None else "",
            headers={
                "Authorization": f"Basic {credentials}",
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
            },
            data={
                "grant_type": "authorization_code",
                "code": code,
                "redirect_uri": self.redirect_uri,
            },
        )
        if response.status_code != 200:
            raise Exception(f"Error {response.status_code}: {response.text}")
        tokens = response.json()
        if not tokens["access_token"]:
            raise Exception("access_token not returned by token endpoint.")
        print("\nIntuit tid:", response.headers["intuit_tid"], "\n")
        return (tokens["access_token"], realm_id)