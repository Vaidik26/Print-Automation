"""
Adobe Acrobat Sign Handler Module
Handles interactions with Adobe Sign API to generate persistent signing links.
"""
import requests
import time
import json
import os

class AdobeSignHandler:
    def __init__(self, integration_key, client_secret, technical_account_email, shard="secure"):
        """
        Initialize Adobe Sign Handler.
        
        :param integration_key: Client ID (Application ID)
        :param client_secret: Client Secret
        :param technical_account_email: Email of the technical account (admin)
        :param shard: Adobe Sign shard (e.g., 'secure', 'na1', 'eu1'). defaults to 'secure' for login.
        """
        self.client_id = integration_key
        self.client_secret = client_secret
        self.tech_account_email = technical_account_email
        self.base_url = f"https://{shard}.echosign.com"
        self.access_token = None
        self.expiry_time = 0
        
        # Note: For Server-to-Server (Internal App), we typically use an Integration Key 
        # that allows JWT or standard OAuth. 
        # Detailed auth flow usually requires a refresh token for recurring usage without user interaction.
        # For simplicity in this script, we'll assume the user provides a REFRESH TOKEN 
        # or we implement the flow to get one. 
        # However, Adobe Sign 'Integration Key' often acts as a Bearer token generator if set up correctly.
        
        # ACTUALLY: The easiest way for a script is using a permanent "Integration Key" as a token 
        # OR using a Refresh Token. 
        # Let's assume the user provides a Refresh Token or we guide them to get one.
        
        # For this implementation, we will ask the user for:
        # 1. Integration Key (Client ID)
        # 2. Client Secret
        # 3. Refresh Token (One-time setup required to get this)
        
        self.refresh_token = None # Will be set via method or init
        
    def set_refresh_token(self, token):
        self.refresh_token = token

    def _get_access_token(self):
        """Exchange Refresh Token for Access Token"""
        if self.access_token and time.time() < self.expiry_time:
            return self.access_token
            
        if not self.refresh_token:
            raise Exception("Access Token expired and no Refresh Token provided.")
            
        url = f"{self.base_url}/oauth/v2/refresh"
        payload = {
            "refresh_token": self.refresh_token,
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "refresh_token"
        }
        
        response = requests.post(url, data=payload)
        
        if response.status_code == 200:
            data = response.json()
            self.access_token = data['access_token']
            self.expiry_time = time.time() + int(data['expires_in']) - 60 # Buffer
            return self.access_token
        else:
            raise Exception(f"Failed to refresh token: {response.text}")

    def upload_transient_document(self, file_name, file_bytes):
        """Upload file temporarily to get a transient document ID"""
        token = self._get_access_token()
        url = f"{self.base_url}/api/rest/v6/transientDocuments"
        
        headers = {
            "Authorization": f"Bearer {token}"
        }
        
        # Adobe expects multipart/form-data with 'File-Name' and 'File'
        files = {
            'File': (file_name, file_bytes, 'application/pdf') # Suggest converting to PDF if possible, or keep original mime
        }
        data = {
            'File-Name': file_name
        }
        
        response = requests.post(url, headers=headers, files=files, data=data)
        
        if response.status_code == 201:
            return response.json()['transientDocumentId']
        else:
            raise Exception(f"Failed to upload document: {response.text}")

    def create_agreement_signing_url(self, transient_doc_id, recipient_email, agreement_name):
        """
        Create an agreement and get the signing URL.
        Note: To get a signing URL (not email), the agreement state must be IN_PROCESS.
        """
        token = self._get_access_token()
        url = f"{self.base_url}/api/rest/v6/agreements"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "fileInfos": [{
                "transientDocumentId": transient_doc_id
            }],
            "name": agreement_name,
            "participantSetsInfo": [{
                "memberInfos": [{
                    "email": recipient_email
                }],
                "order": 1,
                "role": "SIGNER"
            }],
            "signatureType": "ESIGN",
            "state": "IN_PROCESS",
            # This is CRITICAL for getting the link manually
            "emailOption": {
                "sendOptions": {
                    "initiationEmails": "NONE", # Don't send email from Adobe
                    "completionEmails": "NONE",
                    "inFlightEmails": "NONE"
                }
            }
        }
        
        # 1. Create Agreement
        response = requests.post(url, headers=headers, json=payload)
        
        if response.status_code == 201:
            agreement_id = response.json()['id']
            
            # 2. Fetch Signing URL
            # Note: It might take a moment to be available
            time.sleep(1) 
            
            sign_url_endpoint = f"{self.base_url}/api/rest/v6/agreements/{agreement_id}/signingUrls"
            link_response = requests.get(sign_url_endpoint, headers=headers)
            
            if link_response.status_code == 200:
                 # Check response structure
                 link_data = link_response.json()
                 if 'signingUrlSetInfos' in link_data and link_data['signingUrlSetInfos']:
                      return link_data['signingUrlSetInfos'][0]['signingUrls'][0]['email'], agreement_id
                 else:
                      raise Exception("No signing URLs returned.")
            else:
                 raise Exception(f"Failed to get signing URL: {link_response.text}")
            
        else:
            raise Exception(f"Failed to create agreement: {response.text}")
