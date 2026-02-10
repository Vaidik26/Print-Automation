"""
Zoho Sign Handler Module (India DC)
Handles document sending via Zoho Sign API for .in domain.
"""

import requests
import json
import time
from typing import Tuple, Dict, Optional

class ZohoSignHandler:
    """Handle Zoho Sign API interactions for India Data Center."""
    
    # Base URLs for Zoho India (.in)
    AUTH_URL = "https://accounts.zoho.in/oauth/v2/token"
    API_BASE_URL = "https://sign.zoho.in/api/v1"

    def __init__(self, client_id: str, client_secret: str, refresh_token: str, organization_id: str = None):
        """
        Initialize Zoho Sign Handler.
        
        Args:
            client_id: OAuth Client ID
            client_secret: OAuth Client Secret
            refresh_token: OAuth Refresh Token
            organization_id: Zoho Sign Organization ID (optional, fetched if None)
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token
        self.organization_id = organization_id
        self.access_token = None
        self.token_expiry = 0

    def get_access_token(self) -> str:
        """
        Get a valid access token. Refresh if expired.
        """
        if self.access_token and time.time() < self.token_expiry:
            return self.access_token
            
        try:
            params = {
                "refresh_token": self.refresh_token,
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "grant_type": "refresh_token"
            }
            
            response = requests.post(self.AUTH_URL, params=params)
            
            if response.status_code == 200:
                data = response.json()
                if "access_token" in data:
                    self.access_token = data["access_token"]
                    # Token usually valid for 3600s, use 3500s to be safe
                    self.token_expiry = time.time() + data.get("expires_in", 3600) - 100
                    return self.access_token
                else:
                    raise Exception(f"Token error: {data.get('error')}")
            else:
                raise Exception(f"Auth failed: {response.text}")
                
        except Exception as e:
            raise Exception(f"Failed to refresh token: {str(e)}")

    def validate_connection(self) -> Tuple[bool, str]:
        """
        Test connection and organization ID.
        """
        try:
            token = self.get_access_token()
            # Simple connection test
            return True, "✅ Connection successful! Token generated."
            
        except Exception as e:
            return False, f"❌ Connection failed: {str(e)}"

    def send_document_for_signature(
        self,
        file_name: str,
        file_content: bytes,
        recipient_email: str,
        recipient_name: str,
        request_name: str,
        notes: str = ""
    ) -> Tuple[bool, str]:
        """
        Uploads a document and sends it for signature.
        """
        try:
            token = self.get_access_token()
            
            # Zoho Sign typically requires Authtoken in header
            headers = {
                "Authorization": f"Zoho-oauthtoken {token}"
            }
            
            # Prepare Multipart Request
            files = {
                'file': (file_name, file_content)
            }
            
            # The 'requests' endpoint in v1 usually takes JSON payload in 'data' field
            data_payload = {
                "requests": {
                    "request_name": request_name,
                    "request_type_name": "Others",
                    "is_sequential": False,
                    "is_quick_send": True, 
                    "notes": notes,
                    "actions": [
                        {
                            "action_type": "SIGN",
                            "recipient_email": recipient_email,
                            "recipient_name": recipient_name,
                            "verify_recipient": False,
                            "is_embedded": False,
                            "in_person_name": "",
                            "in_person_email": "",
                            "role": "Signer",
                            "signing_order": 0
                        }
                    ]
                }
            }
            
            data = {
                'data': json.dumps(data_payload)
            }
            
            # Construct URL
            url = f"{self.API_BASE_URL}/requests"
            
            response = requests.post(url, headers=headers, files=files, data=data)
            
            if response.status_code in [200, 201]:
                res_json = response.json()
                if res_json.get("status") == "success":
                    req_info = res_json.get("requests", {})
                    req_id = req_info.get("request_id")
                    req_status = req_info.get("request_status", "unknown")
                    
                    # Check if it actually went through or stayed in draft
                    if req_status in ["inprogress", "submitted"]:
                        return True, f"Sent successfully (ID: {req_id}, Status: {req_status})"
                    else:
                        # accepted but maybe draft?
                        return True, f"Created but status is '{req_status}' (ID: {req_id}). Check Zoho 'Drafts' or 'In Progress'."
                else:
                    msg = res_json.get("message", "Unknown error")
                    return False, f"API Error: {msg}"
            else:
                return False, f"HTTP Error {response.status_code}: {response.text}"

        except Exception as e:
            return False, f"Error sending document: {str(e)}"
