"""
DocuSign Handler Module
Handles basic DocuSign API interactions using JWT Authentication.
"""
import base64
import os
import requests
from docusign_esign import ApiClient, EnvelopesApi, EnvelopeDefinition, Document, Signer, CarbonCopy, SignHere, Tabs, Recipients
from docusign_esign.client.api_exception import ApiException

class DocuSignHandler:
    def __init__(self, integration_key, user_id, account_id, private_key_path, base_url="https://demo.docusign.net"):
        self.integration_key = integration_key
        self.user_id = user_id
        self.account_id = account_id
        self.private_key_path = private_key_path
        self.base_url = base_url
        self.api_client = ApiClient()
        self.api_client.set_base_path(base_url + "/restapi")
        # Always use Demo for now as per user screenshot
        self.api_client.set_oauth_host_name("account-d.docusign.com")

        self._jwt_auth()

    def _jwt_auth(self):
        """Authenticate using JWT Grant to get an access token."""
        try:
            with open(self.private_key_path, "r") as key_file:
                content = key_file.read()
                
            # Extract pure private key if mixed content
            if "-----BEGIN RSA PRIVATE KEY-----" in content:
                start = content.find("-----BEGIN RSA PRIVATE KEY-----")
                end = content.find("-----END RSA PRIVATE KEY-----") + len("-----END RSA PRIVATE KEY-----")
                private_key = content[start:end].encode("ascii")
            else:
                # Assume it's pure bytes if not found (fallback)
                with open(self.private_key_path, "rb") as key_file:
                    private_key = key_file.read()

            # Scopes needed: signature strict impersonation
            scopes = ["signature", "impersonation"]
            
            # Request JWT token
            token_response = self.api_client.request_jwt_user_token(
                client_id=self.integration_key,
                user_id=self.user_id,
                oauth_host_name=self.api_client.oauth_host_name,
                private_key_bytes=private_key,
                expires_in=3600,
                scopes=scopes
            )
            
            self.access_token = token_response.access_token
            # self.api_client.set_default_header("Authorization", "Bearer " + self.access_token)
            
            # Very important: Update base path based on account info (though we passed it, good to verify)
            # For simplicity, we stick to the provided base_url + /restapi
            self.api_client.host = self.base_url + "/restapi"
            self.api_client.set_default_header("Authorization", "Bearer " + self.access_token)

            return True

        except ApiException as e:
            # If consent is needed, return the consent URL
            if "consent_required" in str(e).lower():
                consent_url = f"https://{self.api_client.oauth_host_name}/oauth/auth?response_type=code&scope=signature%20impersonation&client_id={self.integration_key}&redirect_uri=https://apps.docusign.com/"
                raise Exception(f"Consent Required. Please visit this URL to grant consent: {consent_url}")
            raise e

    def get_signing_link(self, signer_email, signer_name, document_name, document_bytes):
        """
        Creates an envelope and generates a recipient view (signing link) for embedded signing.
        However, since we want to send the link via *our* email, strict embedded signing usually 
        redirects immediately. 
        
        A better approach for "Emailing a link" is:
        1. Create Envelope (status='sent') with the recipient marked as 'embedded' (clientUserId set).
        2. Generate the Recipient View URL.
        """
        
        # 1. Create Envelope Definition
        envelope_definition = EnvelopeDefinition(email_subject=f"Please sign: {document_name}")
        
        # Create Document
        b64_doc = base64.b64encode(document_bytes).decode("utf-8")
        doc = Document(
            document_base64=b64_doc,
            name=document_name,
            file_extension=document_name.split('.')[-1],
            document_id="1"
        )
        envelope_definition.documents = [doc]
        
        # Create Signer
        # client_user_id is CRITICAL for embedded signing. It tells DocuSign "don't email them, I will handle it".
        signer = Signer(
            email=signer_email,
            name=signer_name,
            recipient_id="1",
            routing_order="1",
            client_user_id=signer_email # Using email as unique client ID for simplicity
        )
        
        # Add SignHere tab positioned to match template signature area
        # Based on user's document: after "Legal Officer" text, in red box area
        sign_here = SignHere(
            document_id="1",
            page_number="1",
            x_position="220",  # Centered in signature box
            y_position="680",  # Positioned in red box area
            tab_label="Signature"
        )
        
        signer.tabs = Tabs(sign_here_tabs=[sign_here]) 

        envelope_definition.recipients = Recipients(signers=[signer])
        envelope_definition.status = "sent" # Send immediately
        
        # 2. Create Envelope
        envelopes_api = EnvelopesApi(self.api_client)
        results = envelopes_api.create_envelope(account_id=self.account_id, envelope_definition=envelope_definition)
        envelope_id = results.envelope_id
        
        # 3. Generate Signing Link (Recipient View)
        recipient_view_request = {
            "authenticationMethod": "email",
            "clientUserId": signer_email,
            "recipientId": "1",
            "returnUrl": "https://www.google.com", # Redirect after signing
            "userName": signer_name,
            "email": signer_email
        }
        
        view_results = envelopes_api.create_recipient_view(
            account_id=self.account_id,
            envelope_id=envelope_id,
            recipient_view_request=recipient_view_request
        )
        
        return view_results.url, envelope_id
