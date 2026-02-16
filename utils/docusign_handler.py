"""
DocuSign Handler Module
Handles basic DocuSign API interactions using JWT Authentication.
"""
import base64
import os
import requests
from docusign_esign import ApiClient, EnvelopesApi, EnvelopeDefinition, Document, Signer, CarbonCopy, SignHere, Tabs, Recipients, RecipientEmailNotification
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
            
            # Fetch the correct base URI for the account
            # We use requests directly to ensure we hit the correct OAuth host
            try:
                user_info_url = f"https://{self.api_client.oauth_host_name}/oauth/userinfo"
                response = requests.get(
                    user_info_url,
                    headers={"Authorization": "Bearer " + self.access_token}
                )
                # response.raise_for_status() # Optional: ignore if we want to fallback
                
                if response.status_code == 200:
                    user_info = response.json()
                    accounts = user_info.get('accounts', [])
                    target_account = next((acc for acc in accounts if acc['account_id'] == self.account_id), None)
                    
                    if target_account:
                        self.base_url = target_account['base_uri']
                        self.api_client.host = self.base_url + "/restapi"
                    else:
                        # Fallback
                        self.api_client.host = self.base_url + "/restapi"
                else:
                    self.api_client.host = self.base_url + "/restapi"

            except Exception as e:
                # Fallback if discovery fails
                self.api_client.host = self.base_url + "/restapi"
                
            self.api_client.set_default_header("Authorization", "Bearer " + self.access_token)

            return True

        except ApiException as e:
            # If consent is needed, return the consent URL
            if "consent_required" in str(e).lower():
                consent_url = f"https://{self.api_client.oauth_host_name}/oauth/auth?response_type=code&scope=signature%20impersonation&client_id={self.integration_key}&redirect_uri=https://apps.docusign.com/"
                raise Exception(f"Consent Required. Please visit this URL to grant consent: {consent_url}")
            raise e

    def send_envelope(self, signer_email, signer_name, documents, subject=None, body=None, embedded=False, cc_emails=None):
        """
        Creates and sends an envelope.
        
        Args:
            signer_email (str): Email of the signer.
            signer_name (str): Name of the signer.
            documents (list): List of tuples (filename, file_bytes).
            subject (str): Email subject.
            body (str): Email body (blurb).
            embedded (bool): If True, generates a short-lived link for embedded signing. 
                             If False, Docusign sends the email (link valid until envelope expires).
        
        Returns:
            tuple: (signing_url, envelope_id)
                   signing_url will be None if embedded=False.
        """
        
        # 1. Create Envelope Definition
        envelope_definition = EnvelopeDefinition()
        envelope_definition.email_subject = subject if subject else "Please Sign"
        envelope_definition.email_blurb = body if body else "Please sign the attached documents."
        
        # Default Expiration to 120 days (approx 4 months) and warn after 110
        # This addresses "never expire" as best as possible within DocuSign limits
        from docusign_esign import Notification, Expirations
        expirations = Expirations(expire_after="120", expire_warn="110")
        notification = Notification(expirations=expirations)
        envelope_definition.notification = notification
        
        # 1b. Set Envelope Blurb explicitly as well
        envelope_definition.email_subject = subject
        envelope_definition.email_blurb = body
        doc_objects = []
        for i, (filename, file_bytes) in enumerate(documents):
            b64_doc = base64.b64encode(file_bytes).decode("utf-8")
            
            # Safe extension extraction
            _, ext = os.path.splitext(filename)
            file_extension = ext.lstrip('.').lower()
            if not file_extension:
                file_extension = "pdf" # Default fallback
                
            doc = Document(
                document_base64=b64_doc,
                name=filename,
                file_extension=file_extension,
                document_id=str(i + 1)
            )
            doc_objects.append(doc)
            
        envelope_definition.documents = doc_objects
        
        # 3. Create Signer
        signer = Signer(
            email=signer_email,
            name=signer_name,
            recipient_id="1",
            routing_order="1"
        )
        
        # Explicitly set Email Notification for the signer to ensure content appears
        if not embedded: # Only relevant for email delivery
            signer.email_notification = RecipientEmailNotification(
                email_subject=subject if subject else "Please Sign",
                email_body=body if body else "Please review and sign."
            )
        
        # If embedded, set client_user_id (suppresses email, creates short-lived link)
        if embedded:
            signer.client_user_id = signer_email
            
        # 4. Add Tabs (SignHere)
        # We use Anchor Tagging to place the signature exactly where {{Signature:Recipient1}} appears in the doc.
        sign_here_anchor = SignHere(
            anchor_string="{{Signature:Recipient1}}",
            anchor_units="pixels",
            anchor_y_offset="-10", # Adjust slightly up to align with text baseline
            anchor_x_offset="0",
            tab_label="Signature_Anchor"
        )
        
        # Fallback tab in case {Signature} is missing from the document
        # This ensures the envelope doesn't fail or have 0 signature fields
        sign_here_fallback = SignHere(
            document_id="1",
            page_number="1",
            x_position="100", 
            y_position="100", # Top left corner as safety
            tab_label="Signature_Fallback" 
            # Note: We can't easily make this "conditional".
            # If both exist, user might be asked to sign twice.
            # So for now, let's stick to the User's Request: Use {Signature}.
            # If they don't include it, they should!
        )
        
        # Let's use the Anchor strategy primarily.
        signer.tabs = Tabs(sign_here_tabs=[sign_here_anchor])
        
        # Add CC Recipients if provided
        carbon_copies = []
        if cc_emails:
            for idx, cc_email in enumerate(cc_emails):
                cc = CarbonCopy(
                    email=str(cc_email).strip(),
                    name="CC", 
                    recipient_id=f"2{idx}",
                    routing_order="2"
                )
                carbon_copies.append(cc)
        
        envelope_definition.recipients = Recipients(signers=[signer], carbon_copies=carbon_copies)
        envelope_definition.status = "sent"
        
        # 5. Create Envelope
        envelopes_api = EnvelopesApi(self.api_client)
        results = envelopes_api.create_envelope(account_id=self.account_id, envelope_definition=envelope_definition)
        envelope_id = results.envelope_id
        
        signing_url = None
        if embedded:
            recipient_view_request = {
                "authenticationMethod": "email",
                "clientUserId": str(signer_email),
                "recipientId": "1",
                "returnUrl": "https://www.google.com", # Redirect after signing
                "userName": str(signer_name),
                "email": str(signer_email)
            }
            
            view_results = envelopes_api.create_recipient_view(
                account_id=self.account_id,
                envelope_id=envelope_id,
                recipient_view_request=recipient_view_request
            )
            signing_url = view_results.url
            
        return signing_url, envelope_id
