"""
Email Handler Module
Handles SMTP email sending with validation, progress tracking, and error handling.
"""

import smtplib
import time
from email.message import EmailMessage
from email.utils import formataddr
from typing import List, Dict, Tuple, Callable, Optional
import re


class EmailHandler:
    """Handle email operations including validation and sending."""

    def __init__(
        self,
        smtp_server: str,
        smtp_port: int,
        sender_email: str,
        sender_password: str,
        sender_name: Optional[str] = None,
    ):
        """
        Initialize email handler with SMTP configuration.

        Args:
            smtp_server: SMTP server address (e.g., 'smtp.gmail.com')
            smtp_port: SMTP port (e.g., 587 for TLS)
            sender_email: Sender's email address
            sender_password: Sender's password or app password
            sender_name: Optional sender display name
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.sender_name = sender_name or sender_email

    @staticmethod
    def validate_email(email: str) -> bool:
        """
        Validate email format using regex.

        Args:
            email: Email address to validate

        Returns:
            True if valid, False otherwise
        """
        if not email or not isinstance(email, str):
            return False

        # Basic email regex pattern
        pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        return bool(re.match(pattern, email.strip()))

    def test_connection(self) -> Tuple[bool, str]:
        """
        Test SMTP connection and authentication.

        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            server = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=10)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.quit()
            return True, "✅ Connection successful!"
        except smtplib.SMTPAuthenticationError:
            return False, "❌ Authentication failed. Check your email and password."
        except smtplib.SMTPException as e:
            return False, f"❌ SMTP error: {str(e)}"
        except Exception as e:
            return False, f"❌ Connection failed: {str(e)}"

    def send_personalized_email(
        self,
        to_email: str,
        subject: str,
        body: str,
        attachment_filename: str,
        attachment_data: bytes,
    ) -> Tuple[bool, str]:
        """
        Send a single personalized email with attachment.

        Args:
            to_email: Recipient email address
            subject: Email subject
            body: Email body (plain text)
            attachment_filename: Name of the attachment file
            attachment_data: Binary data of the attachment

        Returns:
            Tuple of (success: bool, error_message: str or empty)
        """
        try:
            # Create message
            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = formataddr((self.sender_name, self.sender_email))
            msg["To"] = to_email
            msg.set_content(body)

            # Add attachment
            msg.add_attachment(
                attachment_data,
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename=attachment_filename,
            )

            # Send email
            with smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=30) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)

            return True, ""

        except smtplib.SMTPRecipientsRefused:
            return False, "Invalid recipient email address"
        except smtplib.SMTPAuthenticationError:
            return False, "Authentication failed"
        except smtplib.SMTPException as e:
            return False, f"SMTP error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"

    def send_batch_emails(
        self,
        email_data_list: List[Dict],
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        delay_seconds: float = 1.5,
    ) -> Dict[str, any]:
        """
        Send multiple personalized emails with progress tracking.

        Args:
            email_data_list: List of dicts with keys:
                - 'to_email': recipient email
                - 'subject': email subject
                - 'body': email body
                - 'attachment_filename': attachment filename
                - 'attachment_data': attachment binary data
                - 'row_index': original row index (for tracking)
            progress_callback: Optional callback function(current, total, status)
            delay_seconds: Delay between emails to avoid rate limiting

        Returns:
            Dict with results:
                - 'total': total emails
                - 'sent': successfully sent count
                - 'failed': failed count
                - 'skipped': skipped count
                - 'failed_details': list of failed email details
        """
        results = {
            "total": len(email_data_list),
            "sent": 0,
            "failed": 0,
            "skipped": 0,
            "failed_details": [],
        }

        for idx, email_data in enumerate(email_data_list):
            # Update progress
            if progress_callback:
                progress_callback(
                    idx + 1,
                    len(email_data_list),
                    f"Sending to {email_data['to_email']}...",
                )

            # Send email
            success, error_msg = self.send_personalized_email(
                to_email=email_data["to_email"],
                subject=email_data["subject"],
                body=email_data["body"],
                attachment_filename=email_data["attachment_filename"],
                attachment_data=email_data["attachment_data"],
            )

            if success:
                results["sent"] += 1
            else:
                results["failed"] += 1
                results["failed_details"].append(
                    {
                        "row_index": email_data.get("row_index", idx),
                        "email": email_data["to_email"],
                        "error": error_msg,
                    }
                )

            # Delay to avoid rate limiting (except for last email)
            if idx < len(email_data_list) - 1:
                time.sleep(delay_seconds)

        return results

    @staticmethod
    def render_template(template: str, data: Dict[str, any]) -> str:
        """
        Render template string with data placeholders.

        Args:
            template: Template string with {placeholder} format
            data: Dictionary of placeholder values

        Returns:
            Rendered string
        """
        result = template
        for key, value in data.items():
            placeholder = f"{{{key}}}"
            result = result.replace(placeholder, str(value))
        return result

    @staticmethod
    def get_template_placeholders(template: str) -> List[str]:
        """
        Extract placeholders from template string.

        Args:
            template: Template string with {placeholder} format

        Returns:
            List of placeholder names
        """
        pattern = r"\{([^}]+)\}"
        return re.findall(pattern, template)
