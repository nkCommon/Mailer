import logging
import requests
import msal
import os
from dotenv import load_dotenv
import datetime as dt
from typing import Optional
# import asyncio
### using app registration AbsenceMailReader
class ExchangeHandler:
    """_summary_
    ExchangeHandler is a class that handles the exchange of mail with the Exchange server
    """
    def __init__(self, env_file: str = '.env'):
        load_dotenv(env_file)
        
        # === CONFIG ===
        self.TENANT_ID = os.getenv("MAIL_TENANT_ID")
        self.CLIENT_ID = os.getenv("MAIL_CLIENT_ID")
        self.CLIENT_SECRET = os.getenv("MAIL_CLIENT_SECRET")
        self.USER_ID = os.getenv("MAIL_USER_ID", "rpa@naestved.dk")  

        self.AUTHORITY = f"https://login.microsoftonline.com/{self.TENANT_ID}"
        self.SCOPE = ["https://graph.microsoft.com/.default"]
        self.GRAPH_BASE = "https://graph.microsoft.com/v1.0"

        self.app = msal.ConfidentialClientApplication(
            self.CLIENT_ID, 
            authority=self.AUTHORITY, 
            client_credential=self.CLIENT_SECRET)
        result = self.app.acquire_token_silent(self.SCOPE, account=None)
        if not result:
            result = self.app.acquire_token_for_client(scopes=self.SCOPE)

        self.access_token = result["access_token"]
        self.headers = {"Authorization": f"Bearer {self.access_token}"}


    def send_mail(self, data:dict):
        """_summary_
        send_mail is a method that sends a mail
        Args:
            data (dict): mail data dict
        """
        logger = logging.getLogger(__name__)    
        logger.info(f"Sending mail on behalf of {data['on_behalf_of']} to {data['send_to']}")
        
        send_to = data['send_to']
        on_behalf_of = data['on_behalf_of']
        subject = f"{data['name']} er fraværende." if data['subject'] in self.ALLOWED_MAIL_SUBJECTS_WITH_CC else f"{data['name']} er til stede igen."
        body = "...."
        
        self.send_mail_to(send_to=send_to, on_behalf_of=on_behalf_of, subject=subject, body=body)
        
    def send_warning_error_mail(self, data:dict, subject:str, body:str):
        """_summary_
        send_mail is a method that sends a mail
        Args:
            data (dict): mail data dict
            error_message (str): error message
        """
        logger = logging.getLogger(__name__)    
        

        if data['cc']:
            send_to = f"{data['from']};{data['cc']}"
        else:
            send_to = data['from']

        logger.info(f"Sending warning/error mail on behalf of {data['on_behalf_of']} to {send_to}")
            
        on_behalf_of = data['on_behalf_of']
        
        body_content = f"* OBS * tjek OPUS for at sikre at fraværet er oprettet korrekt. \n\n{body}"
        
        self.send_mail_to(send_to=send_to, on_behalf_of=on_behalf_of, subject=subject, body=body_content)

    def send_mail_to_employee(self, data:dict):
        """_summary_
        send_mail_to_employee is a method that sends a mail to the employee
        Args:
            data (dict): mail data dict
        """
        if data['subject'] in self.ALLOWED_MAIL_SUBJECTS_WITH_NO_CC:
            return
        logger = logging.getLogger(__name__)    
        user_upn = data['cc'] if data['cc'] else data['from']
        logger.info(f"Sending mail to employee {user_upn}")
        
        # Build recipient list correctly

        body = self.MAIL_MESSAGE_TO_EMPLOYEE_SICK if data['subject'] == 'Sygemelding' else self.MAIL_MESSAGE_TO_EMPLOYEE_CHILD_SICK
        subject = f"Sygefravær er registreret"
        on_behalf_of = data['on_behalf_of']
        
        self.send_mail_to(send_to=user_upn, on_behalf_of=on_behalf_of, subject=subject, body=body)
 
    def send_mail_to(self, send_to: str, on_behalf_of:str, subject: str, body: str)->dict:
        """_summary_
        send_mail_to is a method that sends a mail to a recipient
        Args:
            send_to (str): recipient email address
            on_behalf_of (str): email address of the sender
            subject (str): mail subject
            body (str): mail body
        """
        logger = logging.getLogger(__name__)    
        
        # Build recipient list correctly
        to_recipients = [
            {"emailAddress": {"address": addr.strip()}}
            for addr in send_to.replace(",", ";").split(";")
            if addr.strip()
        ]
        url = f"{self.GRAPH_BASE}/users/{on_behalf_of}/sendMail"

        message = {
        "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": f"{body}",
                },
                "from": {
                    "emailAddress": {"address": on_behalf_of}
                },
                "sender": {
                    "emailAddress": {"address": on_behalf_of}
                },
                "toRecipients": to_recipients,
            },
            "saveToSentItems": "true"
        }
        resp = requests.post(
                url,
                headers=self.headers,
                json=message,
                timeout=30,
            )
        if resp.status_code >= 300:
            logger.error(f"Cannot send mail: Graph error {resp.status_code}: {resp.text}")
            raise RuntimeError(f"Cannot send mail: Graph error {resp.status_code}: {resp.text}")
        
