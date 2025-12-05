from Mail.src.mail import ExchangeHandler
import os
from dotenv import load_dotenv

def main():
    load_dotenv()
    
    MAIL_TENANT_ID = os.getenv("MAIL_TENANT_ID")
    MAIL_CLIENT_ID = os.getenv("MAIL_CLIENT_ID")
    MAIL_CLIENT_SECRET = os.getenv("MAIL_CLIENT_SECRET")
    MAIL_USER_ID = os.getenv("MAIL_USER_ID")

    mail = ExchangeHandler(
        tenant_id=MAIL_TENANT_ID,
        client_id=MAIL_CLIENT_ID,
        client_secret=MAIL_CLIENT_SECRET,
        user_id=MAIL_USER_ID
    )

    mail.send_mail_to(
        send_to="lakas@naestved.dk", 
        on_behalf_of="rpa@naestved.dk", 
        subject="Test", 
        body="Test"
        )

if __name__ == "__main__":
    main()