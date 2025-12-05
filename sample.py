from Mail.src.mail import ExchangeHandler

def main():
    mail = ExchangeHandler()
    mail.send_mail_to(
        send_to="lakas@naestved.dk", 
        on_behalf_of="rpa@naestved.dk", 
        subject="Test", 
        body="Test"
        )

if __name__ == "__main__":
    main()