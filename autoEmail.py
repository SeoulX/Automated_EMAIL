import smtplib
import time
import pyautogui
from email.header import decode_header
import imaplib
import email
import speech_recognition as sr
import pyttsx3
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

recognizer = sr.Recognizer()
engine = pyttsx3.init()
engine.setProperty('rate', 150) 
engine.setProperty('volume', 1.0)

sender_email = 'aandrianbinas01@gmail.com'
password = 'muhumsuzrippngfy'

def get_email_body(msg):
    main_body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body_part = payload.decode(charset, errors='replace')
                    reply_delimiters = ["On "] 
                    for delimiter in reply_delimiters:
                        if delimiter in body_part:
                            main_body = body_part.split(delimiter)[0]
                            break
                        else:
                            main_body = body_part
                            break
    else:
        charset = msg.get_content_charset() or "utf-8"
        payload = msg.get_payload(decode=True)
        if payload:
            main_body = payload.decode(charset, errors='replace')
            reply_delimiters = ["On "]
            for delimiter in reply_delimiters:
                if delimiter in main_body:
                    main_body = main_body.split(delimiter)[0]
                    break

    return main_body.strip()

def check_new_emails():
    try:
        with imaplib.IMAP4_SSL('imap.gmail.com') as mail:
            mail.login(sender_email, password)
            mail.select('inbox')
            _, data = mail.search(None, '(UNSEEN)')
            email_ids = data[0].split()
            if email_ids:
                latest_email_id = email_ids[-1] 
                _, data = mail.fetch(latest_email_id, '(RFC822)')
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)
                subject = msg['Subject']
                sender_info = decode_header(msg['From'])[0]
                sender_name = sender_info[0].decode() if isinstance(sender_info[0], bytes) else sender_info[0]
                Rsender_email = email.utils.parseaddr(msg['From'])[1]  

                body = get_email_body(msg)

                engine.say(f"New email from {sender_name}, {Rsender_email}. Subject: {subject}")
                engine.say(f"Message: {body}") 
                engine.runAndWait()

                while True: 
                    with sr.Microphone() as source:
                        engine.say("Would you like to reply or ignore?")
                        engine.runAndWait()
                        audio = recognizer.listen(source)

                    try:
                        command = recognizer.recognize_google(audio).lower()
                        if "reply" in command:
                            automate_email_sending(Rsender_email, subject) 
                            break  
                        elif "ignore" in command:
                            engine.say("Ignoring email")
                            engine.runAndWait()
                            break 
                    except sr.UnknownValueError:
                        engine.say("Could not understand audio. Please say 'reply' or 'ignore'.")
                        engine.runAndWait()
                        continue 
                    except sr.RequestError as e:
                        engine.say(f"Could not request results; {e}")
                        engine.runAndWait()
                        break
    except imaplib.IMAP4.error as e:
        engine.say(f"An error occurred while checking emails: {e}")
        engine.runAndWait()
        time.sleep(2)

    time.sleep(10)

def automate_email_sending(recipient, reply_subject):
    gmail_url = 'https://mail.google.com/'
    gmail_window = None
    for window in pyautogui.getAllWindows():
        if gmail_url in window.title.lower() or 'aandrianbinas01@gmail.com - gmail' in window.title.lower():
            gmail_window = window
            break

    if gmail_window:
        gmail_window.activate()
        pyautogui.hotkey('ctrl', 'l')
        time.sleep(2)
        pyautogui.write(gmail_url)
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(5)
    else:
        pyautogui.press('win')
        time.sleep(2)
        pyautogui.write('chrome')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(2)  
        pyautogui.moveTo(625, 360, duration=0.5)
        time.sleep(2)
        pyautogui.click()
        time.sleep(2)
        pyautogui.write(gmail_url)
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(5)
    
    pyautogui.press('o')
    time.sleep(2)
    
    pyautogui.press('r')
    time.sleep(2)

    # if recipient:
    #     pyautogui.write(recipient) 
    #     time.sleep(2)
    #     pyautogui.press('enter') 
    #     time.sleep(1)
    # pyautogui.press('tab')
    # time.sleep(2)

    # if reply_subject:
    #     pyautogui.write(f"Re: {reply_subject}")
    #     time.sleep(2)
    # pyautogui.press('tab')
    # time.sleep(2)

    with sr.Microphone() as source:
        engine.say("Please dictate your email:")
        engine.runAndWait()
        audio = recognizer.listen(source)

    try:
        email_body = recognizer.recognize_google(audio)
        pyautogui.write(email_body)
    except sr.UnknownValueError:
        engine.say("Could not understand audio. Please try dictating again.")
        engine.runAndWait()
        return
    except sr.RequestError as e:
        engine.say(f"Could not request results from Google Speech Recognition service; {e}")
        engine.runAndWait()
        return

    pyautogui.hotkey('ctrl', 'enter')
    engine.say("Email sent successfully!")
    engine.runAndWait()

while True:
    check_new_emails()
    time.sleep(8) 
