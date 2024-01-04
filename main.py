import os
import webbrowser
import datetime
import openai
import speech_recognition as sr
import win32com.client
from config import apikey
import imaplib
import email
from email.header import decode_header

speaker = win32com.client.Dispatch("SAPI.SpVoice")

'''
def read_emails():
    email_username = "nikhil.cheryala@spsu.ac.in"
    email_password = ""
    imap_server = "imap.example.com"  # Use the IMAP server of your email provider

    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_username, email_password)
    except Exception as e:
        return "Failed to connect and log in to email account."

    mailbox = "INBOX"
    mail.select(mailbox)

    result, data = mail.search(None, "ALL")
    email_contents = []

    if result == "OK":
        email_ids = data[0].split()
        for email_id in email_ids:
            result, message_data = mail.fetch(email_id, "(RFC822)")
            if result == "OK":
                raw_email = message_data[0][1]
                msg = email.message_from_bytes(raw_email)

                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")

                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))

                        if "attachment" not in content_disposition:
                            body = part.get_payload(decode=True).decode()
                else:
                    body = msg.get_payload(decode=True).decode()

                email_contents.append(f"Subject: {subject}\nMessage Body: {body}\n")

    mail.logout()
    return email_contents
'''


chatStr = ""


def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Nikhil: {query}\n Jarvis: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    speaker.Speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]


def ai(prompt):
    openai.api_key = apikey
    text = f"openAI response for Prompt: {prompt} \n *************************\n\n"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("openai files"):
        os.makedirs("openai" , exist_ok=True)
    # with open(f"openai/prompt- {random.randint(1, 2343434356)}", "w") as f:
    with open(f"openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)



def commmand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold = 0.4
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"


if __name__ == '__main__':
    print('PyCharm')
    speaker.Speak("Hello I am neha A.I")
    while True:
        print("Listening....")
        query = commmand()
        # speaker.Speak(query)
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["google", "https://www.google.com"], ["spsu","https://www.spsu.ac.in/"], ["facebook", "https://www.facebook.com/"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
        if "open music" in query:
            musicPath = r"\Users\nithin\Downloads\music.mp3"
            os.startfile(musicPath)
            # strfTime = datetime.datetime.now().strftime("%H:%M:%S")
        elif "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.Speak(f"Sir the time is now {strfTime}")
        elif "open instagram".lower() in query.lower():
            os.system(f"start \\Users\\nithin\\OneDrive\\Documents\\Desktop\\Instagram.lnk")
        elif "open hotstar".lower() in query.lower():
            os.system(f"start \\Users\\nithin\\OneDrive\\Documents\\Desktop\\Hotstar.lnk")

        elif "Using artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "Jarvis Quit".lower() in query.lower():
            exit()

        elif "reset chat".lower() in query.lower():
            chatStr = ""
        else:
            print("Chatting...")
            chat(query)

