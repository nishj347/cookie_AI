#Author :  nishj347 @github
import speech_recognition as sr
import win32com.client
import webbrowser
import openai
import datetime
import cv2

speaker = win32com.client.Dispatch("SAPI.SpVoice")


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.8
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"You said: {query}")
    except Exception as e:
        print("Say that again please")
        return None
    return query


def open_camera():
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        if not ret:
            print("Failed to grab frame")
            break
        cv2.imshow('Camera', frame)

        # Exit the camera window when 'q' key is pressed
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()


if __name__ == '__main__':
    print('pycharm')
    speaker.Speak("Hello, I'm cookie AI")
    while True:
        query = takecommand()
        if query:
            sites = [["youtube", "https://www.youtube.com/"], ["google", "https://www.google.com/"],
                     ["notion", "https://www.notion.so/"],["spotify","https://www.spotify.com/"],["github","https://github.com/nishj347"]]
            for site in sites:
                if f"open {site[0]}".lower() in query.lower():
                    speaker.Speak(f"Opening {site[0]}, ma'am")
                    webbrowser.open(site[1])

            if "the time" in query.lower():
                hour = datetime.datetime.now().strftime("%H")
                minute = datetime.datetime.now().strftime("%M")
                speaker.Speak(f"Ma'am, the time is {hour} hours and {minute} minutes")

            if "open camera".lower() in query.lower():
                speaker.Speak("Opening the camera, ma'am")
                open_camera()
