import win32com.client
import speech_recognition as sr
import webbrowser
import os
import datetime
import openai
import requests
from Config import ApiKey
from flask import Flask, request, jsonify

speaker = win32com.client.Dispatch("SAPI.SpVoice")
chatStr = ''

app = Flask(__name__)

def say(text):
    # while True:
        # print("Enter the word you want to speak it out by speaker.")
        # text = input()
        print(text)
        speaker.speak(text)
        # if s == "Allah Hafiz":
        #     break 

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.8  # Adjust the pause threshold according to your environment
        r.energy_threshold = 4000  # Adjust the energy threshold according to your environment
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-PK")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            return "Sorry, I couldn't understand what you said."
        except sr.RequestError:
            return "Sorry, I'm currently unable to process your request."

def openwebsite():
    while True:
        print("Listening...")
        website_name = takeCommand()

        if website_name:
            website_url = f"https://www.{website_name.lower()}.com"
            say(f"Shaafeeq is Opening {website_name.lower()} sir...")
            webbrowser.open(website_url)
            return
        return
    return

def playMusic():
    while True:
        print("Listening...")
        folder_name = takeCommand()

        if folder_name:
            folder_path = f"C:\\Users\\muham\\Music\\Playlists\\Playlists from Groove Music\\{folder_name}.zpl"
            say(f"Shaafeeq is Opening {folder_name.lower()} sir...")
            try:
                os.startfile(folder_path)
                print(f"Opened folder: {folder_path}")
            except Exception as e:
                print(f"Error opening folder: {e}")
                return
        return
    return

def chat(prompt):
    global chatStr
    print(chatStr)
    openai.api_key = ApiKey
    chatStr += f"Ammar: {prompt}\n Shaafeeq: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: Wrap this inside of a  try catch block
    say(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

def ai(prompt):
    openai.api_key = ApiKey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    
    text += response["choices"][0]["text"]
    if not os.path.exists("F:\\Python\\JARVIS\\OpenAi"):
        os.mkdir("F:\\Python\\JARVIS\\OpenAi")

    with open(f"F:\\Python\\JARVIS\\OpenAi\\{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
        f.write(text)

    print("done")


def get_weather(city):
    api_key = "4ec8f8756c9386f12e816acbb896eb8a"
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}"

    try:
        response = requests.get(url)
        data = response.json()

        if 'weather' in data:
            weather_description = data['weather'][0]['description']
            temperature = data['main']['temp']
            temp = temperature - 273.15
            humidity = data['main']['humidity']

            say(f"Weather in {city}:")
            say(f"Description: {weather_description}")
            say(f"Temperature: {temp} C")
            say(f"Humidity: {humidity}%")
        else:
            say("Weather data not available for the specified location.")
    except requests.exceptions.RequestException as e:
        say(f"An error occurred: {e}")

def get_news(category):
    api_key = "79c41f0094ed41008f2b1b3187d4451e"
    url = f"https://newsapi.org/v2/top-headlines?country=us&category={category}&apiKey={api_key}"

    try:
        response = requests.get(url)
        data = response.json()

        if data['status'] == 'ok':
            articles = data['articles']
            for article in articles:
                title = article['title']
                description = article['description']
                url = article['url']

                say(title)
                say(description)
                print("URL:", url)
                print("------")
        else:
            say("Failed to fetch news.")

    except requests.exceptions.RequestException as e:
        say(f"An error occurred: {e}")

def get_newsSub(subcategory):
    api_key = "79c41f0094ed41008f2b1b3187d4451e"

    say(f"Fetching {subcategory} news...")
    get_news.__get__(category)
    url = f"https://newsapi.org/v2/top-headlines?category={category}&q={subcategory}&apiKey={api_key}"

    try:
        response = requests.get(url)
        data = response.json()

        if data['status'] == 'ok':
            articles = data['articles']
            for article in articles:
                title = article['title']
                description = article['description']
                url = article['url']
                if not None:
                    say(title)
                if description is not None:
                    say(description)
                else:                
                    print("URL:", url)
                    print("------")
        else:
            say("Failed to fetch news.")

    except requests.exceptions.RequestException as e:
        say(f"An error occurred: {e}")

if __name__ == '__main__':
    app.run()
    say("Shaafeeq AI is here for your service sir")
    say("How can i help you?")
    # print(ApiKey)
    while True:
        query = takeCommand()
        # say(query)

        if "Open Website".lower() in query.lower():
            say("Which website do you want to open sir...")
            openwebsite()

        elif "Listen Music".lower() in query.lower():
            say("Which music do you want to listen sir...")
            playMusic()

        elif "time".lower() in query.lower():
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir the time is {strfTime}")

        elif "date".lower() in query.lower():
            strfDate = datetime.datetime.now().strftime("%Y-%m-%d")
            say(f"Sir the date is {strfDate}")

        elif "Artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "weather".lower() in query.lower():
            say("Which city do you want to know the weather sir...")
            city = takeCommand()
            get_weather(city)

        elif "news".lower() in query.lower():
            say("Which news category do you want to hear sir...")
            category = takeCommand()
            # get_news(category)
            say(f"Which type of {category} news do you want to hear sir...")
            subcategory = takeCommand()
            get_newsSub(subcategory)

        elif "Allah Hafiz".lower() in query.lower():
            say("Allah Hafiz")
            break 

        else:
            chat(query)











@app.route('/sendMessage', methods=['POST'])
def process_user_message():
    user_message = request.form.get('userMessage')
    
    # Code to interact with your AI model and get the AI response
    ai_response = "AI response will be generated based on user_message"

    # Code to fetch weather information from Weather Map API
    weather_info = "Weather information will be fetched from Weather Map API"

    # Code to fetch latest news from News API
    latest_news = "Latest news will be fetched from News API"

    response = {
        'aiResponse': ai_response,
        'weatherInfo': weather_info,
        'latestNews': latest_news
    }

    return jsonify(response)