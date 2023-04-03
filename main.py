import requests
import json
import win32com.client as wn
try:
    if __name__ == '__main__':
        while True:
            exit_message = wn.Dispatch("SAPI.SpVoice")
            print("\n\t\t Welcome to Weather App 1.1. Created by Ahmad\n")
            print('''Instructions: 
            1)Please Enter Your city name with correct Spelling.
            2)Press 0 when you want to exit from the weather app \n''')

            city = input("Enter Name of Your City : ")
            if city == "0":
                print("Goodbye, thanks for using Weather App, See you soon again")
                exit_message.Speak("Goodbye, thanks for using Weather App, See you soon again")
                break
            url = f"https://api.weatherapi.com/v1/current.json?key=65ad539364ff4768924163315230204&q={city}"

            s = requests.get(url)
            weatherdisc = json.loads(s.text)
            test = weatherdisc["current"]["is_day"]
            if test == 0:
                time = "Night"
            else:
                time = "Day"
            print(f'''
            City Name : {weatherdisc["location"]["name"]}
            Region : {weatherdisc["location"]["region"]}
            Country : {weatherdisc["location"]["country"]}
            Temprature : {weatherdisc["current"]["temp_c"]}(C) or {weatherdisc["current"]["temp_f"]}(f)
            Feels Like : {weatherdisc["current"]["feelslike_c"]}(C)
            Wind Speed : {weatherdisc["current"]["wind_kph"]} (Kph)
            Wind Degree : {weatherdisc["current"]["wind_degree"]}
            Preciption : {weatherdisc["current"]["precip_in"]}
            Humidity : {weatherdisc["current"]["humidity"]}%
            Cloud : {weatherdisc["current"]["cloud"]}%
            Time Zone : {weatherdisc["location"]["tz_id"]}
            Local Time : {weatherdisc["location"]["localtime"]}
            Day/Night : It's {time} Time in {city}
            Last Updated : {weatherdisc["current"]["last_updated"]}
            ''')
            responce = int(input('''You Want to Check Weather of other city (Press 1 for "Yes" and 0 for "exit") : '''))
            if responce == 1:
                pass
            elif responce == 0:
                print("Goodbye, thanks for using Weather App, See you soon again")
                exit_message.Speak("Goodbye, thanks for using Weather App, See you again soon ")
                break
            else:
                print("Invalid Choice, Thanks,  See you again soon")
                exit_message.Speak("Invalid Choice, Thanks,  See you again soon")
                break


except Exception as e:
    print("Something Went Wrong, Wrong Input : Error = ", e)
