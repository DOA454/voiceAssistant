import requests
from datetime import datetime


# api_key = '38af91f97ea3a0243ec6cb45019bfb4d'
# city = input("Enter City:")
# date_str = input("Enter a date in YYYY-MM-DD format: ")
# date_obj = datetime.strptime(date_str, "%Y-%m-%d")

# url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"

# response = requests.get(url)

# if response.status_code == 200:
#     data = response.json()
#     if date_obj.date() == datetime.now().date():
#         temp = data["main"]["temp"]
#         feels_like = data["main"]["feels_like"]
#         description = data["weather"][0]["description"]
#         print(f"Current weather in {city}: {description}. Temperature: {temp}°C. Feels like: {feels_like}°C.")
#     else:
#         timestamp = int(date_obj.timestamp())
#         url = f"http://api.openweathermap.org/data/2.5/onecall/timemachine?lat={data['coord']['lat']}&lon={data['coord']['lon']}&dt={timestamp}&appid={api_key}&units=metric"
#         response = requests.get(url)
#         if response.status_code == 200:
#             data = response.json()
#             temp = data["current"]["temp"]
#             feels_like = data["current"]["feels_like"]
#             description = data["current"]["weather"][0]["description"]
#             print(f"Weather in {city} on {date_str}: {description}. Temperature: {temp}°C. Feels like: {feels_like}°C.")
#         else:
#             print(f"Error retrieving weather data. Error code: {response.status_code}")
# else:
#     print(f"Error retrieving weather data. Error code: {response.status_code}")


api_key = '38af91f97ea3a0243ec6cb45019bfb4d'
city = input("Enter City:")

url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"

response = requests.get(url)

if response.status_code == 200:
    data = response.json()
    temp = data["main"]["temp"]
    feels_like = data["main"]["feels_like"]
    description = data["weather"][0]["description"]
    print(f"Current weather in {city}: {description}. Temperature: {temp}°C. Feels like: {feels_like}°C.")
else:
    print(f"Error retrieving weather data. Error code: {response.status_code}")