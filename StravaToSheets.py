import requests
import urllib3
import pandas as pd
from openpyxl import Workbook

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

auth_url = "https://www.strava.com/oauth/token"
activities_url = "https://www.strava.com/api/v3/athlete/activities"

payload = {
    'client_id': "CLIENT_ID",
    'client_secret': 'CLIENT_SECRET',
    'refresh_token': 'REFRESH_TOKEN',
    'grant_type': "refresh_token",
    'f': 'json'
}

print("Requesting Token...\n")
res = requests.post(auth_url, data=payload, verify=False)
access_token = res.json().get('access_token')

if not access_token:
    print("Failed to retrieve access token.")
    exit()

print("Access Token = {}\n".format(access_token))
header = {'Authorization': 'Bearer ' + access_token}

# Pagination setup
request_page_num = 1
all_activities = []

while True:
    param = {'per_page': 200, 'page': request_page_num}
    response = requests.get(activities_url, headers=header, params=param)

    if response.status_code != 200:
        print(f"Error: {response.status_code}, {response.text}")
        break

    my_dataset = response.json()

    if not my_dataset:
        print("No more activities found, stopping fetch.")
        break

    all_activities.extend(my_dataset)
    request_page_num += 1

print(f"Total activities retrieved: {len(all_activities)}")

# Data storage
activity_data = []
total_kudos = 0
total_active_time = 0  # in seconds

for count, activity in enumerate(all_activities, start=1):
    name = activity.get("name", "Unnamed Activity")
    activity_type = activity.get("type", "Unknown")
    moving_time = activity.get("moving_time", 0)  # in seconds
    kudos_count = activity.get("kudos_count", 0)

    # Convert moving_time from seconds to HH:MM:SS format
    hours, remainder = divmod(moving_time, 3600)
    minutes, seconds = divmod(remainder, 60)
    formatted_time = f"{hours:02}:{minutes:02}:{seconds:02}"

    # Store data in list
    activity_data.append([count, name, activity_type, formatted_time, kudos_count])

    # Totals
    total_kudos += kudos_count
    total_active_time += moving_time

# Convert total active time to HH:MM:SS format
total_hours, remainder = divmod(total_active_time, 3600)
total_minutes, total_seconds = divmod(remainder, 60)
formatted_total_time = f"{total_hours:02}:{total_minutes:02}:{total_seconds:02}"

# Create a DataFrame for activity data
df_activities = pd.DataFrame(activity_data, columns=["Activity Number", "Activity Name", "Activity Type", "Active Time", "Kudos"])

# Create a DataFrame for totals
df_totals = pd.DataFrame(
    [["Total Kudos", total_kudos], ["Total Active Time", formatted_total_time]],
    columns=["DataType", "Values"]
)

# Save to Excel
file_name = "Strava_Activities.xlsx"
with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
    df_activities.to_excel(writer, sheet_name="Activities", index=False)
    df_totals.to_excel(writer, sheet_name="Summary", index=False)

print(f"Data successfully written to {file_name}")
