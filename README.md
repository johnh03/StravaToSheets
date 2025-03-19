
---
# StravaToSheets

A Python program that uses the Strava API to fetch a user's activity data and export it to a local Excel sheet. The spreadsheet includes activity type, time spent, and kudos received.

## ⚡ Setup Instructions  

Before running the program, you must set up the Strava API. Using **Postman** is recommended, but you can also refer to the [Strava Developers Portal](https://developers.strava.com) for a more detailed guide.

### 🔹 1. Create a Strava API Application  

1. Go to [Strava Developers Portal](https://developers.strava.com) and click **"Create & Manage Your App"**  
2. Sign into Strava and create an API application  
3. Take note of the following credentials:
   - **Client ID** – Your application ID  
   - **Client Secret** – Your client secret (**keep this confidential**)  
   - **Authorization Token** – Expires every six hours (**keep this confidential**)  
   - **Refresh Token** – Used to generate a new authorization token (**keep this confidential**)  
   - **Rate Limits** – Your application's current API request limits  
   - **Authorization Callback Domain** – Set this to `localhost` during development  

---

### 🔹 2. Obtain an Authorization Code  

1. Open the following URL in a web browser, replacing `CLIENT_ID` with your actual client ID:  

   ```
   https://www.strava.com/oauth/authorize?client_id=CLIENT_ID&redirect_uri=http://localhost&response_type=code&scope=activity:read_all
   ```

2. The page will redirect to `localhost` (which won’t load).  
3. Copy the **authorization code** from the URL and save it:  

   ```
   Authorization Code: YOUR_CODE
   ```

---

### 🔹 3. Exchange the Authorization Code for Tokens  

1. Use **Postman** (or any API tool) to send a `POST` request to:  

   ```
   https://www.strava.com/oauth/token?client_id=CLIENT_ID&client_secret=CLIENT_SECRET&code=YOUR_CODE&grant_type=authorization_code
   ```

2. Save the returned **refresh token** and **access token**:  

   ```
   Refresh Token: REFRESH_TOKEN
   Access Token: ACCESS_TOKEN
   ```

---

### 🔹 4. Fetch Activities Using the Access Token  

Use **Postman** to send a `GET` request:  

```
https://www.strava.com/api/v3/athlete/activities?access_token=ACCESS_TOKEN
```

This will return a **JSON response** containing all of the user's activities.

---

### 🔹 5. Refresh the Access Token  

Since the access token expires, use the refresh token to generate a new one.  
Send a `POST` request to:  

```
https://www.strava.com/oauth/token?client_id=CLIENT_ID&client_secret=CLIENT_SECRET&refresh_token=REFRESH_TOKEN&grant_type=refresh_token
```

- This will return a **new access token** and inform you of its expiration time.  
- The **refresh token does not expire**, so you can keep using it to get new access tokens.  

---

## 🚀 Running the Program  

Once setup is complete:  

1. Open `StravaToSheets.py`  
2. Replace `CLIENT_ID`, `CLIENT_SECRET`, and `REFRESH_TOKEN` with your credentials  
3. Run the script  

The program will generate an **Excel spreadsheet** with:  
✔ **Activity list** (most recent to least)  
✔ **Activity name, type, and duration**  
✔ **Kudos received**  
✔ **Summary page** with total kudos and total time spent  

---

### 🎯 Notes  

- The setup process is **one-time only**; the script will handle token refreshing automatically.  
- Keep your **Client Secret, Authorization Token, and Refresh Token confidential** to prevent unauthorized access.  

---