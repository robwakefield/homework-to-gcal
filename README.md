# Homework to Google Calendar script

Python script which retrieves current homework assignments from school website and adds them to Google Calendar.

This script uses the python library Selenium (which is used for web-testing) in order to scrape the School's website as elements only become visible after the user clicks and interacts with them. Therefore a process of simulating these clicks is used in place of Requests and BeautifulSoup for example.

client_secret.json, credentials.json, storage.json can all be found by using the Google Developer Console and creating a new app.
Updated version of Chromedriver.exe is needed in order to run script.
