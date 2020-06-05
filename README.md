# Webscrape to Google Sheets #

## Overview
This project is a python program to web scrape tabular data, reorganize it, and create and push the data onto a google sheets spreadsheet.

Run data.py to webscrape and create and push to google sheets. It also calls the funtion in formatting.py to format the google sheets, like adding a color gradient conditional formatting. eg: https://docs.google.com/spreadsheets/d/1gysR3bUe8qhgDt55AAuy1jf5gIAjPT98mp8DoEmbmK0/edit?usp=sharing

## Objective
Learn and practice webscraping and basic Google API.

## Features
Scrapes the web page at https://api2.sgx.com/sites/default/files/reports/settlement-prices/2020/05/wcm%40sgx_en%40iron_dsp%4018-May-2020%40Iron_Ore_Options_DSP.html and creates a reordered python dictionary, then creates a google spreadsheet, pushes the data, and styles the sheet.

## Instructions
The .json authentication key has been removed for privacy reasons. Please create a Google service account and place the client_secret.json file into the same directory.
1. Go to https://console.developers.google.com/
2. Create a new project.
3. Click Enable API. Search for and enable the Google Drive and Google Sheets API.
4. Create credentials for a Web Server to access Application Data.
5. Name the service account and grant it a Project Role of Editor.
6. Download the JSON file.
7. Copy the JSON file to your code directory and rename it to client_secret.json
8. Edit SHEET_EMAILS in data.py by copying client_email from client_secret.json and adding own email.
