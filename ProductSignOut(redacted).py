import datetime
from flask import Flask, render_template, request, redirect, url_for
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Google Sheets API setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = ''

credentials = service_account.Credentials.from_service_account_file('filepath.json', scopes=SCOPES)
sheetsInstance = build('sheets', 'v4', credentials=credentials)

# gathers available devices for sign out sheet
def get_available_devices():
    result = sheetsInstance.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{'ASSET: Training Laptop Database'}!B2:C").execute()
    rows = result.get('values', [])
    devices = []
    
    for row in rows:
        if row[0] == 'Available':
            devices.append(row[1])
    return devices

# groups each device by status
def get_devices_by_status():
    # get 'ASSET: Training Laptop Database' sheet and analyze data from column B to F
    result = sheetsInstance.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{'ASSET: Training Laptop Database'}!B2:C").execute()
    rows = result.get('values', [])
    
    availableDevices = []
    checkedOutDevices = []
    serviceDevices = [] 

    for row in rows:
        status = row[0] # element in column B (STATUS)
        deviceId = row[1] # element in column C (DEVICE ID)

        # categorized id by status
        if status == 'Available':
            availableDevices.append(deviceId)
        elif status == 'Checked Out':
            checkedOutDevices.append(deviceId)
        elif status == 'Service':
            serviceDevices.append(deviceId)
    
    return availableDevices, checkedOutDevices, serviceDevices

def sign_out_device(deviceId, assignedTo, notes):
    result = sheetsInstance.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{'ASSET: Training Laptop Database'}!B2:F").execute()
    rows = result.get('values', [])
    
    # for each row update each column from B to F (column B: status, column D, Assigned To, column E: Notes, column F: Updated Timestamp)
    for idx, row in enumerate(rows):
        if row[1] == deviceId:
            updateRange = f"{'ASSET: Training Laptop Database'}!B{idx+2}:F{idx+2}"
            updateValues = [['Checked Out', deviceId, assignedTo, notes, datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')]]
            sheetsInstance.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=updateRange, valueInputOption='RAW', body={'values': updateValues}).execute()
            break

def sign_in_device(deviceId):
    result = sheetsInstance.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{'ASSET: Training Laptop Database'}!B2:F").execute()
    rows = result.get('values', [])
    
    for idx, row in enumerate(rows):
        if row[1] == deviceId:
            updateRange = f"{'ASSET: Training Laptop Database'}!B{idx+2}:F{idx+2}"
            updateValues = [['Available', deviceId, '', '', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')]]
            sheetsInstance.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=updateRange, valueInputOption='RAW', body={'values': updateValues}).execute()
            break

# Flask app setup
app = Flask(__name__)

@app.route('/')
def home():
    availableDevices, checkedOutDevices, serviceDevices = get_devices_by_status()
    return render_template('index.html', available_devices=availableDevices, checked_out_devices=checkedOutDevices, service_devices=serviceDevices)

@app.route('/sign_out', methods=['GET', 'POST'])
def sign_out():
    if request.method == 'POST':
        device_id = request.form['device_id']
        assigned_to = request.form['assigned_to']
        notes = request.form['notes']
        sign_out_device(device_id, assigned_to, notes)
        return redirect(url_for('home'))
    else:
        available_devices = get_available_devices()
        return render_template('signOut.html', devices=available_devices)

@app.route('/sign_in', methods=['GET', 'POST'])
def sign_in():
    if request.method == 'POST':
        device_id = request.form['device_id']
        sign_in_device(device_id)
        return redirect(url_for('home'))
    else:
        return render_template('signIn.html')

if __name__ == '__main__':
    app.run(debug=True)