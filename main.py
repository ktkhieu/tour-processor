#!/usr/bin/env python3

import os
import json
import tempfile
from datetime import datetime
from typing import Dict, Optional
from flask import Flask, render_template_string, request, jsonify, session
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

app = Flask(__name__)
app.secret_key = os.urandom(24)  # For session management

class TourProcessor:
    def __init__(self):
        self.service = None
        self.spreadsheet_id = None
        
    def authenticate_with_user_credentials(self, credentials_json: str, spreadsheet_id: str) -> tuple[bool, str]:
        """
        Authenticate with user-provided credentials
        
        Returns:
            (success: bool, message: str)
        """
        try:
            # Parse the JSON credentials
            credentials_data = json.loads(credentials_json)
            
            # Validate it's a service account
            if credentials_data.get('type') != 'service_account':
                return False, "Invalid credentials: Must be a service account JSON file"
            
            # Create credentials from the JSON data
            credentials = ServiceAccountCredentials.from_service_account_info(
                credentials_data, 
                scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            
            # Build the service
            service = build('sheets', 'v4', credentials=credentials)
            
            # Test the connection by trying to read from the spreadsheet
            test_result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range='A1:A1'
            ).execute()
            
            # If we get here, authentication and spreadsheet access worked
            self.service = service
            self.spreadsheet_id = spreadsheet_id
            
            return True, "Successfully connected to Google Sheets"
            
        except json.JSONDecodeError:
            return False, "Invalid JSON format in credentials"
        except HttpError as e:
            if e.resp.status == 403:
                return False, "Permission denied: Make sure you've shared the spreadsheet with the service account email"
            elif e.resp.status == 404:
                return False, "Spreadsheet not found: Check your spreadsheet ID"
            else:
                return False, f"Google Sheets API error: {e.resp.status}"
        except Exception as e:
            return False, f"Authentication failed: {str(e)}"
    
    def extract_data(self, raw_data: str) -> Dict[str, str]:
        """Extract the key fields from tour request data"""
        lines = [line.strip() for line in raw_data.split('\n') if line.strip()]
        data = {}
        
        # Parse basic fields
        for line in lines:
            if ':' in line:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    key, value = parts
                    data[key.strip()] = value.strip()
        
        # Extract Name/Value pairs
        name_value_pairs = {}
        i = 0
        while i < len(lines):
            if lines[i].startswith('Name:') and i + 1 < len(lines) and lines[i + 1].startswith('Value:'):
                name = lines[i].replace('Name:', '').strip()
                value = lines[i + 1].replace('Value:', '').strip()
                name_value_pairs[name] = value
                i += 2
            else:
                i += 1
        
        # Build the final extracted data
        date_requested = name_value_pairs.get('Tour Requested Date', '')
        if not date_requested:
            date1 = name_value_pairs.get('White House Date 1', '')
            date3 = name_value_pairs.get('White House Date 3', '')
            if date1 and date3:
                date_requested = f"{date1} - {date3}"
            elif date1:
                date_requested = date1
        
        # Build comments
        comments_parts = []
        if name_value_pairs.get('Tours Requested'):
            comments_parts.append(f"Tour: {name_value_pairs['Tours Requested']}")
        if name_value_pairs.get('Name of Visitors'):
            comments_parts.append(f"Visitors: {name_value_pairs['Name of Visitors']}")
        if name_value_pairs.get('Channel'):
            comments_parts.append(f"Channel: {name_value_pairs['Channel']}")
        
        # Add address
        address_parts = []
        for field in ['Address1', 'Address2', 'City', 'State', 'ZipCode']:
            if data.get(field):
                address_parts.append(data[field])
        if address_parts:
            comments_parts.append(f"Address: {', '.join(address_parts)}")
        
        full_name = f"{data.get('FirstName', '')} {data.get('LastName', '')}".strip()
        
        return {
            'name': full_name,
            'email': data.get('EmailAddress', ''),
            'phone': data.get('PhoneNumber', ''),
            'dates_requested': date_requested,
            'party_of': name_value_pairs.get('Total Number of People in Party', ''),
            'rsvp_link': '',
            'unable': '',
            'comments': '; '.join(comments_parts)
        }
    
    def add_to_sheet(self, extracted_data: Dict[str, str]) -> bool:
        """Add the extracted data to Google Sheets"""
        if not self.service or not self.spreadsheet_id:
            return False
        
        # Check if sheet has headers, add them if not
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range='A1:H1'
            ).execute()
            
            values = result.get('values', [])
            if not values:
                # Add headers
                headers = [[
                    'Name', 'Email', 'Phone Number', 'Dates Requested', 
                    'Party Of', 'RSVP Link', 'Unable', 'Comments'
                ]]
                
                self.service.spreadsheets().values().append(
                    spreadsheetId=self.spreadsheet_id,
                    range='A1:H1',
                    valueInputOption='RAW',
                    body={'values': headers}
                ).execute()
        except:
            pass  # If we can't check/add headers, continue anyway
        
        # Add the data row
        row = [[
            extracted_data['name'],
            extracted_data['email'],
            extracted_data['phone'],
            extracted_data['dates_requested'],
            extracted_data['party_of'],
            extracted_data['rsvp_link'],
            extracted_data['unable'],
            extracted_data['comments']
        ]]
        
        try:
            self.service.spreadsheets().values().append(
                spreadsheetId=self.spreadsheet_id,
                range='A:H',
                valueInputOption='RAW',
                body={'values': row}
            ).execute()
            return True
        except HttpError:
            return False

# Initialize processor
processor = TourProcessor()

# HTML Template
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tour Request Processor</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: #f8fafc;
            color: #334155;
            line-height: 1.6;
        }

        .container {
            max-width: 800px;
            margin: 40px auto;
            padding: 0 20px;
        }

        .header {
            text-align: center;
            margin-bottom: 40px;
        }

        .header h1 {
            font-size: 28px;
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 8px;
        }

        .header p {
            color: #64748b;
            font-size: 16px;
        }

        .card {
            background: white;
            border-radius: 12px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
            border: 1px solid #e2e8f0;
            padding: 24px;
            margin-bottom: 24px;
        }

        .setup-card {
            border-left: 4px solid #3b82f6;
        }

        .form-group {
            margin-bottom: 20px;
        }

        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 500;
            color: #374151;
        }

        input, textarea {
            width: 100%;
            padding: 12px;
            border: 1px solid #d1d5db;
            border-radius: 8px;
            font-size: 14px;
            transition: border-color 0.2s ease;
        }

        input:focus, textarea:focus {
            outline: none;
            border-color: #3b82f6;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
        }

        textarea {
            font-family: 'SF Mono', 'Monaco', 'Inconsolata', monospace;
            background: #fafafa;
            resize: vertical;
        }

        .credentials-textarea {
            min-height: 150px;
            font-size: 12px;
        }

        .data-textarea {
            min-height: 200px;
        }

        .button-group {
            display: flex;
            gap: 12px;
            margin-top: 20px;
        }

        button {
            padding: 12px 20px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .btn-primary {
            background: #3b82f6;
            color: white;
            flex: 1;
        }

        .btn-primary:hover:not(:disabled) {
            background: #2563eb;
        }

        .btn-success {
            background: #10b981;
            color: white;
            flex: 1;
        }

        .btn-success:hover:not(:disabled) {
            background: #059669;
        }

        .btn-secondary {
            background: white;
            color: #6b7280;
            border: 1px solid #d1d5db;
            flex: 1;
        }

        .btn-secondary:hover {
            background: #f9fafb;
        }

        .result {
            display: none;
            padding: 16px;
            border-radius: 8px;
            margin-top: 20px;
        }

        .result.success {
            background: #f0fdf4;
            border: 1px solid #bbf7d0;
            color: #166534;
        }

        .result.error {
            background: #fef2f2;
            border: 1px solid #fecaca;
            color: #dc2626;
        }

        .extracted-data {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            overflow: hidden;
        }

        .data-row {
            display: flex;
            padding: 12px 16px;
            border-bottom: 1px solid #f3f4f6;
        }

        .data-row:last-child {
            border-bottom: none;
        }

        .data-row:nth-child(even) {
            background: #f9fafb;
        }

        .data-label {
            font-weight: 500;
            width: 140px;
            color: #374151;
            flex-shrink: 0;
        }

        .data-value {
            color: #6b7280;
            word-break: break-word;
        }

        .action-buttons {
            display: flex;
            gap: 12px;
            margin-top: 16px;
        }

        .loading {
            display: none;
            text-align: center;
            padding: 20px;
            color: #6b7280;
        }

        .spinner {
            border: 2px solid #f3f4f6;
            border-top: 2px solid #3b82f6;
            border-radius: 50%;
            width: 20px;
            height: 20px;
            animation: spin 1s linear infinite;
            margin: 0 auto 8px;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .status {
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 500;
            display: inline-block;
            margin-bottom: 16px;
        }

        .status.connected {
            background: #f0fdf4;
            color: #166534;
            border: 1px solid #bbf7d0;
        }

        .status.disconnected {
            background: #fff7ed;
            color: #9a3412;
            border: 1px solid #fed7aa;
        }

        .instructions {
            background: #fffbeb;
            border: 1px solid #fed7aa;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 20px;
        }

        .instructions h4 {
            color: #92400e;
            margin-bottom: 8px;
            font-size: 14px;
        }

        .instructions ol {
            color: #b45309;
            font-size: 13px;
            padding-left: 20px;
        }

        .instructions li {
            margin-bottom: 4px;
        }

        .instructions a {
            color: #2563eb;
        }

        button:disabled {
            opacity: 0.6;
            cursor: not-allowed;
        }

        .step-title {
            font-size: 18px;
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 12px;
        }

        @media (max-width: 768px) {
            .container {
                padding: 0 16px;
            }
            
            .button-group {
                flex-direction: column;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Tour Request Processor</h1>
            <p>Connect your Google Sheets and process tour request data</p>
        </div>

        <!-- Step 1: Google Sheets Setup -->
        <div class="card setup-card" id="setupCard">
            <div class="step-title">Connect Your Google Sheets</div>
            
            <div class="instructions">
                <h4>Setup Instructions:</h4>
                <ol>
                    <li>Go to <a href="https://console.cloud.google.com/" target="_blank">Google Cloud Console</a></li>
                    <li>Create a project and enable Google Sheets API</li>
                    <li>Create a Service Account and download the JSON key</li>
                    <li>Share your Google Sheet with the service account email</li>
                    <li>Paste the JSON content and spreadsheet ID below</li>
                </ol>
            </div>

            <form id="setupForm">
                <div class="form-group">
                    <label for="credentials">Service Account JSON Credentials</label>
                    <textarea 
                        id="credentials" 
                        class="credentials-textarea"
                        placeholder='Paste your service account JSON here, e.g.:
{
  "type": "service_account",
  "project_id": "your-project",
  "private_key_id": "...",
  "private_key": "...",
  "client_email": "...",
  ...
}'
                        required>
                    </textarea>
                </div>

                <div class="form-group">
                    <label for="spreadsheetId">Google Spreadsheet ID</label>
                    <input 
                        type="text" 
                        id="spreadsheetId" 
                        placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
                        required>
                    <small style="color: #6b7280; font-size: 12px;">
                        Get this from your Google Sheets URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
                    </small>
                </div>

                <div class="button-group">
                    <button type="submit" class="btn-primary" id="connectBtn">
                        Connect to Google Sheets
                    </button>
                </div>
            </form>

            <div class="loading" id="setupLoading">
                <div class="spinner"></div>
                <p>Connecting to Google Sheets...</p>
            </div>

            <div class="result" id="setupResult"></div>
        </div>

        <!-- Step 2: Process Tour Data -->
        <div class="card" id="dataCard" style="display: none;">
            <div class="step-title">
                Tour Request Processor
                <button onclick="logout()" class="btn-secondary" style="float: right; padding: 8px 16px; font-size: 12px;">
                    Disconnect
                </button>
            </div>
            <div class="status connected" id="connectionStatus">
                ✓ Connected to Google Sheets
            </div>

            <form id="tourForm">
                <div class="form-group">
                    <label for="rawData">Raw Tour Request Data</label>
                    <textarea 
                        id="rawData" 
                        class="data-textarea"
                        placeholder="Paste your tour request data here..."
                        required>
                    </textarea>
                </div>

                <div class="button-group">
                    <button type="submit" class="btn-success">
                        Process & Save to Google Sheets
                    </button>
                    <button type="button" class="btn-secondary" onclick="clearForm()">
                        Clear
                    </button>
                </div>
            </form>

            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p>Processing...</p>
            </div>

            <div class="result" id="result"></div>
        </div>
    </div>

    <script>
        // Setup form handler
        document.getElementById('setupForm').addEventListener('submit', function(e) {
            e.preventDefault();
            connectToSheets();
        });

        // Tour form handler
        document.getElementById('tourForm').addEventListener('submit', function(e) {
            e.preventDefault();
            processTourData();
        });

        function connectToSheets() {
            const credentials = document.getElementById('credentials').value.trim();
            const spreadsheetId = document.getElementById('spreadsheetId').value.trim();
            const connectBtn = document.getElementById('connectBtn');
            
            if (!credentials || !spreadsheetId) {
                showSetupResult('Please fill in both credentials and spreadsheet ID.', 'error');
                return;
            }

            // Disable button and show loading
            connectBtn.disabled = true;
            connectBtn.textContent = 'Connecting...';
            document.getElementById('setupLoading').style.display = 'block';
            document.getElementById('setupResult').style.display = 'none';

            fetch('/connect', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    credentials: credentials,
                    spreadsheetId: spreadsheetId
                })
            })
            .then(response => response.json())
            .then(data => {
                document.getElementById('setupLoading').style.display = 'none';
                connectBtn.disabled = false;
                connectBtn.textContent = 'Connect to Google Sheets';
                
                if (data.success) {
                    // Hide the setup card and show the data processing card
                    document.getElementById('setupCard').style.display = 'none';
                    document.getElementById('dataCard').style.display = 'block';
                    
                    // Update header to show logged in state
                    document.querySelector('.header h1').textContent = 'Tour Request Processor';
                    document.querySelector('.header p').textContent = 'Process tour request data and save to your Google Sheets';
                    
                    // Focus on the textarea
                    document.getElementById('rawData').focus();
                } else {
                    showSetupResult('Connection failed: ' + data.error, 'error');
                }
            })
            .catch(error => {
                document.getElementById('setupLoading').style.display = 'none';
                connectBtn.disabled = false;
                connectBtn.textContent = 'Connect to Google Sheets';
                showSetupResult('Error connecting: ' + error.message, 'error');
            });
        }

        function logout() {
            // Show confirmation
            if (confirm('Are you sure you want to disconnect? You will need to re-enter your credentials.')) {
                // Reset the interface
                document.getElementById('setupCard').style.display = 'block';
                document.getElementById('dataCard').style.display = 'none';
                
                // Clear forms
                document.getElementById('credentials').value = '';
                document.getElementById('spreadsheetId').value = '';
                document.getElementById('rawData').value = '';
                document.getElementById('result').style.display = 'none';
                document.getElementById('setupResult').style.display = 'none';
                
                // Reset header
                document.querySelector('.header h1').textContent = 'Tour Request Processor';
                document.querySelector('.header p').textContent = 'Connect your Google Sheets and process tour request data';
                
                // Call server to clear session
                fetch('/logout', {
                    method: 'POST'
                });
            }
        }

        function processTourData() {
            const rawData = document.getElementById('rawData').value.trim();
            
            if (!rawData) {
                showResult('Please paste some tour request data first.', 'error');
                return;
            }

            document.getElementById('loading').style.display = 'block';
            document.getElementById('result').style.display = 'none';

            fetch('/process_and_save', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({rawData: rawData})
            })
            .then(response => response.json())
            .then(data => {
                document.getElementById('loading').style.display = 'none';
                if (data.success) {
                    displayExtractedData(data.extractedData, data.saved);
                } else {
                    showResult('Error: ' + data.error, 'error');
                }
            })
            .catch(error => {
                document.getElementById('loading').style.display = 'none';
                showResult('Error processing data: ' + error.message, 'error');
            });
        }

        function displayExtractedData(data, saved) {
            const resultDiv = document.getElementById('result');
            
            const statusMessage = saved ? 
                '<div style="background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; padding: 12px; border-radius: 6px; margin-bottom: 16px; font-weight: 500;">✅ Successfully saved to Google Sheets!</div>' :
                '<div style="background: #fef2f2; border: 1px solid #fecaca; color: #dc2626; padding: 12px; border-radius: 6px; margin-bottom: 16px; font-weight: 500;">⚠️ Processed but not saved to Google Sheets</div>';
            
            const html = `
                ${statusMessage}
                <div class="extracted-data">
                    <div class="data-row">
                        <div class="data-label">Name</div>
                        <div class="data-value">${data.name}</div>
                    </div>
                    <div class="data-row">
                        <div class="data-label">Email</div>
                        <div class="data-value">${data.email}</div>
                    </div>
                    <div class="data-row">
                        <div class="data-label">Phone Number</div>
                        <div class="data-value">${data.phone}</div>
                    </div>
                    <div class="data-row">
                        <div class="data-label">Dates Requested</div>
                        <div class="data-value">${data.dates_requested}</div>
                    </div>
                    <div class="data-row">
                        <div class="data-label">Party Of</div>
                        <div class="data-value">${data.party_of}</div>
                    </div>
                    <div class="data-row">
                        <div class="data-label">RSVP Link</div>
                        <div class="data-value">(empty)</div>
                    </div>
                    <div class="data-row">
                        <div class="data-label">Unable</div>
                        <div class="data-value">(empty)</div>
                    </div>
                    <div class="data-row">
                        <div class="data-label">Comments</div>
                        <div class="data-value">${data.comments}</div>
                    </div>
                </div>
                <div class="action-buttons">
                    <button onclick="copyToClipboard()" class="btn-secondary">
                        Copy to Clipboard
                    </button>
                    <button onclick="clearForm()" class="btn-primary">
                        Process Another Request
                    </button>
                </div>
            `;

            resultDiv.innerHTML = html;
            resultDiv.className = 'result success';
            resultDiv.style.display = 'block';

            window.currentExtractedData = data;
        }

        function copyToClipboard() {
            const data = window.currentExtractedData;
            if (!data) return;

            const row = [
                data.name,
                data.email,
                data.phone,
                data.dates_requested,
                data.party_of,
                data.rsvp_link,
                data.unable,
                data.comments
            ].join('\\t');

            navigator.clipboard.writeText(row).then(() => {
                showResult('✓ Copied to clipboard! Paste into your spreadsheet.', 'success');
            });
        }

        function saveToSheets() {
            // This function is no longer needed since we auto-save
            // Keeping it for the copy to clipboard functionality
        }

        function showSetupResult(message, type) {
            const resultDiv = document.getElementById('setupResult');
            resultDiv.innerHTML = `<p>${message}</p>`;
            resultDiv.className = `result ${type}`;
            resultDiv.style.display = 'block';
        }

        function showResult(message, type) {
            const resultDiv = document.getElementById('result');
            resultDiv.innerHTML = `<p>${message}</p>`;
            resultDiv.className = `result ${type}`;
            resultDiv.style.display = 'block';
        }

        function clearForm() {
            document.getElementById('rawData').value = '';
            document.getElementById('result').style.display = 'none';
            window.currentExtractedData = null;
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    """Main page"""
    return render_template_string(HTML_TEMPLATE)

@app.route('/logout', methods=['POST'])
def logout():
    """Clear user session"""
    session.clear()
    return jsonify({'success': True})

@app.route('/connect', methods=['POST'])
def connect():
    """Connect to user's Google Sheets"""
    try:
        data = request.json
        credentials_json = data.get('credentials', '')
        spreadsheet_id = data.get('spreadsheetId', '')
        
        if not credentials_json or not spreadsheet_id:
            return jsonify({'success': False, 'error': 'Missing credentials or spreadsheet ID'})
        
        success, message = processor.authenticate_with_user_credentials(credentials_json, spreadsheet_id)
        
        if success:
            # Store connection status in session
            session['connected'] = True
            return jsonify({'success': True, 'message': message})
        else:
            return jsonify({'success': False, 'error': message})
            
    except Exception as e:
        return jsonify({'success': False, 'error': f'Unexpected error: {str(e)}'})

@app.route('/process_and_save', methods=['POST'])
def process_and_save():
    """Process tour request data and automatically save to Google Sheets"""
    try:
        data = request.json
        raw_data = data.get('rawData', '')
        
        if not raw_data:
            return jsonify({'success': False, 'error': 'No data provided'})
        
        if not session.get('connected'):
            return jsonify({'success': False, 'error': 'Not connected to Google Sheets'})
        
        # Extract the data
        extracted_data = processor.extract_data(raw_data)
        
        # Try to save to Google Sheets
        saved = processor.add_to_sheet(extracted_data)
        
        return jsonify({
            'success': True,
            'extractedData': extracted_data,
            'saved': saved
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/process', methods=['POST'])
def process():
    """Process tour request data (legacy endpoint)"""
    try:
        data = request.json
        raw_data = data.get('rawData', '')
        
        if not raw_data:
            return jsonify({'success': False, 'error': 'No data provided'})
        
        extracted_data = processor.extract_data(raw_data)
        
        return jsonify({
            'success': True,
            'extractedData': extracted_data
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/save', methods=['POST'])
def save():
    """Save to Google Sheets"""
    try:
        data = request.json
        
        if not session.get('connected'):
            return jsonify({'success': False, 'error': 'Not connected to Google Sheets'})
        
        if processor.add_to_sheet(data):
            return jsonify({'success': True})
        else:
            return jsonify({'success': False, 'error': 'Failed to save to Google Sheets'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

if __name__ == '__main__':
    print("🚀 Starting Multi-User Tour Request Processor")
    print("📍 Open your browser to: http://localhost:5000")
    print("💡 Users can connect their own Google Sheets!")
    
    app.run(debug=True, host='0.0.0.0', port=5000)