import requests
from datetime import datetime, timedelta
from textblob import TextBlob
import os

# Set up Microsoft Graph API credentials
microsoft_client_id = os.environ.get('CLIENT_ID')
microsoft_client_secret = os.environ.get('CLIENT_SECRET')
microsoft_tenant_id = os.environ.get('TENANT_ID')

# Set up Cisco Email Security credentials
cisco_client_id = os.environ.get('CISCO_CLIENT_ID')
cisco_client_secret = os.environ.get('CISCO_CLIENT_SECRET')

# Get yesterday's date
yesterday = datetime.now() - timedelta(days=1)
yesterday_str = yesterday.strftime('%Y-%m-%dT%H:%M:%SZ')

# Construct the request URL
microsoft_graph_url = f'https://graph.microsoft.com/v1.0/me/messages?$filter=receivedDateTime ge {yesterday_str}'

# Make the request to Microsoft Graph API
response = requests.get(microsoft_graph_url, headers={'Authorization': 'Bearer YOUR_ACCESS_TOKEN'})

# Check if the request was successful
if response.status_code == 200:
    messages = response.json().get('value', [])
    
    # Analyze sentiment for each message
    for message in messages:
        # Check if the email is incoming or outgoing
        if message.get('isReadReceiptRequested') or message.get('isDeliveryReceiptRequested'):
            continue  # Ignore outgoing emails and move to the next one
        
        # Check if the email was detected as a threat by Cisco Email Threat Defense
        if message.get('ciscoThreatDetection'):
            continue  # Ignore emails detected as threats and move to the next one
        
        # Extract the email body
        body = message.get('body', {}).get('content', '')
        
        # Perform sentiment analysis
        blob = TextBlob(body)
        sentiment = blob.sentiment.polarity
        
        # Print the sentiment result
        print(f'Sentiment for email with subject "{message.get("subject")}": {sentiment}')
else:
    print('Failed to retrieve email messages from Microsoft Graph API.')