from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from collections import Counter, OrderedDict
from operator import itemgetter
from datetime import datetime

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)

    #Building list to hold senders
    senderEmail = []
    
    #Set Page token to 1 so we begin the API call
    pageToken = '1'
    #This variable is just for counting messages
    count = 1
    #While there are still more messages to retreive, the API will return a nextPageToken attribute. If this is returned, we grab it
    #and send another API call with its value to return more messages. If it doesn't return one, that is all the messages, so end the loop
    while pageToken:
        #Build and execute API call to user.messages.list that returns a list of unread messages in inbox.
        results = service.users().messages().list(userId='me',q="is:unread in:inbox",maxResults="500",pageToken=pageToken).execute()
        #Save messages object in the HTTP response to variable named messages
        messages = results.get('messages', [])
        #For every message returned by the API
        for message in messages:
            #Save the message ID, as we will need to make another API call to see message attributes
            messageId=message['id']
            #Build and execute API call to user.message.get, which will get the message objects metadata (format=metadata)
            #and will only return the "From" field of the metadata
            messageResults = service.users().messages().get(userId='me',id=messageId,format="metadata",metadataHeaders="From").execute()
            #Save payload object (contains message attributes)
            messagePayload = messageResults.get('payload',{})
            #Save "headers" object from the payload (contains our metadata)
            messageHeaders = messagePayload['headers']
            #Saving the 1st entry of the metadata list (which is the From field, as we only are returning one field per API call)
            #We split that field into a list as it contains the name of the sender and the address both in the same string
            messageSender = messageHeaders[0]['value']
            #We append the message sender (From field from headers) to a list
            senderEmail.append(messageSender)
            #Print out # and sender just so we can see the script working :)
            print("Fetched Message #{} from {}".format(count,messageSender))
            #interate the count +1
            count = count + 1
            #if count > 30:
            #    break
        #Here is where we do pagenation
        #If the HTTP response does have the nextPageToken, there is more messages, so we update the pageToken variable with the nextPageToken
        #This will then be placed into the next API call we make to for lists when the loop reiterates
        if results.get('nextPageToken'):
            pageToken = results.get('nextPageToken')
        #If the nextPageToken does not exist, that is all the messages given our API call criteria, so break the loop and stop calling the API
        else:
            break
    
    #Here we take the list of senders and we count duplicates, and output to a dict
    repeatedCount = dict(Counter(senderEmail))
    #Here we sort the dictionary by most frequent senders, and retain the key/value structure
    repeatedCount = OrderedDict(sorted(repeatedCount.items(), key=itemgetter(1), reverse=True))
    #We get the datetime and save it to string variable
    cur_time_dt = datetime.now()
    cur_time_str = cur_time_dt.strftime("%m_%d-%H:%M:%S")
    #We open (or create and open if does not exist) a text file to save the data. We open it as append to not accidently overwrite data
    txt_file = open('frequency_lists/email_list_{}.txt'.format(cur_time_str),'a+')
    #iterate over the key and values. Must called myDict.items() in order to actually print both the name of the key and the value
    for key,value in repeatedCount.items():
        #Writing the data to a txt file
        txt_file.write("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n")
        txt_file.write("{} sent an email {} times\n".format(key,value))
    #Don't forget to close your text file :)
    txt_file.close()
   
if __name__ == '__main__':
    main()