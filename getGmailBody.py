
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import email
import base64
import csv

#don't forget to add the credentials file in your project

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

#global variables where all the info which will be saved in the SCV are
SAVE_TO_INFO_SCV = [["Date and Time", "Message Id", "Useful_info1", "Useful_info2", "Useful_info3"]]


def main():
    #connect to the gmail account
    service = getService()

    getFundingInfo(service)
    getRefundingInfo(service)
    
def getFundingInfo(service):
    #search for label
    searchString = "from:(@trading212.com) subject:(Real Account Funded Successfully)"
    #get the id list of the emails that match as well as their number
    list_of_funding_ids, numberResults = search_messages(service, 'me', searchString)
    #print the total matching emails for the user
    print("Number of Trading 212 funding transactions: " + str(numberResults) + "\n")
    #specify where the script will find the usefull information
    start_str1 = "start of useful info 1"
    end_str1 = "end of useful info 1"
    start_str2 = "start of useful info 2"
    end_str2 = "end of useful info 2"
    start_str3 = "start of useful info 3"
    end_str3 = "end of useful info 3"

    #Take each email (only applicable if there is at least one email)
    if numberResults > 0:
        #loop for each email
        for i in range (numberResults):
            get_message(service, 'me', list_of_funding_ids[i], start_str1, end_str1, start_str2, end_str2, start_str3, end_str3)
        #save everything in an SCV
        save_info_to_CSV()


def getService():
    
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    return service


def search_messages(service, user_id, searchString):
    try:
        search_id = service.users().messages().list(userId=user_id, q=searchString).execute()

        #find the number of results that match
        numberResults = search_id['resultSizeEstimate']

        #get the ids of the result
        if numberResults > 0:

            #create a list to put all the ids
            final_list = []

            #get all the ids
            message_ids = search_id['messages']
            
            for ids in message_ids:                
                #add each id in the list
                final_list.append(ids["id"])
                
            #return the list
            return final_list, numberResults
        else:
            
            print("There were no results for that search. Returning an empty string")
            #if there are no messages return an empty string
            return "", numberResults
    
    except Exception as error:
        print('An error occurred in the search_messages function: %s' % error)


def get_message(service, user_id, msg_id, start_str1, end_str1, start_str2, end_str2, start_str3, end_str3):
    try:
        #get the message in raw format
        message = service.users().messages().get(userId=user_id, id=msg_id, format='raw').execute()
        
        #encode the message in ASCII format
        msg_str = base64.urlsafe_b64decode(message['raw'].encode("utf-8")).decode("utf-8")
        #put it in an email object
        mime_msg = email.message_from_string(msg_str)
        #get the date and time of the email
        date = mime_msg['date']
        
        
        #get the plain and html text from the payload
        for part in mime_msg.walk():
            
        #each part is a either non-multipart, or another multipart message
        #that contains further parts... Message is organized like a tree
            if part.get_content_type() == 'text/plain':
                #prints the raw text that we want
                text = part.get_payload()              
                print_formatted_text(msg_id, text, date, start_str1, end_str1, start_str2, end_str2, start_str3, end_str3)
    
    except Exception as error:
        print('An error occurred in the get_message function: %s' % error)




def print_formatted_text(msg_id, text, date, start_str1, end_str1, start_str2, end_str2, start_str3, end_str3):
    #get all the common info
    useful_info1 = split_text(text, start_str1, end_str1)
    #print all the common info
    print("Date: " + date)
    print("Email id: " + msg_id)

    useful_info2 = split_text(text, start_str2, end_str2)
    useful_info3 = split_text(text, start_str3, end_str3)
    #print the remaining info for funding
    print(start_str1 + " " + deposited_withdrawn)
    print(start_str2 + " " + transactionFee)
    print(start_str3 + " " + total  + "\n")
    #refresh the global variable SAVE_TO_FUNDING_SCV
    SAVE_TO_INFO_SCV.append([date, msg_id, useful_info1, useful_info2, useful_info3])
    

def save_info_to_CSV():
    with open('Trading212Funding.csv', 'w', newline='') as file:
        writer = csv.writer(file, delimiter='|')
        writer.writerows(SAVE_TO_INFO_SCV)



def split_text(plainText, start, end):
    usefullText = plainText.split(start)[1]
    usefullText = usefullText.split(end)[0]
    return usefullText.strip()
    

if __name__ == '__main__':
    main()
