import pandas as pd
import requests

filename = "survey_answers.csv"
max_meeting_size=  4
min_meeting_size = 2
response_prompt = 'What type of meeting would you be interested in having with other EDG members?'
email_prompt = "What's your MathWorks email?"

def download_data():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQzboGzFB0L-ogAsEArqPOPABbeX2BdXffkmLEAYYMSTQFkID_kPK26JTXtRhszM6OASX2-cEg-fOsa/pub?output=csv"
    
    r = requests.get(url)

    with open("survey_answers.csv",'wb') as f:
        f.write(r.content)

def create_meetings():
    df = pd.read_csv(filename)

    topics = df[response_prompt].unique()

    groups = {}

    for topic in topics:
        groups[topic] = list(df[df[response_prompt]==topic][email_prompt])

    meetings = []
    
    for topic in topics:
        members = groups[topic]

        if(len(members)) == 1:
            print("Only 1 Member Found")

        if(len(members)<=max_meeting_size):
            cur_meeting = {}
            cur_meeting['topic'] = topic
            cur_meeting['members'] = members
            meetings.append(cur_meeting)
        
        else:   
            cur_meeting = {}
            cur_meeting['topic'] = topic
            cur_meeting['members'] = []
            size_counter = 0
            meeting_counter = 0
            overflow_index = 0
            for member in members:
                if meeting_counter>=(len(members)//max_meeting_size):
                    meetings[overflow_index][members].append(member)
                    overflow_index+=1
                else:
                    size_counter+=1
                    if(size_counter>=max_meeting_size):
                        metings.append(cur_meeting)
                        size_counter = 0
                
    return meetings


def main():
    download_data()
    create_meetings()
main()