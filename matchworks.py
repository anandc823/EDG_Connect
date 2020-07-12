import pandas as pd
import requests
import win32com.client
import datetime

outlook = win32com.client.Dispatch("Outlook.Application")
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
        if len(members)==6:
            print("6")   # Need to fill in this part to just create two groups of three
        else:
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
                    size_counter+=1
                    if(size_counter>=max_meeting_size):
                        cur_meeting['members'].append(member)
                        meetings.append(cur_meeting.copy())
                        print("built 4")
                        print(meetings)
                        size_counter = 0
                        cur_meeting['members'] = []
                        print(meetings)
                    else:
                        cur_meeting['members'].append(member)
                        print('building block of 4',cur_meeting)
                if len(cur_meeting['members'])==1 or len(cur_meeting['members'])==2:
                    index = -1
                    for extraMember in cur_meeting['members']:
                        meetings[index]['members'].append(extraMember)
                        index-=1
                else:
                    meetings.append(cur_meeting)

    return meetings

def getFriday():
    d = datetime.date.today()+datetime.timedelta(1)

    while d.weekday() != 4:
        d += datetime.timedelta(1)
    
    meetingTime = str(d)+" 12:00"
    print(meetingTime)

    return meetingTime


def sendMeeting(startdt,topic,recipients):    
  appt = outlook.CreateItem(1) # AppointmentItem
  appt.Start = startdt # yyyy-MM-dd hh:mm
  appt.Subject = f'MatchWorks Meeting Invitation: {topic}!'
  appt.Duration = 30 # In minutes (60 Minutes)
  appt.Location = "Follow Up On Location"
  appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added
  
  for email in recipients:
    appt.Recipients.Add(email) # Don't end ; as delimiter

  appt.Save()
  appt.Send()

def main():
    download_data()
    meetings = create_meetings()
    meetingdt = getFriday()

    for meeting in meetings:
        sendMeeting(meetingdt,meeting['topic'],meeting['members'])
main()