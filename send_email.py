import win32com.client
import datetime

outlook = win32com.client.Dispatch("Outlook.Application")


def getFriday():
    d = datetime.date.today()+timedelta(1)

    while d.weekday() != 4:
        d += datetime.timedelta(1)
    
    meetingTime = str(d)+" 12:00"
    print(meetingTime)

    return meetingTime


def sendMeeting(startdt):    
  appt = outlook.CreateItem(1) # AppointmentItem
  appt.Start = startdt # yyyy-MM-dd hh:mm
  appt.Subject = "MatchWorks Meeting Invitation!"
  appt.Duration = 30 # In minutes (60 Minutes)
  appt.Location = "Follow Up On Location"
  appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added
  
  appt.Recipients.Add("ageannop@mathworks.com") # Don't end ; as delimiter

  appt.Save()
  appt.Send()

def main():
    startdt = getFriday()
    sendMeeting(startdt)
    print("sent meeting invite")
main()