import pandas as pd
filename = "survey_answers.xlsx"
max_meeting_size=  4
min_meeting_size = 2

def create_meetings():
    df = pd.read_excel(filename)

    topics = df['response'].unique()

    groups = {}

    for topic in topics:
        groups[topic] = list(df[df.response==topic]["Email"])

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
    meetings = create_meetings()
    print(meetings)
main()