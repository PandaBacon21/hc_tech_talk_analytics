import pandas as pd 
from datetime import datetime, timedelta


def get_dfs(file_path): 
    dfs = pd.read_excel(file_path, sheet_name=None)

    invite_list_df = dfs['Invite_List']
    event_attendance_df = dfs['Event_Attendance_Details']
    general_ticket_reg_df = dfs['General_Ticket_Registration']
    poll_results_df = dfs['Poll_Results']
    session_summary_df = dfs['Session_Summary']
    #resource_link_engagement = dfs['Resource_Link_Engagement']
    #event_source_tracking = dfs['Event_Source_Tracking']
    #survey_df = dfs['Survey']
    
    return invite_list_df, event_attendance_df, general_ticket_reg_df, poll_results_df, session_summary_df


# GET SESSION NAME, ID, DATE
def session_details(session_summary_df):
    session_name = session_summary_df['Name'].iloc[0]
    session_id = session_summary_df['Session ID'].iloc[0]
    session_date = session_summary_df['Start time'].iloc[0]
    print(f'session name: {session_name}')
    print(f'session id: {session_id}')
    print(f'session date: {session_date}')
    return session_date

# GET TOTAL NUMBER OF INVITES, FIRST INVITE SENT (BY WHICH SE), WHEN FIRST INVITE WAS SENT
def invite_details(invite_list_df, session_date):
    total_invites = invite_list_df['Email Sent'].count()
    print(f'total invites: {total_invites}')
    
    # GET FIRST INVITE SENT AND BY WHICH SE
    invite_list_by_date = invite_list_df.sort_values('Email Sent')
    first_invite_sent = invite_list_by_date.iloc[0]
    first_invite_sent_date = first_invite_sent['Email Sent'].strftime('%Y-%m-%d')
    print(f"first SE to send invite: {first_invite_sent['SE']}")
    print(f"first invite recipient: {first_invite_sent['Recipient']}")
    print(f'first invite sent: {first_invite_sent_date}')

    # convert dates to datetime object and find the delta
    first_invite_sent_string = first_invite_sent['Email Sent'].strftime('%Y-%m-%d')
    first_invite_sent_datetime = datetime.strptime(first_invite_sent_string, '%Y-%m-%d')

    session_date_split = session_date.split(' ')
    session_date_fixed = f'{session_date_split[0]}'
    start_time_datetime = datetime.strptime(session_date_fixed, '%Y-%m-%d')
    delta = start_time_datetime - first_invite_sent_datetime
    print(f'first invite sent {delta.days} days before event')


# GET TOTAL NUMBER OF REGISTRATIONS AND THE FIRST TO REGISTER
def registration_details(general_ticket_reg_df):
    total_reg = 0
    for index, row in general_ticket_reg_df.iterrows():
        if '@zoom.us' not in row['Registrant email']:
            total_reg += 1
    print(f'total registrations: {total_reg}')
    
    # FIRST CUSTOMER TO REGISTER
    reg_list_by_date = general_ticket_reg_df.sort_values('Register date (UTC)')
    first_reg_utc = reg_list_by_date.iloc[0]
    first_reg_cust = first_reg_utc['Display name']
    print(f"first registered customer: {first_reg_cust}")

    # convert time from UTC to EST (because tz_localize is currently EDT)
    first_reg_edt = first_reg_utc['Register date (UTC)'] - timedelta(hours=5)
    print(f"registered on: {first_reg_edt} EST")
    return first_reg_cust

def first_reg_se(general_ticket_reg_df, invite_list_df, first_reg_cust):
    # REGISTRATIONS BY SE
    # drop unneeded columns
    reduced_reg_df = general_ticket_reg_df.drop(['First name', 'Last name', 'Authentication method', 'Agree to receive marketing communication?', 'Register date (UTC)', 'Ticket name',
                                                 'Marketing consent pre-checked?', 'Job title_*_Webinar Participant Ticket', 'Organization_*_Webinar Participant Ticket', 
                                                 'Event experience', 'Event access name', 'Registration method', 'Access type', 'Source of registration', 'Unique identifier'], axis=1)

    # drop rows with Zoom users
    for index, row in reduced_reg_df.iterrows():
        if '@zoom.us' in row['Registrant email']: 
            reduced_reg_df.drop([index], inplace=True)
    reduced_reg_df.reset_index(inplace=True, drop=True)
    # split the domain to compare
    reduced_reg_df[['temp','domain']] = reduced_reg_df['Registrant email'].str.split('@', expand=True)

    # drop the temp column
    reduced_reg_df.drop(['temp'], axis=1, inplace=True)

    # add SE column based on matching email in invite list
    ## make the invite_list_df and general_ticket_reg_df identically-labeled
    invite_list_renamed = invite_list_df.rename(columns={'Recipient': 'Registrant email'})
    invite_list_renamed['Registrant email'] = invite_list_renamed['Registrant email'].str.strip()
    invite_list_renamed['Registrant email'] = invite_list_renamed['Registrant email'].str.lower()
    invite_list_renamed[['temp', 'domain']] = invite_list_renamed['Registrant email'].str.split('@', expand=True)
    invite_list_renamed.drop(['temp', 'Registrant email'], axis=1, inplace=True)
    invite_list_renamed.drop(['First name', 'Last name', 'Email Sent', '<Custom info by SE>'], axis=1, inplace=True)
    invite_list_renamed.dropna(inplace=True)
    invite_list_renamed.drop_duplicates(inplace=True)
    reg_per_se = reduced_reg_df.merge(invite_list_renamed, how='left', on='domain')
    first_reg_se = reg_per_se.loc[reg_per_se['Display name'] == first_reg_cust, 'SE'].iloc[0]
    # SE FOR FIRST REGISTRATION ABOVE
    print(f"first registered customer's SE: {first_reg_se}")
    return reg_per_se

def all_reg_se(reg_per_se):
    
    # SE WITH MOST REGISTRATIONS: 
    print(f"SE with most registrations: {reg_per_se.value_counts('SE').index[0]} with {reg_per_se.value_counts('SE').iloc[0]}")
    # TOTAL REGISTRATIONS BY SE: 
    print('SE registration totals:')
    print(reg_per_se.value_counts('SE').to_string(dtype=False))
        
    # NUMBER OF REGISTRATIONS BY CUSTOMER:
    print(f"customer with the most registrations: {reg_per_se.value_counts('domain').index[0]} with {reg_per_se.value_counts('domain').iloc[0]}")
    print('number of registrations by domain: ')
    print(reg_per_se.value_counts('domain').head().to_string())
    
# GET SESSION ATTENDEES
def attendee_details(event_attendance_df):
    total_attendees_minus_zoom = 0

    for index, row in event_attendance_df.iterrows():
        if '@zoom.us' not in row['Registrant email']: 
            total_attendees_minus_zoom += 1
    print(f'total attendees: {total_attendees_minus_zoom}')
    
def all_attendees_se(reg_per_se):
    # SE WITH MOST ATTENDEES: 
    attendees_per_se = reg_per_se.drop(reg_per_se[reg_per_se['Event attendance'] == 'Absent'].index)

    # TOTAL ATTENDEES BY SE: 
    print(f"SE with most attendees: {attendees_per_se.value_counts('SE').index[0]} with {attendees_per_se.value_counts('SE').iloc[0]}")
    print('SE attendee totals: ')
    print(attendees_per_se.value_counts('SE').to_string(dtype=False))
        
    # NUMBER OF ATTENDEES BY CUSTOMER: 
    print(f"customer with the most attendees: {attendees_per_se.value_counts('domain').index[0]} with {attendees_per_se.value_counts('domain').iloc[0]}")
    print('number of attendees per customer')
    print(attendees_per_se.value_counts('domain').head().to_string())
    return attendees_per_se


def poll_ratings_details(poll_results_df, attendees_per_se):
    # POLLS PER CUSTOMER: 
    poll_results_simplified = poll_results_df.drop(['#', 'User name', 'First name', 'Last name', 'Submitted date/time'], axis=1)
    poll_results_simplified = poll_results_simplified.rename(columns={'User email': 'Registrant email'})
    attendees_with_poll = attendees_per_se.merge(poll_results_simplified, how='left', on='Registrant email')

    # EVENT RATING
    event_rating_avg = attendees_with_poll['1.Overall, how would you rate this event?'].mean(skipna=True)
    event_rating_avg = round(event_rating_avg, 2)
    print(f'Average Event Rating {event_rating_avg}/5')
    lowest_rating = attendees_with_poll['1.Overall, how would you rate this event?'].min()
    lowest_rating_cust = attendees_with_poll.loc[attendees_with_poll['1.Overall, how would you rate this event?'] == lowest_rating, 'Display name']
    print(f'lowest rating: {lowest_rating}/5, by: ')
    for customer in lowest_rating_cust:
        print(customer)

    # EVENT RECOMMENDATION
    event_recommend_avg = attendees_with_poll['2.How likely would you be to recommend this event to others?'].mean(skipna=True)
    event_recommend_avg = round(event_recommend_avg, 2)
    print(f'average recommendation score: {event_recommend_avg}/10')
    lowest_recommend = attendees_with_poll['2.How likely would you be to recommend this event to others?'].min()
    lowest_recommend_cust = attendees_with_poll.loc[attendees_with_poll['2.How likely would you be to recommend this event to others?'] == lowest_recommend, 'Display name']
    print(f'lowest recommendation rating: {lowest_recommend}/10 by: ')
    for customer in lowest_recommend_cust: 
        print(customer)
        
    return attendees_with_poll

def interest_feedback(attendees_with_poll):
    # GET FOLLOWING TOPIC INTEREST AND FEEDBACK
    topics = []
    topics2 = []
    attendees_with_poll.fillna('nan', inplace=True)
    for index, row in attendees_with_poll.iterrows(): 
        if row['3.Would you like to have a dedicated session for you and your team to hear more about any of the following topics? (select all that apply) Your Zoom Solution Engineer will reach out directly to schedule time with you.'] != 'nan':
            topics.append(row['3.Would you like to have a dedicated session for you and your team to hear more about any of the following topics? (select all that apply) Your Zoom Solution Engineer will reach out directly to schedule time with you.'])
    for topic in topics: 
        t = topic.split(';')
        for item in t: 
            topics2.append(item)
            
    product_interest_df = pd.DataFrame(topics2, columns=['products'])
    product_interest_df['products'].value_counts()


if __name__=='__main__': 
    
    
    print('Zoom Healthcare Quarterly Tech Talks Reporting Analytics')
    file_path = input('Paste in the FULL file path to combined .xlsx file: ')
    invite_list_df, event_attendance_df, general_ticket_reg_df, poll_results_df, session_summary_df = get_dfs(file_path)
    session_date = session_details(session_summary_df)
    invite_details(invite_list_df, session_date)
    first_reg_cust = registration_details(general_ticket_reg_df)
    attendee_details(event_attendance_df)
    reg_per_se = first_reg_se(general_ticket_reg_df, invite_list_df, first_reg_cust)
    all_reg_se(reg_per_se)
    attendees_per_se = all_attendees_se(reg_per_se)
    attendees_with_poll = poll_ratings_details(poll_results_df, attendees_per_se)
    interest_feedback(attendees_with_poll)
    
    
    
    
    
    
    