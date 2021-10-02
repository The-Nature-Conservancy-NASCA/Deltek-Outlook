import datetime as dt
import pandas as pd
import numpy as np
import win32com.client

def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

def get_appointments(calendar,subject_kw = None,exclude_subject_kw = None, body_kw = None):
    if subject_kw == None:
        appointments = [app for app in calendar]
    else:
        appointments = [app for app in calendar if subject_kw in app.subject]

    if exclude_subject_kw != None:
        appointments = [app for app in appointments if exclude_subject_kw not in app.subject]

    Tmp         = [app.subject for app in appointments]
    cal_subject = [app.split('|') for app in Tmp]
    cal_subject = np.array(cal_subject)
    cal_date    = [app.start.strftime('%d/%m/%Y') for app in appointments]
    cal_date    = pd.to_datetime(cal_date,format='%d/%m/%Y')
    cal_start   = [app.start.strftime('%d/%m/%Y %H:%M:%S')for app in appointments]
    cal_start   = pd.to_datetime(cal_start)
    cal_end     = [app.end.strftime('%d/%m/%Y %H:%M:%S') for app in appointments]
    cal_end     = pd.to_datetime(cal_end)
    cal_body    = [app.body for app in appointments]
    Tmp         = cal_end - cal_start
    Hour        = [(app.seconds)/3600 for app in Tmp]

    df = pd.DataFrame({'Project': cal_subject[:,1],
                       'Activity': cal_subject[:, 2],
                       'Date': cal_date,
                       'Start_Time': cal_start,
                       'End_Time': cal_end,
                       'Hours': Hour,
                       'Description': cal_body})

    df.index = cal_date
    return df

# ----------------------------------------------------------------------------------------------------------------------
# Input Data
# ----------------------------------------------------------------------------------------------------------------------
Start_Time  = dt.datetime(2021,10,1)
End_Time    = dt.datetime(2021,10,31)
keyword     = 'Deltek'

# ----------------------------------------------------------------------------------------------------------------------
# Input Data
# ----------------------------------------------------------------------------------------------------------------------
RawData     = get_calendar(Start_Time, End_Time)
Results     = get_appointments(RawData, keyword)

# ----------------------------------------------------------------------------------------------------------------------
# Result Aggregation for Deltek
# ----------------------------------------------------------------------------------------------------------------------
Tmp         = Results.groupby(by=['Project','Date'], as_index=False)['Hours'].sum()
Tmp         = Tmp.pivot(index='Project', columns='Date', values='Hours')
Tmp1        = pd.DataFrame(columns=np.arange(Start_Time, End_Time, dtype='datetime64[D]'))
Report      = pd.concat([Tmp1, Tmp])

# ----------------------------------------------------------------------------------------------------------------------
# Save Results
# ----------------------------------------------------------------------------------------------------------------------
Results.to_excel('01-Report.xlsx')
Report.to_excel('02-Deltek.xlsx')