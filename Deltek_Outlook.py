import datetime as dt
import pandas as pd
import numpy as np
import win32com.client

def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%d/%m/%Y') + "' AND [END] <= '" + end.strftime('%d/%m/%Y') + "'"
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
    cal_subject = np.asarray(cal_subject)
    [f,c] = np.shape(cal_subject)
    for i in range(0,f):
        cal_subject[i,1] = cal_subject[i,1].replace(' ','')
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
Start_Time  = dt.datetime(2022,1,1)
End_Time    = dt.datetime(2022,2,1)
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
Tmp[np.isnan(Tmp)] = 0
Tmp1        = pd.DataFrame(columns=np.arange(Start_Time, End_Time, dtype='datetime64[D]'))
Report      = pd.concat([Tmp1, Tmp])

# ----------------------------------------------------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------------------------------------------------
Summary = pd.DataFrame(index=Report.index.values,columns=['Horas', 'Porc'])
Summary['Horas'] = Report.sum(1)
Summary['Porc']  = Summary['Horas']/(Summary['Horas'].sum() - Summary.loc['01-Holiday']['Horas'])

Report = Report.drop('01-Holiday')

# ----------------------------------------------------------------------------------------------------------------------
# Save Results
# ----------------------------------------------------------------------------------------------------------------------
Results.to_excel('01-Report.xlsx')
Report.to_csv('02-Deltek.csv', index_label='Name')
Summary.to_csv('03-Total_Deltek.csv', index_label='Name')
