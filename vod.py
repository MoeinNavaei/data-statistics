    
def vod(vod_tva, vod_lenz, vod_aio):
        
    import xlsxwriter  
    import pandas as pd
    #from pandas import DataFrame
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    import arabic_reshaper
    # from bidi.algorithm import get_display
    import matplotlib as mpl
    import matplotlib.ticker as tkr
    import numpy as np
    from matplotlib.ticker import FuncFormatter
    from mpl_toolkits.mplot3d import Axes3D
    import time
    import re
    import requests
        
    print("start vod")
    
    print("start tva vod")
    vod_tva['Video'].replace('', 'nan', inplace=True)
    vod_tva.dropna(subset=['Video'], inplace=True)
    
    vod_tva=vod_tva.groupby(['Video']).sum().reset_index()
    vod_tva['duration (minute)']=round(vod_tva['Sessions']*vod_tva['Avg. Duration (sec)']/60, 0)
    
    vod_tva.insert(5, 'Name content summary', '')
    vod_tva_content_name=vod_tva['Video']
    n=len(vod_tva_content_name)
    vod_tva['Video'] = vod_tva['Video'].astype(str)
    for i in range(0,n):
         xx_name_content=vod_tva_content_name[i]
         head, sep, tail = xx_name_content.partition('(')
         vod_tva.iat[i, 5] = head
    
    vod_tva.insert(6, 'episode', 1)
    tva_vod_pivot=vod_tva.groupby(['Name content summary']).sum().reset_index()
    
    tva_vod_serial=tva_vod_pivot.query("episode != 1")
    tva_vod_film=tva_vod_pivot.query("episode == 1")
    
    tva_vod_serial_content=tva_vod_serial.copy()
    tva_vod_serial_content=len(tva_vod_serial_content)
    tva_vod_serial_visit=tva_vod_serial['Sessions'].sum()
    tva_vod_serial_duration=tva_vod_serial['duration (minute)'].sum()
    tva_vod_serial.sort_values('Sessions', axis = 0, ascending = False, inplace = True, na_position ='last')
    tva_vod_serial_popular_visit=tva_vod_serial.iloc[0:10 , [0, 1]]
    tva_vod_serial_popular_visit=tva_vod_serial_popular_visit.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی- تیوا', 'Sessions': 'تعداد بازدید'})
    
    tva_vod_serial.sort_values('duration (minute)', axis = 0, ascending = False, inplace = True, na_position ='last')
    tva_vod_serial_popular_duration=tva_vod_serial.iloc[0:10 , [0, 6]]
    tva_vod_serial_popular_duration=tva_vod_serial_popular_duration.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ مدت زمان بازدید- تیوا', 'duration (minute)': 'مدت زمان بازدید (به دقیقه)'})
    
    tva_vod_serial['visit middle']=round(tva_vod_serial['Sessions']/tva_vod_serial['episode'], 0)
    tva_vod_serial.sort_values('visit middle', axis = 0, ascending = False, inplace = True, na_position ='last')
    tva_vod_serial_popular_visit_middle=tva_vod_serial.iloc[0:10 , [0, 7]]
    tva_vod_serial_popular_visit_middle=tva_vod_serial_popular_visit_middle.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ متوسط بازدید هر قسمت- تیوا', 'visit middle': 'تعداد بازدید'})
    
    tva_vod_serial['duration middle']=round(tva_vod_serial['Avg. Duration (sec)']/tva_vod_serial['episode'], 0)
    tva_vod_serial.sort_values('duration middle', axis = 0, ascending = False, inplace = True, na_position ='last')
    tva_vod_serial_popular_duration_middle=tva_vod_serial.iloc[0:10 , [0, 8]]
    tva_vod_serial_popular_duration_middle=tva_vod_serial_popular_duration_middle.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ متوسط زمان بازدید هر قسمت- تیوا', 'duration middle': 'مدت زمان بازدید (به دقیقه)'})
    
    tva_vod_film_content=tva_vod_film.copy()
    tva_vod_film_content=len(tva_vod_film_content)
    tva_vod_film_visit=tva_vod_film['Sessions'].sum()
    tva_vod_film_duration=tva_vod_film['duration (minute)'].sum()
    tva_vod_film.sort_values('Sessions', axis = 0, ascending = False, inplace = True, na_position ='last')
    tva_vod_film_popular_visit=tva_vod_film.iloc[0:10 , [0, 1]]
    tva_vod_film_popular_visit=tva_vod_film_popular_visit.rename(columns={'Name content summary': 'نام محتواهای پربازدید فیلم- تیوا', 'Sessions': 'تعداد بازدید'})
    tva_vod_film.sort_values('duration (minute)', axis = 0, ascending = False, inplace = True, na_position ='last')
    tva_vod_film_popular_duration=tva_vod_film.iloc[0:10 , [0, 6]]
    tva_vod_film_popular_duration=tva_vod_film_popular_duration.rename(columns={'Name content summary': 'نام محتواهای پربازدید فیلم به لحاظ مدت زمان بازدید- تیوا', 'duration (minute)': 'مدت زمان بازدید (به دقیقه)'})
    
    tva_vod_content=tva_vod_film_content+tva_vod_serial_content
    tva_vod_visit=tva_vod_serial_visit+tva_vod_film_visit
    tva_vod_duration=tva_vod_serial_duration+tva_vod_film_duration
    
    print("End tva vod")
    
    print("start lenz vod")
    vod_lenz = vod_lenz.drop([len(vod_lenz)-1])
    vod_lenz = vod_lenz.rename(columns={"content name":"content_name"})
    vod_lenz['content_name'].replace('', 'nan', inplace=True)
    vod_lenz.dropna(subset=['content_name'], inplace=True)
    vod_lenz['content_name'].replace('', 'NO', inplace=True)
    vod_lenz = vod_lenz [~vod_lenz.content_name.str.contains('NO')]
    
    vod_lenz.insert(6, 'Name content summary', '')
    vod_lenz_content_name=vod_lenz['content_name']
    n=len(vod_lenz_content_name)
    
    for i in range(0,n):
         x_name_content=vod_lenz_content_name[i]
         head, sep, tail = x_name_content.partition('قسمت')
         vod_lenz.iat[i, 6] = head
    
    vod_lenz.insert(7, 'episode', 1)
    lenz_vod_pivot=vod_lenz.groupby(['Name content summary']).sum().reset_index()
    
    lenz_vod_serial=lenz_vod_pivot.query("episode != 1")
    lenz_vod_film=lenz_vod_pivot.query("episode == 1")
    
    lenz_vod_serial_content=lenz_vod_serial.copy()
    lenz_vod_serial_content=len(lenz_vod_serial_content)
    lenz_vod_serial_visit=lenz_vod_serial['access times'].sum()
    lenz_vod_serial_duration=lenz_vod_serial['access duration (hour)'].sum()
    lenz_vod_serial_duration=round(lenz_vod_serial_duration*60, 0)
    lenz_vod_serial.sort_values('access times', axis = 0, ascending = False, inplace = True, na_position ='last')
    lenz_vod_serial_popular_visit=lenz_vod_serial.iloc[0:10 , [0, 2]]
    lenz_vod_serial_popular_visit=lenz_vod_serial_popular_visit.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی- لنز', 'access times': 'تعداد بازدید'})
    
    lenz_vod_serial.sort_values('access duration (hour)', axis = 0, ascending = False, inplace = True, na_position ='last')
    lenz_vod_serial_popular_duration=lenz_vod_serial.iloc[0:10 , [0, 3]]
    lenz_vod_serial_popular_duration=lenz_vod_serial_popular_duration.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ مدت زمان بازدید- لنز', 'access duration (hour)': 'مدت زمان بازدید (به دقیقه)'})
    
    lenz_vod_serial['visit middle']=round(lenz_vod_serial['access times']/lenz_vod_serial['episode'], 0)
    lenz_vod_serial.sort_values('visit middle', axis = 0, ascending = False, inplace = True, na_position ='last')
    lenz_vod_serial_popular_visit_middle=lenz_vod_serial.iloc[0:10 , [0, 6]]
    lenz_vod_serial_popular_visit_middle=lenz_vod_serial_popular_visit_middle.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ متوسط بازدید هر قسمت- لنز', 'visit middle': 'تعداد بازدید'})
    
    lenz_vod_serial['duration middle']=round(lenz_vod_serial['access duration (hour)']/lenz_vod_serial['episode'], 0)
    lenz_vod_serial.sort_values('duration middle', axis = 0, ascending = False, inplace = True, na_position ='last')
    lenz_vod_serial_popular_duration_middle=lenz_vod_serial.iloc[0:10 , [0, 7]]
    lenz_vod_serial_popular_duration_middle=lenz_vod_serial_popular_duration_middle.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ متوسط زمان بازدید هر قسمت- لنز', 'duration middle': 'مدت زمان بازدید (به دقیقه)'})
    
    lenz_vod_film_content=lenz_vod_film.copy()
    lenz_vod_film_content=len(lenz_vod_film_content)
    lenz_vod_film_visit=lenz_vod_film['access times'].sum()
    lenz_vod_film_duration=lenz_vod_film['access duration (hour)'].sum()
    lenz_vod_film_duration=round(lenz_vod_film_duration*60, 0)
    lenz_vod_film.sort_values('access times', axis = 0, ascending = False, inplace = True, na_position ='last')
    lenz_vod_film_popular_visit=lenz_vod_film.iloc[0:10 , [0, 2]]
    lenz_vod_film_popular_visit=lenz_vod_film_popular_visit.rename(columns={'Name content summary': 'نام محتواهای پربازدید فیلم- لنز', 'access times': 'تعداد بازدید'})
    
    lenz_vod_film.sort_values('access duration (hour)', axis = 0, ascending = False, inplace = True, na_position ='last')
    lenz_vod_film_popular_duration=lenz_vod_film.iloc[0:10 , [0, 3]]
    lenz_vod_film_popular_duration=lenz_vod_film_popular_duration.rename(columns={'Name content summary': 'نام محتواهای پربازدید فیلم به لحاظ مدت زمان بازدید- لنز', 'access duration (hour)': 'مدت زمان بازدید (به دقیقه)'})
    
    
    lenz_vod_content=lenz_vod_film_content+lenz_vod_serial_content
    lenz_vod_visit=lenz_vod_serial_visit+lenz_vod_film_visit
    lenz_vod_duration=lenz_vod_serial_duration+lenz_vod_film_duration
    print("End lenz vod")
    
    print("start aio vod")
    vod_aio['title_fa'].replace('', 'NO', inplace=True)
    vod_aio = vod_aio [~vod_aio.title_fa.str.contains('NO')]
    
    
#    del vod_aio['id']
    del vod_aio['publish_date']
    del vod_aio['imdb_rank']
    del vod_aio['country_fa']
    del vod_aio['genre']
    del vod_aio['cast']
    
    vod_aio.insert(2, 'season', '')
    vod_aio.insert(3, 'epizode', '')
    vod_aio.insert(4, 'duration (minute)', '')
    
    vod_aio_serial = vod_aio [vod_aio.title_fa.str.contains('فصل')]  
    vod_aio_film = vod_aio [~vod_aio.title_fa.str.contains('فصل')]  
    vod_aio_film = vod_aio_film.reset_index()
    del vod_aio_film['index']
    vod_aio_serial = vod_aio_serial.reset_index()
    del vod_aio_serial['index']    
    
    for i in range(0, len(vod_aio_serial)):
        print(i)
        x_name_content = vod_aio_serial.loc[i, 'title_fa']
        head, sep, tail = x_name_content.partition('فصل')
        vod_aio_serial.loc[i, 'title_fa'] = head
        vod_aio_serial.loc[i, 'season'] = tail
        
    for i in range(0, len(vod_aio_serial)):
        print(i)
        x_name_content = vod_aio_serial.loc[i, 'season']
        head, sep, tail = x_name_content.partition('قسمت')
        vod_aio_serial.loc[i, 'season'] = head
        vod_aio_serial.loc[i, 'epizode'] = tail

    vod_aio_serial['title_fa'] = vod_aio_serial['title_fa'].str.strip()
    vod_aio_serial['season'] = vod_aio_serial['season'].str.strip()
    vod_aio_serial['epizode'] = vod_aio_serial['epizode'].str.strip()
    
    aio_vod_serial = vod_aio_serial.copy()
    aio_vod_film = vod_aio_film.copy()
    aio_vod_serial_content = vod_aio_serial.copy()
    aio_vod_serial_content.drop_duplicates(subset =['title_fa'], keep = 'last', inplace = True)
    aio_vod_serial_content = len(aio_vod_serial_content)
    aio_vod_film_content = vod_aio_film.copy()
    aio_vod_film_content.drop_duplicates(subset =['title_fa'], keep = 'last', inplace = True)
    aio_vod_film_content = len(aio_vod_film_content)
    aio_vod_content = aio_vod_serial_content + aio_vod_film_content
    aio_vod_film_visit = vod_aio_film['viewer'].sum()
    aio_vod_serial_visit = vod_aio_serial['viewer'].sum()
    aio_vod_visit = aio_vod_film_visit + aio_vod_serial_visit
    aio_vod_serial_duration = aio_vod_serial['duration (minute)'].sum()
    aio_vod_film_duration = aio_vod_film['duration (minute)'].sum()
    aio_vod_duration = aio_vod_serial_duration + aio_vod_film_duration
    
    vod_aio_film.sort_values('viewer', axis = 0, ascending = False, inplace = True, na_position ='last')
    aio_vod_film_popular_visit = vod_aio_film.iloc[0:10 , [0, 1]]
    aio_vod_film_popular_visit=aio_vod_film_popular_visit.rename(columns={'title_fa': 'نام محتواهای پربازدید فیلم- آیو', 'viewer': 'تعداد بازدید'})
   
    aio_vod_film_popular_duration=aio_vod_film.iloc[0:10 , [0, 4]]
    aio_vod_film_popular_duration=aio_vod_film_popular_duration.rename(columns={'title_fa': 'نام محتواهای پربازدید فیلم به لحاظ مدت زمان بازدید- آیو', 'duration (minute)': 'مدت زمان بازدید (به دقیقه)'})
    
    aio_vod_serial_popular_duration=aio_vod_serial.iloc[0:10 , [0, 4]]
    aio_vod_serial_popular_duration=aio_vod_serial_popular_duration.rename(columns={'title_fa': 'نام محتواهای پربازدید سریالی به لحاظ مدت زمان بازدید- آیو', 'duration (minute)': 'مدت زمان بازدید (به دقیقه)'})
    
    aio_vod_serial_popular_duration_middle=aio_vod_serial.iloc[0:10 , [0, 4]]
    aio_vod_serial_popular_duration_middle=aio_vod_serial_popular_duration_middle.rename(columns={'title_fa': 'نام محتواهای پربازدید سریالی به لحاظ متوسط زمان بازدید هر قسمت- آیو', 'duration middle': 'مدت زمان بازدید (به دقیقه)'})
    
     
    
    aio_vod_serial_popular_visit = vod_aio_serial.groupby(['title_fa', 'season']).sum().reset_index()
    aio_vod_serial_popular_visit.sort_values('viewer', axis = 0, ascending = False, inplace = True, na_position ='last')
    aio_vod_serial_popular_visit = aio_vod_serial_popular_visit.reset_index()
    del aio_vod_serial_popular_visit['index']
    aio_vod_serial_popular_visit = aio_vod_serial_popular_visit.iloc[0:10 , [0, 2]]
    aio_vod_serial_popular_visit=aio_vod_serial_popular_visit.rename(columns={'title_fa': 'نام محتواهای پربازدید سریالی- آیو', 'viewer': 'تعداد بازدید'})
    

    
    aio_vod_serial_popular_visit_middle1 = aio_vod_serial.groupby(['title_fa', 'season']).count().reset_index()
    aio_vod_serial_popular_visit_middle2 = aio_vod_serial.groupby(['title_fa', 'season']).sum().reset_index()
    aio_vod_serial_popular_visit_middle1=aio_vod_serial_popular_visit_middle1.rename(columns={"viewer":"count"})
    aio_vod_serial_popular_visit_middle2=aio_vod_serial_popular_visit_middle2.rename(columns={"viewer":"visit"})
    aio_vod_serial_popular_visit_middle = pd.merge(aio_vod_serial_popular_visit_middle1, aio_vod_serial_popular_visit_middle2, on = ['title_fa', 'season'])
    aio_vod_serial_popular_visit_middle.drop_duplicates(subset =['title_fa', 'season'], keep = 'last', inplace = True)
    aio_vod_serial_popular_visit_middle.sort_values('visit', axis = 0, ascending = False, inplace = True, na_position ='last')
    aio_vod_serial_popular_visit_middle = aio_vod_serial_popular_visit_middle.reset_index()
    del aio_vod_serial_popular_visit_middle['index']
    aio_vod_serial_popular_visit_middle = aio_vod_serial_popular_visit_middle.iloc[0:10 , [0, 5]]
    aio_vod_serial_popular_visit_middle=aio_vod_serial_popular_visit_middle.rename(columns={'title_fa': 'نام محتواهای پربازدید سریالی به لحاظ متوسط بازدید هر قسمت- آیو', 'visit': 'تعداد بازدید'})
    
    
    
    
#    del vod_aio['content_id']
#    del vod_aio['year']
#    del vod_aio['rating']
#    del vod_aio['name']
#    del vod_aio['genre']
#    del vod_aio['category']
#    del vod_aio['country']
#    del vod_aio['عوامل']
#    vod_aio.to_excel('aio_vod_film0.xlsx', index=False)
#    vod_aio.drop_duplicates(subset =['title', 'viewer', 'season_number', 'episode_number'], keep = 'first', inplace = True)
#    vod_aio.insert(4, 'duration (minute)', 0)
#    vod_aio=vod_aio.rename(columns={"episode_number":"episode"})
#    vod_aio=vod_aio.rename(columns={"season_number":"season"})
#    vod_aio.insert(5, 'counter', 1)
#    vod_aio['episode'] = vod_aio['episode'].astype(str)
#    vod_aio['episode'] = vod_aio['episode'].str.replace('nan' , 'NO')
#    aio_vod_film=vod_aio.query("episode == 'NO'")
#    aio_vod_serial=vod_aio.query("episode != 'NO'")
#    aio_vod_serial['episode'] = aio_vod_serial['episode'].astype(str).replace('\.0', '', regex=True)
#    aio_vod_serial['episode'] = aio_vod_serial['episode'].astype(int)
#
#    del aio_vod_serial['episode']
#    aio_vod_serial = aio_vod_serial.rename(columns={"counter":"episode"})
#    aio_vod_serial=aio_vod_serial.groupby(['title', 'season']).sum().reset_index()
#    aio_vod_film=aio_vod_film.groupby(['title']).sum().reset_index()
#    
#    
#    
#    aio_vod_serial_content=aio_vod_serial.copy()
#    aio_vod_serial_content=len(aio_vod_serial_content)
#    aio_vod_serial_visit=aio_vod_serial['viewer'].sum()
#    aio_vod_serial_duration=aio_vod_serial['duration (minute)'].sum()
#    aio_vod_serial.sort_values('viewer', axis = 0, ascending = False, inplace = True, na_position ='last')
#    aio_vod_serial_popular_visit=aio_vod_serial.iloc[0:10 , [0, 2]]
#    aio_vod_serial_popular_visit=aio_vod_serial_popular_visit.rename(columns={'title': 'نام محتواهای پربازدید سریالی- آیو', 'viewer': 'تعداد بازدید'})
#    
#    aio_vod_serial.sort_values('duration (minute)', axis = 0, ascending = False, inplace = True, na_position ='last')
#    aio_vod_serial_popular_duration=aio_vod_serial.iloc[0:10 , [0, 3]]
#    aio_vod_serial_popular_duration=aio_vod_serial_popular_duration.rename(columns={'title': 'نام محتواهای پربازدید سریالی به لحاظ مدت زمان بازدید- آیو', 'duration (minute)': 'مدت زمان بازدید (به دقیقه)'})
#    
#    aio_vod_serial['visit middle']=round(aio_vod_serial['viewer']/aio_vod_serial['episode'], 0)
#    aio_vod_serial.sort_values('visit middle', axis = 0, ascending = False, inplace = True, na_position ='last')
#    aio_vod_serial_popular_visit_middle=aio_vod_serial.iloc[0:10 , [0, 5]]
#    aio_vod_serial_popular_visit_middle=aio_vod_serial_popular_visit_middle.rename(columns={'title': 'نام محتواهای پربازدید سریالی به لحاظ متوسط بازدید هر قسمت- آیو', 'visit middle': 'تعداد بازدید'})
#    
#    aio_vod_serial['duration middle']=round(aio_vod_serial['duration (minute)']/aio_vod_serial['episode'], 0)
#    aio_vod_serial.sort_values('duration middle', axis = 0, ascending = False, inplace = True, na_position ='last')
#    aio_vod_serial_popular_duration_middle=aio_vod_serial.iloc[0:10 , [0, 6]]
#    aio_vod_serial_popular_duration_middle=aio_vod_serial_popular_duration_middle.rename(columns={'title': 'نام محتواهای پربازدید سریالی به لحاظ متوسط زمان بازدید هر قسمت- آیو', 'duration middle': 'مدت زمان بازدید (به دقیقه)'})
#    
#    aio_vod_film_content=aio_vod_film.copy()
#    aio_vod_film_content=len(aio_vod_film_content)
#    aio_vod_film_visit=aio_vod_film['viewer'].sum()
#    aio_vod_film_duration=aio_vod_film['duration (minute)'].sum()
#    aio_vod_film.sort_values('viewer', axis = 0, ascending = False, inplace = True, na_position ='last')
#    aio_vod_film_popular_visit=aio_vod_film.iloc[0:10 , [0, 1]]
#    aio_vod_film_popular_visit=aio_vod_film_popular_visit.rename(columns={'title': 'نام محتواهای پربازدید فیلم- آیو', 'viewer': 'تعداد بازدید'})
#    aio_vod_film.sort_values('duration (minute)', axis = 0, ascending = False, inplace = True, na_position ='last')
#    aio_vod_film_popular_duration=aio_vod_film.iloc[0:10 , [0, 3]]
#    aio_vod_film_popular_duration=aio_vod_film_popular_duration.rename(columns={'title': 'نام محتواهای پربازدید فیلم به لحاظ مدت زمان بازدید- آیو', 'duration (minute)': 'مدت زمان بازدید (به دقیقه)'})
#    
#    aio_vod_content=aio_vod_film_content+aio_vod_serial_content
#    aio_vod_visit=aio_vod_serial_visit+aio_vod_film_visit
#    aio_vod_duration=aio_vod_serial_duration+aio_vod_film_duration
#    
#    print("End aio vod")
#    
#    aio_vod_film.to_excel('aio_vod_film.xlsx', index=False)
#    
    
#    
#    
#    
    print("vod statistics summary")
    
    vod_statistics_summary={'operators': ['تیوا', 'لنز', 'آیو'],
           'all_content': [tva_vod_content, lenz_vod_content, aio_vod_content],
           'total_visit': [tva_vod_visit, lenz_vod_visit, aio_vod_visit],
           'total_duration': [tva_vod_duration, lenz_vod_duration, aio_vod_duration],
           'vod_serial_content': [tva_vod_serial_content, lenz_vod_serial_content, aio_vod_serial_content],
           'vod_film_content': [tva_vod_film_content, lenz_vod_film_content, aio_vod_film_content],
           'vod_serial_visit': [tva_vod_serial_visit, lenz_vod_serial_visit, aio_vod_serial_visit],
           'vod_film_visit': [tva_vod_film_visit, lenz_vod_film_visit, aio_vod_film_visit],
           'vod_serial_duration': [tva_vod_serial_duration, lenz_vod_serial_duration, aio_vod_serial_duration],
           'vod_film_duration': [tva_vod_film_duration, lenz_vod_film_duration, aio_vod_film_duration],}
    vod_statistics_summary=pd.DataFrame(vod_statistics_summary, columns=['operators', 'all_content', 
                                                                         'total_visit', 'total_duration', 
                                                                         'vod_serial_content', 'vod_film_content', 
                                                                         'vod_serial_visit', 'vod_film_visit', 
                                                                         'vod_serial_duration', 'vod_film_duration'])
    
    vod_statistics_summary=vod_statistics_summary.rename(columns={'operators': 'اپراتور', 'all_content': 'تعداد کل محتوا',
                                                                  'total_visit': 'تعداد کل بازدید', 'total_duration': 'کل مدت زمان بازدید (به دقیقه)',
                                                                  'vod_serial_content': 'تعداد محتواهای سریالی', 'vod_film_content': 'تعداد محتواهای فیلم',
                                                                  'vod_serial_visit': 'تعداد بازدید از محتواهای سریالی', 'vod_film_visit': 'تعداد بازدید از محتواهای فیلم',
                                                                  'vod_serial_duration': 'مدت زمان بازدید از محتواهای سریالی (به دقیقه)', 'vod_film_duration': 'مدت زمان بازدید از محتواهای فیلم (به دقیقه)',})
    
    print("vod statistics popular content")
    
    print("edit of tva vod")
    tva_vod_serial_popular_visit.to_excel('busy/tva_vod_serial_popular_visit.xlsx')
    tva_vod_serial_popular_visit=pd.read_excel('busy/tva_vod_serial_popular_visit.xlsx')
    tva_vod_serial_popular_duration.to_excel('busy/tva_vod_serial_popular_duration.xlsx')
    tva_vod_serial_popular_duration=pd.read_excel('busy/tva_vod_serial_popular_duration.xlsx')
    tva_vod_serial_popular_visit_middle.to_excel('busy/tva_vod_serial_popular_visit_middle.xlsx')
    tva_vod_serial_popular_visit_middle=pd.read_excel('busy/tva_vod_serial_popular_visit_middle.xlsx')
    tva_vod_serial_popular_duration_middle.to_excel('busy/tva_vod_serial_popular_duration_middle.xlsx')
    tva_vod_serial_popular_duration_middle=pd.read_excel('busy/tva_vod_serial_popular_duration_middle.xlsx')
    tva_vod_film_popular_visit.to_excel('busy/tva_vod_film_popular_visit.xlsx')
    tva_vod_film_popular_visit=pd.read_excel('busy/tva_vod_film_popular_visit.xlsx')
    tva_vod_film_popular_duration.to_excel('busy/tva_vod_film_popular_duration.xlsx')
    tva_vod_film_popular_duration=pd.read_excel('busy/tva_vod_film_popular_duration.xlsx')
    
    del tva_vod_serial_popular_visit['Unnamed: 0']
    del tva_vod_serial_popular_duration['Unnamed: 0']
    del tva_vod_serial_popular_visit_middle['Unnamed: 0']
    del tva_vod_serial_popular_duration_middle['Unnamed: 0']
    del tva_vod_film_popular_visit['Unnamed: 0']
    del tva_vod_film_popular_duration['Unnamed: 0']
    
    print("edit of lenz vod")
    lenz_vod_serial_popular_visit.to_excel('busy/lenz_vod_serial_popular_visit.xlsx')
    lenz_vod_serial_popular_visit=pd.read_excel('busy/lenz_vod_serial_popular_visit.xlsx')
    lenz_vod_serial_popular_duration.to_excel('busy/lenz_vod_serial_popular_duration.xlsx')
    lenz_vod_serial_popular_duration=pd.read_excel('busy/lenz_vod_serial_popular_duration.xlsx')
    lenz_vod_serial_popular_visit_middle.to_excel('busy/lenz_vod_serial_popular_visit_middle.xlsx')
    lenz_vod_serial_popular_visit_middle=pd.read_excel('busy/lenz_vod_serial_popular_visit_middle.xlsx')
    lenz_vod_serial_popular_duration_middle.to_excel('busy/lenz_vod_serial_popular_duration_middle.xlsx')
    lenz_vod_serial_popular_duration_middle=pd.read_excel('busy/lenz_vod_serial_popular_duration_middle.xlsx')
    lenz_vod_film_popular_visit.to_excel('busy/lenz_vod_film_popular_visit.xlsx')
    lenz_vod_film_popular_visit=pd.read_excel('busy/lenz_vod_film_popular_visit.xlsx')
    lenz_vod_film_popular_duration.to_excel('busy/lenz_vod_film_popular_duration.xlsx')
    lenz_vod_film_popular_duration=pd.read_excel('busy/lenz_vod_film_popular_duration.xlsx')
    
    del lenz_vod_serial_popular_visit['Unnamed: 0']
    del lenz_vod_serial_popular_duration['Unnamed: 0']
    del lenz_vod_serial_popular_visit_middle['Unnamed: 0']
    del lenz_vod_serial_popular_duration_middle['Unnamed: 0']
    del lenz_vod_film_popular_visit['Unnamed: 0']
    del lenz_vod_film_popular_duration['Unnamed: 0']
    
#    print("edit of aio vod")
#    aio_vod_serial_popular_visit.to_excel('busy/aio_vod_serial_popular_visit.xlsx')
#    aio_vod_serial_popular_visit=pd.read_excel('busy/aio_vod_serial_popular_visit.xlsx')
#    aio_vod_serial_popular_duration.to_excel('busy/aio_vod_serial_popular_duration.xlsx')
#    aio_vod_serial_popular_duration=pd.read_excel('busy/aio_vod_serial_popular_duration.xlsx')
#    aio_vod_serial_popular_visit_middle.to_excel('busy/aio_vod_serial_popular_visit_middle.xlsx')
#    aio_vod_serial_popular_visit_middle=pd.read_excel('busy/aio_vod_serial_popular_visit_middle.xlsx')
#    aio_vod_serial_popular_duration_middle.to_excel('busy/aio_vod_serial_popular_duration_middle.xlsx')
#    aio_vod_serial_popular_duration_middle=pd.read_excel('busy/aio_vod_serial_popular_duration_middle.xlsx')
#    aio_vod_film_popular_visit.to_excel('busy/aio_vod_film_popular_visit.xlsx')
#    aio_vod_film_popular_visit=pd.read_excel('busy/aio_vod_film_popular_visit.xlsx')
#    aio_vod_film_popular_duration.to_excel('busy/aio_vod_film_popular_duration.xlsx')
#    aio_vod_film_popular_duration=pd.read_excel('busy/aio_vod_film_popular_duration.xlsx')
#    
#    del aio_vod_serial_popular_visit['Unnamed: 0']
#    del aio_vod_serial_popular_duration['Unnamed: 0']
#    del aio_vod_serial_popular_visit_middle['Unnamed: 0']
#    del aio_vod_serial_popular_duration_middle['Unnamed: 0']
#    del aio_vod_film_popular_visit['Unnamed: 0']
#    del aio_vod_film_popular_duration['Unnamed: 0']
    
    print("dataframe of vod popular content")
    
    vod_popular_content=pd.DataFrame()
    vod_popular_content=pd.concat([tva_vod_serial_popular_visit, tva_vod_serial_popular_duration,
                                                  tva_vod_serial_popular_visit_middle, tva_vod_serial_popular_duration_middle,
                                                  tva_vod_film_popular_visit, tva_vod_film_popular_duration,
                                                  lenz_vod_serial_popular_visit, lenz_vod_serial_popular_duration,
                                                  lenz_vod_serial_popular_visit_middle, lenz_vod_serial_popular_duration_middle,
                                                  lenz_vod_film_popular_visit, lenz_vod_film_popular_duration,
                                                  aio_vod_serial_popular_visit, aio_vod_serial_popular_duration,
                                                  aio_vod_serial_popular_visit_middle, aio_vod_serial_popular_duration_middle,
                                                  aio_vod_film_popular_visit, aio_vod_film_popular_duration,],axis=1)
    
    
    writer = pd.ExcelWriter('output/pedram/آمار VOD.xlsx', engine='xlsxwriter')
    vod_statistics_summary.to_excel(writer, 'خلاصه آمار VOD')
    vod_popular_content.to_excel(writer, 'محتواهای پربازدید')
    writer.save()

    writer = pd.ExcelWriter('output/output.sending.hard/آمار VOD.xlsx', engine='xlsxwriter')
    vod_statistics_summary.to_excel(writer, 'خلاصه آمار VOD')
    vod_popular_content.to_excel(writer, 'محتواهای پربازدید')
    writer.save()
    

    return vod_statistics_summary, vod_popular_content
    
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        