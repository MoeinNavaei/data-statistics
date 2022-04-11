
def content_type(sima):
    
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
    
    print("content type")
    
    Akhbar=sima.query("content_type == 'اخبار'")
    FilmSinamaei=sima.query("content_type == 'فیلم سینمایی'")
    Kodak=sima.query("content_type == 'کودک'")
    MajmoeTV=sima.query("content_type == 'مجموعه تلویزیونی'")
    Mostanad=sima.query("content_type == 'مستند'")
    Varzeshi=sima.query("content_type == 'ورزشی'")
    
    print("Akhbar")
    
    Akhbar_pivot=Akhbar.groupby(['نام برنامه','channel']).sum().reset_index()
    Akhbar_content=len(Akhbar_pivot['نام برنامه'])
    Akhbar_visit=Akhbar['تعداد بازدید'].sum()
    Akhbar_duration=Akhbar['مدت بازدید'].sum()
    Akhbar_pivot_channel=Akhbar.groupby(['channel']).sum().reset_index()
    Akhbar_channels=len(Akhbar_pivot_channel['channel'])
    Akhbar_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    Akhbar_popular_visit=Akhbar_pivot.iloc[0:10 , [0, 3]]
    Akhbar_popular_visit.to_excel('busy/Akhbar_popular_visit.xlsx')
    Akhbar_popular_visit=pd.read_excel('busy/Akhbar_popular_visit.xlsx')
    del Akhbar_popular_visit['Unnamed: 0']
    Akhbar_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    Akhbar_popular_duration=Akhbar_pivot.iloc[0:10 , [0, 2]]
    Akhbar_popular_duration.to_excel('busy/Akhbar_popular_duration.xlsx')
    Akhbar_popular_duration=pd.read_excel('busy/Akhbar_popular_duration.xlsx')
    del Akhbar_popular_duration['Unnamed: 0']
    
    Akhbar_statistics=pd.DataFrame()
    Akhbar_statistics={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد شبکه پخش کننده',],
           'statistics': [Akhbar_content, Akhbar_visit, Akhbar_duration, Akhbar_channels,],}
    Akhbar_statistics=pd.DataFrame(Akhbar_statistics, columns=['parameters', 'statistics'])
    Akhbar_statistics=Akhbar_statistics.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    Akhbar_data=pd.DataFrame()
    Akhbar_data=pd.concat([Akhbar_statistics, Akhbar_popular_visit, Akhbar_popular_duration], axis=1)    
    
    print("FilmSinamaei")
    
    FilmSinamaei_pivot=FilmSinamaei.groupby(['نام برنامه','channel']).sum().reset_index()
    FilmSinamaei_content=len(FilmSinamaei_pivot['نام برنامه'])
    FilmSinamaei_visit=FilmSinamaei['تعداد بازدید'].sum()
    FilmSinamaei_duration=FilmSinamaei['مدت بازدید'].sum()
    FilmSinamaei_pivot_channel=FilmSinamaei.groupby(['channel']).sum().reset_index()
    FilmSinamaei_channels=len(FilmSinamaei_pivot_channel['channel'])
    FilmSinamaei_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    FilmSinamaei_popular_visit=FilmSinamaei_pivot.iloc[0:10 , [0, 3]]
    FilmSinamaei_popular_visit.to_excel('busy/FilmSinamaei_popular_visit.xlsx')
    FilmSinamaei_popular_visit=pd.read_excel('busy/FilmSinamaei_popular_visit.xlsx')
    del FilmSinamaei_popular_visit['Unnamed: 0']
    FilmSinamaei_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    FilmSinamaei_popular_duration=FilmSinamaei_pivot.iloc[0:10 , [0, 2]]
    FilmSinamaei_popular_duration.to_excel('busy/FilmSinamaei_popular_duration.xlsx')
    FilmSinamaei_popular_duration=pd.read_excel('busy/FilmSinamaei_popular_duration.xlsx')
    del FilmSinamaei_popular_duration['Unnamed: 0']
    
    FilmSinamaei_statistics=pd.DataFrame()
    FilmSinamaei_statistics={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد شبکه پخش کننده',],
           'statistics': [FilmSinamaei_content, FilmSinamaei_visit, FilmSinamaei_duration, FilmSinamaei_channels,],}
    FilmSinamaei_statistics=pd.DataFrame(FilmSinamaei_statistics, columns=['parameters', 'statistics'])
    FilmSinamaei_statistics=FilmSinamaei_statistics.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    FilmSinamaei_data=pd.DataFrame()
    FilmSinamaei_data=pd.concat([FilmSinamaei_statistics, FilmSinamaei_popular_visit, FilmSinamaei_popular_duration], axis=1)    
    
    print("Kodak")
    
    Kodak_pivot=Kodak.groupby(['نام برنامه','channel']).sum().reset_index()
    Kodak_content=len(Kodak_pivot['نام برنامه'])
    Kodak_visit=Kodak['تعداد بازدید'].sum()
    Kodak_duration=Kodak['مدت بازدید'].sum()
    Kodak_pivot_channel=Kodak.groupby(['channel']).sum().reset_index()
    Kodak_channels=len(Kodak_pivot_channel['channel'])
    Kodak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    Kodak_popular_visit=Kodak_pivot.iloc[0:10 , [0, 3]]
    Kodak_popular_visit.to_excel('busy/Kodak_popular_visit.xlsx')
    Kodak_popular_visit=pd.read_excel('busy/Kodak_popular_visit.xlsx')
    del Kodak_popular_visit['Unnamed: 0']
    Kodak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    Kodak_popular_duration=Kodak_pivot.iloc[0:10 , [0, 2]]
    Kodak_popular_duration.to_excel('busy/Kodak_popular_duration.xlsx')
    Kodak_popular_duration=pd.read_excel('busy/Kodak_popular_duration.xlsx')
    del Kodak_popular_duration['Unnamed: 0']
    
    Kodak_statistics=pd.DataFrame()
    Kodak_statistics={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد شبکه پخش کننده',],
           'statistics': [Kodak_content, Kodak_visit, Kodak_duration, Kodak_channels,],}
    Kodak_statistics=pd.DataFrame(Kodak_statistics, columns=['parameters', 'statistics'])
    Kodak_statistics=Kodak_statistics.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    Kodak_data=pd.DataFrame()
    Kodak_data=pd.concat([Kodak_statistics, Kodak_popular_visit, Kodak_popular_duration], axis=1)    
    
    print("MajmoeTV")
    
    MajmoeTV_pivot=MajmoeTV.groupby(['نام برنامه','channel']).sum().reset_index()
    MajmoeTV_content=len(MajmoeTV_pivot['نام برنامه'])
    MajmoeTV_visit=MajmoeTV['تعداد بازدید'].sum()
    MajmoeTV_duration=MajmoeTV['مدت بازدید'].sum()
    MajmoeTV_pivot_channel=MajmoeTV.groupby(['channel']).sum().reset_index()
    MajmoeTV_channels=len(MajmoeTV_pivot_channel['channel'])
    MajmoeTV_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    MajmoeTV_popular_visit=MajmoeTV_pivot.iloc[0:10 , [0, 3]]
    MajmoeTV_popular_visit.to_excel('busy/MajmoeTV_popular_visit.xlsx')
    MajmoeTV_popular_visit=pd.read_excel('busy/MajmoeTV_popular_visit.xlsx')
    del MajmoeTV_popular_visit['Unnamed: 0']
    MajmoeTV_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    MajmoeTV_popular_duration=MajmoeTV_pivot.iloc[0:10 , [0, 2]]
    MajmoeTV_popular_duration.to_excel('busy/MajmoeTV_popular_duration.xlsx')
    MajmoeTV_popular_duration=pd.read_excel('busy/MajmoeTV_popular_duration.xlsx')
    del MajmoeTV_popular_duration['Unnamed: 0']
    
    MajmoeTV_statistics=pd.DataFrame()
    MajmoeTV_statistics={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد شبکه پخش کننده',],
           'statistics': [MajmoeTV_content, MajmoeTV_visit, MajmoeTV_duration, MajmoeTV_channels,],}
    MajmoeTV_statistics=pd.DataFrame(MajmoeTV_statistics, columns=['parameters', 'statistics'])
    MajmoeTV_statistics=MajmoeTV_statistics.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    MajmoeTV_data=pd.DataFrame()
    MajmoeTV_data=pd.concat([MajmoeTV_statistics, MajmoeTV_popular_visit, MajmoeTV_popular_duration], axis=1)
    
    print("Mostanad")
    
    Mostanad_pivot=Mostanad.groupby(['نام برنامه','channel']).sum().reset_index()
    Mostanad_content=len(Mostanad_pivot['نام برنامه'])
    Mostanad_visit=Mostanad['تعداد بازدید'].sum()
    Mostanad_duration=Mostanad['مدت بازدید'].sum()
    Mostanad_pivot_channel=Mostanad.groupby(['channel']).sum().reset_index()
    Mostanad_channels=len(Mostanad_pivot_channel['channel'])
    Mostanad_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    Mostanad_popular_visit=Mostanad_pivot.iloc[0:10 , [0, 3]]
    Mostanad_popular_visit.to_excel('busy/Mostanad_popular_visit.xlsx')
    Mostanad_popular_visit=pd.read_excel('busy/Mostanad_popular_visit.xlsx')
    del Mostanad_popular_visit['Unnamed: 0']
    Mostanad_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    Mostanad_popular_duration=Mostanad_pivot.iloc[0:10 , [0, 2]]
    Mostanad_popular_duration.to_excel('busy/Mostanad_popular_duration.xlsx')
    Mostanad_popular_duration=pd.read_excel('busy/Mostanad_popular_duration.xlsx')
    del Mostanad_popular_duration['Unnamed: 0']
    
    Mostanad_statistics=pd.DataFrame()
    Mostanad_statistics={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد شبکه پخش کننده',],
           'statistics': [Mostanad_content, Mostanad_visit, Mostanad_duration, Mostanad_channels,],}
    Mostanad_statistics=pd.DataFrame(Mostanad_statistics, columns=['parameters', 'statistics'])
    Mostanad_statistics=Mostanad_statistics.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    Mostanad_data=pd.DataFrame()
    Mostanad_data=pd.concat([Mostanad_statistics, Mostanad_popular_visit, Mostanad_popular_duration], axis=1)
    
    print("Varzeshi")
    
    Varzeshi_pivot=Varzeshi.groupby(['نام برنامه','channel']).sum().reset_index()
    Varzeshi_content=len(Varzeshi_pivot['نام برنامه'])
    Varzeshi_visit=Varzeshi['تعداد بازدید'].sum()
    Varzeshi_duration=Varzeshi['مدت بازدید'].sum()
    Varzeshi_pivot_channel=Varzeshi.groupby(['channel']).sum().reset_index()
    Varzeshi_channels=len(Varzeshi_pivot_channel['channel'])
    Varzeshi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    Varzeshi_popular_visit=Varzeshi_pivot.iloc[0:10 , [0, 3]]
    Varzeshi_popular_visit.to_excel('busy/Varzeshi_popular_visit.xlsx')
    Varzeshi_popular_visit=pd.read_excel('busy/Varzeshi_popular_visit.xlsx')
    del Varzeshi_popular_visit['Unnamed: 0']
    Varzeshi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    Varzeshi_popular_duration=Varzeshi_pivot.iloc[0:10 , [0, 2]]
    Varzeshi_popular_duration.to_excel('busy/Varzeshi_popular_duration.xlsx')
    Varzeshi_popular_duration=pd.read_excel('busy/Varzeshi_popular_duration.xlsx')
    del Varzeshi_popular_duration['Unnamed: 0']
    
    Varzeshi_statistics=pd.DataFrame()
    Varzeshi_statistics={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد شبکه پخش کننده',],
           'statistics': [Varzeshi_content, Varzeshi_visit, Varzeshi_duration, Varzeshi_channels,],}
    Varzeshi_statistics=pd.DataFrame(Varzeshi_statistics, columns=['parameters', 'statistics'])
    Varzeshi_statistics=Varzeshi_statistics.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    Varzeshi_data=pd.DataFrame()
    Varzeshi_data=pd.concat([Varzeshi_statistics, Varzeshi_popular_visit, Varzeshi_popular_duration], axis=1)
    
    
    writer = pd.ExcelWriter('output/zomorrodi/انواع محتوا در سیما.xlsx', engine='xlsxwriter')
    Akhbar_data.to_excel(writer, 'اخبار', index = False)
    FilmSinamaei_data.to_excel(writer, 'فیلم سینمایی', index = False)
    Kodak_data.to_excel(writer, 'کودک', index = False)
    MajmoeTV_data.to_excel(writer, 'مجموعه تلویزیونی', index = False)
    Mostanad_data.to_excel(writer, 'مستند', index = False)
    Varzeshi_data.to_excel(writer, 'ورزشی', index = False)
    writer.save()
    
    writer = pd.ExcelWriter('output/output.sending.hard/انواع محتوا در سیما.xlsx', engine='xlsxwriter')
    Akhbar_data.to_excel(writer, 'اخبار', index = False)
    FilmSinamaei_data.to_excel(writer, 'فیلم سینمایی', index = False)
    Kodak_data.to_excel(writer, 'کودک', index = False)
    MajmoeTV_data.to_excel(writer, 'مجموعه تلویزیونی', index = False)
    Mostanad_data.to_excel(writer, 'مستند', index = False)
    Varzeshi_data.to_excel(writer, 'ورزشی', index = False)
    writer.save()
    
    
    return Akhbar_statistics, Akhbar_popular_visit, Akhbar_popular_duration, \
           FilmSinamaei_statistics, FilmSinamaei_popular_visit, FilmSinamaei_popular_duration, \
           Kodak_statistics, Kodak_popular_visit, Kodak_popular_duration, \
           MajmoeTV_statistics, MajmoeTV_popular_visit, MajmoeTV_popular_duration, \
           Mostanad_statistics, Mostanad_popular_visit, Mostanad_popular_duration, \
           Varzeshi_statistics, Varzeshi_popular_visit, Varzeshi_popular_duration

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    