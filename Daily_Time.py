
def Daily_Time(all_data, all_data_Time):
    
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
    import itertools
    from datetime import datetime as dt
    from persiantools.jdatetime import JalaliDate
    import datetime
    import jalali
    import jdatetime
    
    print("Time visit") 
    
    del all_data_Time['channel']
    del all_data_Time['نام برنامه اولیه']
    del all_data_Time['نام برنامه']
    del all_data_Time['تاریخ شروع']
    del all_data_Time['تاریخ پایان']
    del all_data_Time['میانگین']
    del all_data_Time['operator']
    del all_data_Time['تاریخ']
    del all_data_Time['ردیف']
    del all_data_Time['content_type']
    del all_data_Time['tag']
    
    # all_data_Time['ساعت'] = all_data_Time['ساعت'].str.strip()   # remove space
    all_data_Time['ساعت'].replace('', 'no', inplace=True)  # write "no" instead of nan
    all_data_Time=all_data_Time[all_data_Time['ساعت'] != 'no']         # remove no
    
    all_data_Time_statistics_sima = all_data_Time.copy()
    all_data_Time_statistics_sima = all_data_Time_statistics_sima.query("type == 'سراسری'")
    all_data_Time_statistics_sima=all_data_Time_statistics_sima.groupby(['ساعت']).sum().reset_index()
    
    all_data_Time_statistics_radio = all_data_Time.copy()
    all_data_Time_statistics_radio = all_data_Time_statistics_radio.query("type == 'رادیویی'")
    all_data_Time_statistics_radio=all_data_Time_statistics_radio.groupby(['ساعت']).sum().reset_index()
    
    all_data_Time_statistics_ostani = all_data_Time.copy()
    all_data_Time_statistics_ostani = all_data_Time_statistics_ostani.query("type == 'استانی'")
    all_data_Time_statistics_ostani=all_data_Time_statistics_ostani.groupby(['ساعت']).sum().reset_index()
    
    all_data_Time_statistics_ekhtesasi = all_data_Time.copy()
    all_data_Time_statistics_ekhtesasi = all_data_Time_statistics_ekhtesasi.query("type == 'اختصاصی'")
    all_data_Time_statistics_ekhtesasi=all_data_Time_statistics_ekhtesasi.groupby(['ساعت']).sum().reset_index()
    
    all_data_Time_statistics_boronmarzi = all_data_Time.copy()
    all_data_Time_statistics_boronmarzi = all_data_Time_statistics_boronmarzi.query("type == 'برون مرزی'")
    all_data_Time_statistics_boronmarzi=all_data_Time_statistics_boronmarzi.groupby(['ساعت']).sum().reset_index()
    
    
    print("Daily visit")
    
    all_data_Daily_statistics=all_data.copy()
    all_data_Daily_statistics=all_data_Daily_statistics.query("operator != 'تلوبیون'")
    all_data_Daily_statistics=all_data_Daily_statistics.query("operator != 'سپهر'")
    all_data_Daily_statistics=all_data_Daily_statistics.query("operator != 'سایت شبکه ها'")
    all_data_Daily_statistics.insert(15, 'day', '')
    all_data_Daily_statistics.insert(16, 'jalali', '')
    
    all_data_Daily_statistics['تاریخ شروع'] = all_data_Daily_statistics['تاریخ شروع'].astype(str)
    all_data_Daily_statistics=all_data_Daily_statistics.rename(columns={"تاریخ شروع":"start_date"})
    date2 = pd.to_datetime(all_data_Daily_statistics.start_date, errors='coerce')
    all_data_Daily_statistics = all_data_Daily_statistics.assign(s_date=date2.dt.date)
    
    all_data_Daily_statistics['s_date'] = all_data_Daily_statistics['s_date'].astype(str)
    all_data_Daily_statistics['day'] = all_data_Daily_statistics['s_date'].str[8:10]
    all_data_Daily_statistics['month'] = all_data_Daily_statistics['s_date'].str[5:7]
    all_data_Daily_statistics['year'] = all_data_Daily_statistics['s_date'].str[0:4]
    all_data_Daily_statistics['day'] = all_data_Daily_statistics['day'].astype(int)
    all_data_Daily_statistics['month'] = all_data_Daily_statistics['month'].astype(int)
    all_data_Daily_statistics['year'] = all_data_Daily_statistics['year'].astype(int)
    all_data_Daily_statistics1 = all_data_Daily_statistics['day']
    all_data_Daily_statistics2 = all_data_Daily_statistics['month']
    all_data_Daily_statistics3 = all_data_Daily_statistics['year']

    for i in range(0, len(all_data_Daily_statistics)):
        print("i_Date: ", i)
        try:
            date_jalali = JalaliDate(datetime.date(all_data_Daily_statistics3[i], all_data_Daily_statistics2[i], all_data_Daily_statistics1[i]))
            all_data_Daily_statistics.loc[i, 'jalali'] = date_jalali
        except: pass
    
    del all_data_Daily_statistics['year']
    del all_data_Daily_statistics['month']   
    del all_data_Daily_statistics['day']   
    del all_data_Daily_statistics['s_date']   
    all_data_Daily_statistics['jalali'] = all_data_Daily_statistics['jalali'].astype(str)
    all_data_Daily_statistics['day'] = all_data_Daily_statistics['jalali'].str[8:10]
       
    del all_data_Daily_statistics['channel']
    del all_data_Daily_statistics['نام برنامه اولیه']
    del all_data_Daily_statistics['نام برنامه']
    del all_data_Daily_statistics['start_date']
    del all_data_Daily_statistics['تاریخ پایان']
    del all_data_Daily_statistics['میانگین']
    del all_data_Daily_statistics['operator']
    del all_data_Daily_statistics['ساعت']
    del all_data_Daily_statistics['تاریخ']
    del all_data_Daily_statistics['ردیف']
    del all_data_Daily_statistics['content_type']
    del all_data_Daily_statistics['tag']
    del all_data_Daily_statistics['jalali']
#    all_data_Daily_statistics.to_excel('all_data_Daily_statistics1.xlsx', index=False)
    all_data_Daily_statistics['day'] = all_data_Daily_statistics['day'].str.strip()   # remove space
    all_data_Daily_statistics['day'].replace('', 'NO', inplace=True)  
#    all_data_Daily_statistics=all_data_Daily_statistics[all_data_Daily_statistics['day'] != 'no']         # remove no
    all_data_Daily_statistics = all_data_Daily_statistics[~all_data_Daily_statistics.day.str.contains("NO")]
#    all_data_Daily_statistics.to_excel('all_data_Daily_statistics2.xlsx', index=False)
    all_data_Daily_statistics_sima = all_data_Daily_statistics.query("type == 'سراسری'")
    all_data_Daily_statistics_sima=all_data_Daily_statistics_sima.groupby(['day']).sum().reset_index()
    
    all_data_Daily_statistics_radio = all_data_Daily_statistics.query("type == 'رادیویی'")
    all_data_Daily_statistics_radio=all_data_Daily_statistics_radio.groupby(['day']).sum().reset_index()
    
    all_data_Daily_statistics_ostani = all_data_Daily_statistics.query("type == 'استانی'")
    all_data_Daily_statistics_ostani=all_data_Daily_statistics_ostani.groupby(['day']).sum().reset_index()
    
    all_data_Daily_statistics_ekhtesasi = all_data_Daily_statistics.query("type == 'اختصاصی'")
    all_data_Daily_statistics_ekhtesasi=all_data_Daily_statistics_ekhtesasi.groupby(['day']).sum().reset_index()
    
    all_data_Daily_statistics_boronmarzi = all_data_Daily_statistics.query("type == 'برون مرزی'")
    all_data_Daily_statistics_boronmarzi=all_data_Daily_statistics_boronmarzi.groupby(['day']).sum().reset_index()
   
        
    writer = pd.ExcelWriter('output/zomorrodi/آمار ساعتی و روزانه.xlsx', engine='xlsxwriter')
    all_data_Time_statistics_sima.to_excel(writer, 'آمار ساعتی سیما', index = False)
    all_data_Daily_statistics_sima.to_excel(writer, 'آمار روزانه سیما', index = False)
    all_data_Time_statistics_radio.to_excel(writer, 'آمار ساعتی رادیو', index = False)
    all_data_Daily_statistics_radio.to_excel(writer, 'آمار روزانه رادیو', index = False)
    all_data_Time_statistics_ostani.to_excel(writer, 'آمار ساعتی استانی', index = False)
    all_data_Daily_statistics_ostani.to_excel(writer, 'آمار روزانه استانی', index = False)
    all_data_Time_statistics_ekhtesasi.to_excel(writer, 'آمار ساعتی اختصاصی', index = False)
    all_data_Daily_statistics_ekhtesasi.to_excel(writer, 'آمار روزانه اختصاصی', index = False)
    all_data_Time_statistics_boronmarzi.to_excel(writer, 'آمار ساعتی برونمرزی', index = False)
    all_data_Daily_statistics_boronmarzi.to_excel(writer, 'آمار روزانه برونمرزی', index = False)
    writer.save()
    
    writer = pd.ExcelWriter('output/output.sending.hard/آمار ساعتی و روزانه.xlsx', engine='xlsxwriter')
    all_data_Time_statistics_sima.to_excel(writer, 'آمار ساعتی سیما', index = False)
    all_data_Daily_statistics_sima.to_excel(writer, 'آمار روزانه سیما', index = False)
    all_data_Time_statistics_radio.to_excel(writer, 'آمار ساعتی رادیو', index = False)
    all_data_Daily_statistics_radio.to_excel(writer, 'آمار روزانه رادیو', index = False)
    all_data_Time_statistics_ostani.to_excel(writer, 'آمار ساعتی استانی', index = False)
    all_data_Daily_statistics_ostani.to_excel(writer, 'آمار روزانه استانی', index = False)
    all_data_Time_statistics_ekhtesasi.to_excel(writer, 'آمار ساعتی اختصاصی', index = False)
    all_data_Daily_statistics_ekhtesasi.to_excel(writer, 'آمار روزانه اختصاصی', index = False)
    all_data_Time_statistics_boronmarzi.to_excel(writer, 'آمار ساعتی برونمرزی', index = False)
    all_data_Daily_statistics_boronmarzi.to_excel(writer, 'آمار روزانه برونمرزی', index = False)
    writer.save()
    
    return all_data_Time_statistics_sima, all_data_Daily_statistics_sima, \
    all_data_Time_statistics_radio, all_data_Daily_statistics_radio, \
    all_data_Time_statistics_ostani, all_data_Daily_statistics_ostani, \
    all_data_Time_statistics_ekhtesasi, all_data_Daily_statistics_ekhtesasi, \
    all_data_Time_statistics_boronmarzi, all_data_Daily_statistics_boronmarzi



    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    