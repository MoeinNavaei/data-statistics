
import xlsxwriter  
import pandas as pd
#from pandas import DataFrame
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import arabic_reshaper
from bidi.algorithm import get_display
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

################################# input data ########################################
print("get data")
all_data=pd.read_excel('total_dey.xlsx')
all_data=all_data.rename(columns={"نوع":"type"})
all_data=all_data.rename(columns={"اپراتور":"operator"})
all_data=all_data.rename(columns={"نام شبکه":"channel"})

vod_lenz=pd.read_csv('lenz_vod\dey1399.csv')
vod_tva=pd.read_excel('Tva_vod\dey1399.xlsx', sheet_name='Videos')
all_data_Time=all_data.copy()
all_data_Time=all_data_Time.query("operator != 'تلوبیون'")
################################# summary statistics ########################################
print("start summary")
all_visit=all_data['تعداد بازدید'].sum()
all_duration=all_data['مدت بازدید'].sum()
all_duration=round(all_duration*60, 0)

all_data_pivot=all_data.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()

sima=all_data.query("type == 'سراسری'")
radio=all_data.query("type == 'رادیویی'")
ostani=all_data.query("type == 'استانی'")
ekhtesasi=all_data.query("type == 'اختصاصی'")

lenz=all_data.query("operator == 'لنز'")
tva=all_data.query("operator == 'تیوا'")
televebion=all_data.query("operator == 'تلوبیون'")

print("sima summary")
sima_all_visit=sima['تعداد بازدید'].sum()
sima_all_duration=sima['مدت بازدید'].sum()
sima_all_duration=round(sima_all_duration*60, 0)
sima_content=sima.copy()
sima_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
sima_all_content=len(sima_content)
sima_channel=sima.copy()
sima_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
sima_all_channel=len(sima_channel)

print("radio summary")
radio_all_visit=radio['تعداد بازدید'].sum()
radio_all_duration=radio['مدت بازدید'].sum()
radio_all_duration=round(radio_all_duration*60, 0)
radio_content=radio.copy()
radio_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_all_content=len(radio_content)
radio_channel=radio.copy()
radio_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
radio_all_channel=len(radio_channel)

print("ostani summary")
ostani_all_visit=ostani['تعداد بازدید'].sum()
ostani_all_duration=ostani['مدت بازدید'].sum()
ostani_all_duration=round(ostani_all_duration*60, 0)
ostani_content=ostani.copy()
ostani_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_all_content=len(ostani_content)
ostani_channel=ostani.copy()
ostani_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
ostani_all_channel=len(ostani_channel)

print("ekhtesasi summary")
ekhtesasi_all_visit=ekhtesasi['تعداد بازدید'].sum()
ekhtesasi_all_duration=ekhtesasi['مدت بازدید'].sum()
ekhtesasi_all_duration=round(ekhtesasi_all_duration*60, 0)
ekhtesasi_content=ekhtesasi.copy()
ekhtesasi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_all_content=len(ekhtesasi_content)
ekhtesasi_channel=ekhtesasi.copy()
ekhtesasi_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
ekhtesasi_all_channel=len(ekhtesasi_channel)

print("lenz summary")
lenz_all_visit=lenz['تعداد بازدید'].sum()
lenz_all_duration=lenz['مدت بازدید'].sum()
lenz_all_duration=round(lenz_all_duration*60, 0)
lenz_channel=lenz.copy()
lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
lenz_all_channel=len(lenz_channel)

print("tva summary")
tva_all_visit=tva['تعداد بازدید'].sum()
tva_all_duration=tva['مدت بازدید'].sum()
tva_all_duration=round(tva_all_duration*60, 0)
tva_channel=tva.copy()
tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
tva_all_channel=len(tva_channel)

print("televebion summary")
televebion_all_visit=televebion['تعداد بازدید'].sum()
televebion_all_duration=televebion['مدت بازدید'].sum()
televebion_all_duration=round(televebion_all_duration*60, 0)
televebion_channel=televebion.copy()
televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
televebion_all_channel=len(televebion_channel)

print("sima_operator summary")
sima_lenz=sima.query("operator == 'لنز'")
sima_lenz_all_visit=sima_lenz['تعداد بازدید'].sum()
sima_lenz_all_duration=sima_lenz['مدت بازدید'].sum()
sima_lenz_all_duration=round(sima_lenz_all_duration*60, 0)
sima_lenz_channel=sima_lenz.copy()
sima_lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
sima_lenz_all_channel=len(sima_lenz_channel)

sima_tva=sima.query("operator == 'تیوا'")
sima_tva_all_visit=sima_tva['تعداد بازدید'].sum()
sima_tva_all_duration=sima_tva['مدت بازدید'].sum()
sima_tva_all_duration=round(sima_tva_all_duration*60, 0)
sima_tva_channel=sima_tva.copy()
sima_tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
sima_tva_all_channel=len(sima_tva_channel)

sima_televebion=sima.query("operator == 'تلوبیون'")
sima_televebion_all_visit=sima_televebion['تعداد بازدید'].sum()
sima_televebion_all_duration=sima_televebion['مدت بازدید'].sum()
sima_televebion_all_duration=round(sima_televebion_all_duration*60, 0)
sima_televebion_channel=sima_televebion.copy()
sima_televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
sima_televebion_all_channel=len(sima_televebion_channel)

print("radio_operator summary")
radio_lenz=radio.query("operator == 'لنز'")
radio_lenz_all_visit=radio_lenz['تعداد بازدید'].sum()
radio_lenz_all_duration=radio_lenz['مدت بازدید'].sum()
radio_lenz_all_duration=round(radio_lenz_all_duration*60, 0)
radio_lenz_channel=radio_lenz.copy()
radio_lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
radio_lenz_all_channel=len(radio_lenz_channel)

radio_tva=radio.query("operator == 'تیوا'")
radio_tva_all_visit=radio_tva['تعداد بازدید'].sum()
radio_tva_all_duration=radio_tva['مدت بازدید'].sum()
radio_tva_all_duration=round(radio_tva_all_duration*60, 0)
radio_tva_channel=radio_tva.copy()
radio_tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
radio_tva_all_channel=len(radio_tva_channel)

radio_televebion=radio.query("operator == 'تلوبیون'")
radio_televebion_all_visit=radio_televebion['تعداد بازدید'].sum()
radio_televebion_all_duration=radio_televebion['مدت بازدید'].sum()
radio_televebion_all_duration=round(radio_televebion_all_duration*60, 0)
radio_televebion_channel=radio_televebion.copy()
radio_televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
radio_televebion_all_channel=len(radio_televebion_channel)

print("ostani_operator summary")
ostani_lenz=ostani.query("operator == 'لنز'")
ostani_lenz_all_visit=ostani_lenz['تعداد بازدید'].sum()
ostani_lenz_all_duration=ostani_lenz['مدت بازدید'].sum()
ostani_lenz_all_duration=round(ostani_lenz_all_duration*60, 0)
ostani_lenz_channel=ostani_lenz.copy()
ostani_lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
ostani_lenz_all_channel=len(ostani_lenz_channel)

ostani_tva=ostani.query("operator == 'تیوا'")
ostani_tva_all_visit=ostani_tva['تعداد بازدید'].sum()
ostani_tva_all_duration=ostani_tva['مدت بازدید'].sum()
ostani_tva_all_duration=round(ostani_tva_all_duration*60, 0)
ostani_tva_channel=ostani_tva.copy()
ostani_tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
ostani_tva_all_channel=len(ostani_tva_channel)

ostani_televebion=ostani.query("operator == 'تلوبیون'")
ostani_televebion_all_visit=ostani_televebion['تعداد بازدید'].sum()
ostani_televebion_all_duration=ostani_televebion['مدت بازدید'].sum()
ostani_televebion_all_duration=round(ostani_televebion_all_duration*60, 0)
ostani_televebion_channel=ostani_televebion.copy()
ostani_televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
ostani_televebion_all_channel=len(ostani_televebion_channel)

print("ekhtesasi_operator summary")
ekhtesasi_lenz=ekhtesasi.query("operator == 'لنز'")
ekhtesasi_lenz_all_visit=ekhtesasi_lenz['تعداد بازدید'].sum()
ekhtesasi_lenz_all_duration=ekhtesasi_lenz['مدت بازدید'].sum()
ekhtesasi_lenz_all_duration=round(ekhtesasi_lenz_all_duration*60, 0)
ekhtesasi_lenz_channel=ekhtesasi_lenz.copy()
ekhtesasi_lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
ekhtesasi_lenz_all_channel=len(ekhtesasi_lenz_channel)

ekhtesasi_tva=ekhtesasi.query("operator == 'تیوا'")
ekhtesasi_tva_all_visit=ekhtesasi_tva['تعداد بازدید'].sum()
ekhtesasi_tva_all_duration=ekhtesasi_tva['مدت بازدید'].sum()
ekhtesasi_tva_all_duration=ekhtesasi_tva_all_duration*60
ekhtesasi_tva_channel=ekhtesasi_tva.copy()
ekhtesasi_tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
ekhtesasi_tva_all_channel=len(ekhtesasi_tva_channel)

ekhtesasi_televebion=ekhtesasi.query("operator == 'تلوبیون'")
ekhtesasi_televebion_all_visit=ekhtesasi_televebion['تعداد بازدید'].sum()
ekhtesasi_televebion_all_duration=ekhtesasi_televebion['مدت بازدید'].sum()
ekhtesasi_televebion_all_duration=round(ekhtesasi_televebion_all_duration*60, 0)
ekhtesasi_televebion_channel=ekhtesasi_televebion.copy()
ekhtesasi_televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
ekhtesasi_televebion_all_channel=len(ekhtesasi_televebion_channel)

print("dataframe summary")
data_summary_service={'field': ['تعداد محتوا', 'تعداد بازدید', 'مدت بازدید (به دقیقه)','تعداد شبکه'],
       'sima': [sima_all_content, sima_all_visit, sima_all_duration, sima_all_channel],
       'radio': [radio_all_content, radio_all_visit, radio_all_duration, radio_all_channel],
       'ostani': [ostani_all_content, ostani_all_visit, ostani_all_duration, ostani_all_channel],
       'ekhtesasi': [ekhtesasi_all_content, ekhtesasi_all_visit, ekhtesasi_all_duration, ekhtesasi_all_channel],}
data_summary_service=pd.DataFrame(data_summary_service, columns=['field', 'sima', 'radio', 'ostani', 'ekhtesasi'])
data_summary_service=data_summary_service.rename(columns={'field': 'حوزه', 'sima': 'سیما', 'radio': 'تیوا', 'رادیویی': 'استانی', 'ekhtesasi': 'اختصاصی'})

data_summary_operator={'operator': ['تعداد شبکه', 'تعداد بازدید', 'مدت بازدید (به دقیقه)'],
       'lenz': [lenz_all_channel, lenz_all_visit, lenz_all_duration],
       'tva': [tva_all_channel, tva_all_visit, tva_all_duration],
       'televebion': [televebion_all_channel, televebion_all_visit, televebion_all_duration],}
data_summary_operator=pd.DataFrame(data_summary_operator, columns=['operator', 'lenz', 'tva', 'televebion'])
data_summary_operator=data_summary_operator.rename(columns={'operator': 'اپراتور', 'lenz': 'لنز', 'tva': 'تیوا', 'televebion': 'تلوبیون'})

data_summary_service_operator={'parameters': ['تعداد شبکه ', 'تعداد بازدید', 'مدت بازدید (به دقیقه)'],
       'sima_lenz': [sima_lenz_all_channel, sima_lenz_all_visit, sima_lenz_all_duration],
       'sima_tva': [sima_tva_all_channel, sima_tva_all_visit, sima_tva_all_duration],
       'sima_televebion': [sima_televebion_all_channel, sima_televebion_all_visit, sima_televebion_all_duration],
       'radio_lenz': [radio_lenz_all_channel, radio_lenz_all_visit, radio_lenz_all_duration],
       'radio_tva': [radio_tva_all_channel, radio_tva_all_visit, radio_tva_all_duration],
       'radio_televebion': [radio_televebion_all_channel, radio_televebion_all_visit, radio_televebion_all_duration],
       'ostani_lenz': [ostani_lenz_all_channel, ostani_lenz_all_visit, ostani_lenz_all_duration],
       'ostani_tva': [ostani_tva_all_channel, ostani_tva_all_visit, ostani_tva_all_duration],
       'ostani_televebion': [ostani_televebion_all_channel, ostani_televebion_all_visit, ostani_televebion_all_duration],
       'ekhtesasi_lenz': [ekhtesasi_lenz_all_channel, ekhtesasi_lenz_all_visit, ekhtesasi_lenz_all_duration],
       'ekhtesasi_tva': [ekhtesasi_tva_all_channel, ekhtesasi_tva_all_visit, ekhtesasi_tva_all_duration],
       'ekhtesasi_televebion': [ekhtesasi_televebion_all_channel, ekhtesasi_televebion_all_visit, ekhtesasi_televebion_all_duration],}
data_summary_service_operator=pd.DataFrame(data_summary_service_operator, columns=['parameters', 'sima_lenz', 'sima_tva','sima_televebion',
                                                                                   'radio_lenz', 'radio_tva','radio_televebion',
                                                                                   'ostani_lenz', 'ostani_tva','ostani_televebion',
                                                                                   'ekhtesasi_lenz', 'ekhtesasi_tva','ekhtesasi_televebion',])

data_summary_service_operator=data_summary_service_operator.rename(columns={'parameters': 'پارامترها', 
                                                                            'sima_lenz': 'آمار', 'sima_tva': 'آمار', 'sima_televebion': 'آمار',
                                                                            'radio_lenz': 'آمار', 'radio_tva': 'آمار', 'radio_televebion': 'آمار',
                                                                            'ostani_lenz': 'آمار', 'ostani_tva': 'آمار', 'ostani_televebion': 'آمار',
                                                                            'ekhtesasi_lenz': 'آمار', 'ekhtesasi_tva': 'آمار', 'ekhtesasi_televebion': 'آمار'})

writer = pd.ExcelWriter('output/خلاصه آمار.xlsx', engine='xlsxwriter')
data_summary_service.to_excel(writer, 'آمار بازدید شبکه های سیما')
data_summary_operator.to_excel(writer, 'زمان بازدید شبکه ها (به دقیقه)')
data_summary_service_operator.to_excel(writer, 'آمار اپراتورها')
writer.save()

print("End summary")

################################# SIMA ########################################
print("start sima")

sima_all=sima.copy()
sima_all_pivot=sima_all.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
sima_all_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
sima_all_popular_visit=sima_all_pivot.iloc[0:10 , [0, 4]]
sima_all_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
sima_all_popular_duration=sima_all_pivot.iloc[0:10 , [0, 3]]

print("shabake_1")
shabake_1=sima.query("channel == 'شبکه 1'")
shabake_1_visit=shabake_1['تعداد بازدید'].sum()
shabake_1_duration=shabake_1['مدت بازدید'].sum()
shabake_1_duration=round(shabake_1_duration*60, 0)
shabake_1_content=shabake_1.copy()
shabake_1_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_1_content=len(shabake_1_content)
shabake_1_pivot=shabake_1.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_1_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_1_popular_visit=shabake_1_pivot.iloc[0:10 , [0, 5]]
shabake_1_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_1_popular_duration=shabake_1_pivot.iloc[0:10 , [0, 4]]

shabake_1_popular_visit = shabake_1_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 1', 'نام برنامه': 'محتواهای پربازدید شبکه 1'})
shabake_1_popular_duration = shabake_1_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 1 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 1'})

shabake_1_Time_visit=all_data_Time.query("channel == 'شبکه 1'")
shabake_1_Time_visit=shabake_1_Time_visit.copy()
shabake_1_Time_visit=shabake_1_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_1_Time_visit['میانگین']
del shabake_1_Time_visit['تاریخ']
del shabake_1_Time_visit['ردیف']
del shabake_1_Time_visit['tag']
shabake_1_Time_visit = shabake_1_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 1', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 1'})

print("shabake_2")
shabake_2=sima.query("channel == 'شبکه 2'")
shabake_2_visit=shabake_2['تعداد بازدید'].sum()
shabake_2_duration=shabake_2['مدت بازدید'].sum()
shabake_2_duration=round(shabake_2_duration*60, 0)
shabake_2_content=shabake_2.copy()
shabake_2_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_2_content=len(shabake_2_content)
shabake_2_pivot=shabake_2.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_2_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_2_popular_visit=shabake_2_pivot.iloc[0:10 , [0, 5]]
shabake_2_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_2_popular_duration=shabake_2_pivot.iloc[0:10 , [0, 4]]

shabake_2_popular_visit = shabake_2_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 2', 'نام برنامه': 'محتواهای پربازدید شبکه 2'})
shabake_2_popular_duration = shabake_2_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 2 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 2'})

shabake_2_Time_visit=all_data_Time.query("channel == 'شبکه 2'")
shabake_2_Time_visit=shabake_2_Time_visit.copy()
shabake_2_Time_visit=shabake_2_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_2_Time_visit['میانگین']
del shabake_2_Time_visit['تاریخ']
del shabake_2_Time_visit['ردیف']
del shabake_2_Time_visit['tag']
shabake_2_Time_visit = shabake_2_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 2', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 2'})

print("shabake_3")
shabake_3=sima.query("channel == 'شبکه 3'")
shabake_3_visit=shabake_3['تعداد بازدید'].sum()
shabake_3_duration=shabake_3['مدت بازدید'].sum()
shabake_3_duration=round(shabake_3_duration*60, 0)
shabake_3_content=shabake_3.copy()
shabake_3_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_3_content=len(shabake_3_content)
shabake_3_pivot=shabake_3.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_3_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_3_popular_visit=shabake_3_pivot.iloc[0:10 , [0, 5]]
shabake_3_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_3_popular_duration=shabake_3_pivot.iloc[0:10 , [0, 4]]

shabake_3_popular_visit = shabake_3_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 3', 'نام برنامه': 'محتواهای پربازدید شبکه 3'})
shabake_3_popular_duration = shabake_3_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 3 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 3'})

shabake_3_Time_visit=all_data_Time.query("channel == 'شبکه 3'")
shabake_3_Time_visit=shabake_3_Time_visit.copy()
shabake_3_Time_visit=shabake_3_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_3_Time_visit['میانگین']
del shabake_3_Time_visit['تاریخ']
del shabake_3_Time_visit['ردیف']
del shabake_3_Time_visit['tag']
shabake_3_Time_visit = shabake_3_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 3', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 3'})

print("shabake_4")
shabake_4=sima.query("channel == 'شبکه 4'")
shabake_4_visit=shabake_4['تعداد بازدید'].sum()
shabake_4_duration=shabake_4['مدت بازدید'].sum()
shabake_4_duration=round(shabake_4_duration*60, 0)
shabake_4_content=shabake_4.copy()
shabake_4_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_4_content=len(shabake_4_content)
shabake_4_pivot=shabake_4.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_4_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_4_popular_visit=shabake_4_pivot.iloc[0:10 , [0, 5]]
shabake_4_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_4_popular_duration=shabake_4_pivot.iloc[0:10 , [0, 4]]

shabake_4_popular_visit = shabake_4_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 4', 'نام برنامه': 'محتواهای پربازدید شبکه 4'})
shabake_4_popular_duration = shabake_4_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 4 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 4'})

shabake_4_Time_visit=all_data_Time.query("channel == 'شبکه 4'")
shabake_4_Time_visit=shabake_4_Time_visit.copy()
shabake_4_Time_visit=shabake_4_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_4_Time_visit['میانگین']
del shabake_4_Time_visit['تاریخ']
del shabake_4_Time_visit['ردیف']
del shabake_4_Time_visit['tag']
shabake_4_Time_visit = shabake_4_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 4', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 4'})

print("shabake_5")
shabake_5=sima.query("channel == 'شبکه 5'")
shabake_5_visit=shabake_5['تعداد بازدید'].sum()
shabake_5_duration=shabake_5['مدت بازدید'].sum()
shabake_5_duration=round(shabake_5_duration*60, 0)
shabake_5_content=shabake_5.copy()
shabake_5_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_5_content=len(shabake_5_content)
shabake_5_pivot=shabake_5.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_5_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_5_popular_visit=shabake_5_pivot.iloc[0:10 , [0, 5]]
shabake_5_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_5_popular_duration=shabake_5_pivot.iloc[0:10 , [0, 4]]

shabake_5_popular_visit = shabake_5_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 5', 'نام برنامه': 'محتواهای پربازدید شبکه 5'})
shabake_5_popular_duration = shabake_5_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 5 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 5'})

shabake_5_Time_visit=all_data_Time.query("channel == 'شبکه 5'")
shabake_5_Time_visit=shabake_5_Time_visit.copy()
shabake_5_Time_visit=shabake_5_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_5_Time_visit['میانگین']
del shabake_5_Time_visit['تاریخ']
del shabake_5_Time_visit['ردیف']
del shabake_5_Time_visit['tag']
shabake_5_Time_visit = shabake_5_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 5', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 5'})

print("shabake_khabar")
shabake_khabar=sima.query("channel == 'خبر'")
shabake_khabar_visit=shabake_khabar['تعداد بازدید'].sum()
shabake_khabar_duration=shabake_khabar['مدت بازدید'].sum()
shabake_khabar_duration=round(shabake_khabar_duration*60, 0)
shabake_khabar_content=shabake_khabar.copy()
shabake_khabar_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_khabar_content=len(shabake_khabar_content)
shabake_khabar_pivot=shabake_khabar.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_khabar_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_khabar_popular_visit=shabake_khabar_pivot.iloc[0:10 , [0, 5]]
shabake_khabar_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_khabar_popular_duration=shabake_khabar_pivot.iloc[0:10 , [0, 4]]

shabake_khabar_popular_visit = shabake_khabar_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه خبر', 'نام برنامه': 'محتواهای پربازدید شبکه خبر'})
shabake_khabar_popular_duration = shabake_khabar_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه خبر (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه خبر'})

shabake_khabar_Time_visit=all_data_Time.query("channel == 'خبر'")
shabake_khabar_Time_visit=shabake_khabar_Time_visit.copy()
shabake_khabar_Time_visit=shabake_khabar_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_khabar_Time_visit['میانگین']
del shabake_khabar_Time_visit['تاریخ']
del shabake_khabar_Time_visit['ردیف']
del shabake_khabar_Time_visit['tag']
shabake_khabar_Time_visit = shabake_khabar_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه خبر', 'مدت بازدید': 'مدت بازدید پربازدید شبکه خبر'})

print("shabake_ofogh")
shabake_ofogh=sima.query("channel == 'افق'")
shabake_ofogh_visit=shabake_ofogh['تعداد بازدید'].sum()
shabake_ofogh_duration=shabake_ofogh['مدت بازدید'].sum()
shabake_ofogh_duration=round(shabake_ofogh_duration*60, 0)
shabake_ofogh_content=shabake_ofogh.copy()
shabake_ofogh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_ofogh_content=len(shabake_ofogh_content)
shabake_ofogh_pivot=shabake_ofogh.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_ofogh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_ofogh_popular_visit=shabake_ofogh_pivot.iloc[0:10 , [0, 5]]
shabake_ofogh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_ofogh_popular_duration=shabake_ofogh_pivot.iloc[0:10 , [0, 4]]

shabake_ofogh_popular_visit = shabake_ofogh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه افق', 'نام برنامه': 'محتواهای پربازدید شبکه افق'})
shabake_ofogh_popular_duration = shabake_ofogh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه افق (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه افق'})

shabake_ofogh_Time_visit=all_data_Time.query("channel == 'افق'")
shabake_ofogh_Time_visit=shabake_ofogh_Time_visit.copy()
shabake_ofogh_Time_visit=shabake_ofogh_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_ofogh_Time_visit['میانگین']
del shabake_ofogh_Time_visit['تاریخ']
del shabake_ofogh_Time_visit['ردیف']
del shabake_ofogh_Time_visit['tag']
shabake_ofogh_Time_visit = shabake_ofogh_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه افق', 'مدت بازدید': 'مدت بازدید پربازدید شبکه افق'})

print("shabake_pooya")
shabake_pooya=sima.query("channel == 'پویا'")
shabake_pooya_visit=shabake_pooya['تعداد بازدید'].sum()
shabake_pooya_duration=shabake_pooya['مدت بازدید'].sum()
shabake_pooya_duration=round(shabake_pooya_duration*60, 0)
shabake_pooya_content=shabake_pooya.copy()
shabake_pooya_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_pooya_content=len(shabake_pooya_content)
shabake_pooya_pivot=shabake_pooya.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_pooya_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_pooya_popular_visit=shabake_pooya_pivot.iloc[0:10 , [0, 5]]
shabake_pooya_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_pooya_popular_duration=shabake_pooya_pivot.iloc[0:10 , [0, 4]]

shabake_pooya_popular_visit = shabake_pooya_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه پویا', 'نام برنامه': 'محتواهای پربازدید شبکه پویا'})
shabake_pooya_popular_duration = shabake_pooya_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه پویا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه پویا'})

shabake_pooya_Time_visit=all_data_Time.query("channel == 'پویا'")
shabake_pooya_Time_visit=shabake_pooya_Time_visit.copy()
shabake_pooya_Time_visit=shabake_pooya_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_pooya_Time_visit['میانگین']
del shabake_pooya_Time_visit['تاریخ']
del shabake_pooya_Time_visit['ردیف']
del shabake_pooya_Time_visit['tag']
shabake_pooya_Time_visit = shabake_pooya_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه پویا', 'مدت بازدید': 'مدت بازدید پربازدید شبکه پویا'})

print("shabake_omid")
shabake_omid=sima.query("channel == 'امید'")
shabake_omid_visit=shabake_omid['تعداد بازدید'].sum()
shabake_omid_duration=shabake_omid['مدت بازدید'].sum()
shabake_omid_duration=round(shabake_omid_duration*60, 0)
shabake_omid_content=shabake_omid.copy()
shabake_omid_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_omid_content=len(shabake_omid_content)
shabake_omid_pivot=shabake_omid.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_omid_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_omid_popular_visit=shabake_omid_pivot.iloc[0:10 , [0, 5]]
shabake_omid_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_omid_popular_duration=shabake_omid_pivot.iloc[0:10 , [0, 4]]

shabake_omid_popular_visit = shabake_omid_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه امید', 'نام برنامه': 'محتواهای پربازدید شبکه امید'})
shabake_omid_popular_duration = shabake_omid_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه امید (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه امید'})

shabake_omid_Time_visit=all_data_Time.query("channel == 'امید'")
shabake_omid_Time_visit=shabake_omid_Time_visit.copy()
shabake_omid_Time_visit=shabake_omid_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_omid_Time_visit['میانگین']
del shabake_omid_Time_visit['تاریخ']
del shabake_omid_Time_visit['ردیف']
del shabake_omid_Time_visit['tag']
shabake_omid_Time_visit = shabake_omid_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه امید', 'مدت بازدید': 'مدت بازدید پربازدید شبکه امید'})

print("shabake_ifilm")
shabake_ifilm=sima.query("channel == 'آی فیلم'")
shabake_ifilm_visit=shabake_ifilm['تعداد بازدید'].sum()
shabake_ifilm_duration=shabake_ifilm['مدت بازدید'].sum()
shabake_ifilm_duration=round(shabake_ifilm_duration*60, 0)
shabake_ifilm_content=shabake_ifilm.copy()
shabake_ifilm_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_ifilm_content=len(shabake_ifilm_content)
shabake_ifilm_pivot=shabake_ifilm.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_ifilm_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_ifilm_popular_visit=shabake_ifilm_pivot.iloc[0:10 , [0, 5]]
shabake_ifilm_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_ifilm_popular_duration=shabake_ifilm_pivot.iloc[0:10 , [0, 4]]

shabake_ifilm_popular_visit = shabake_ifilm_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آی فیلم', 'نام برنامه': 'محتواهای پربازدید شبکه آی فیلم'})
shabake_ifilm_popular_duration = shabake_ifilm_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه آی فیلم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه آی فیلم'})

shabake_ifilm_Time_visit=all_data_Time.query("channel == 'آی فیلم'")
shabake_ifilm_Time_visit=shabake_ifilm_Time_visit.copy()
shabake_ifilm_Time_visit=shabake_ifilm_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_ifilm_Time_visit['میانگین']
del shabake_ifilm_Time_visit['تاریخ']
del shabake_ifilm_Time_visit['ردیف']
del shabake_ifilm_Time_visit['tag']
shabake_ifilm_Time_visit = shabake_ifilm_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آی فیلم', 'مدت بازدید': 'مدت بازدید پربازدید شبکه آی فیلم'})

print("shabake_namayesh")
shabake_namayesh=sima.query("channel == 'نمایش'")
shabake_namayesh_visit=shabake_namayesh['تعداد بازدید'].sum()
shabake_namayesh_duration=shabake_namayesh['مدت بازدید'].sum()
shabake_namayesh_duration=round(shabake_namayesh_duration*60, 0)
shabake_namayesh_content=shabake_namayesh.copy()
shabake_namayesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_namayesh_content=len(shabake_namayesh_content)
shabake_namayesh_pivot=shabake_namayesh.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_namayesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_namayesh_popular_visit=shabake_namayesh_pivot.iloc[0:10 , [0, 5]]
shabake_namayesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_namayesh_popular_duration=shabake_namayesh_pivot.iloc[0:10 , [0, 4]]

shabake_namayesh_popular_visit = shabake_namayesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه نمایش', 'نام برنامه': 'محتواهای پربازدید شبکه نمایش'})
shabake_namayesh_popular_duration = shabake_namayesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه نمایش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه نمایش'})

shabake_namayesh_Time_visit=all_data_Time.query("channel == 'نمایش'")
shabake_namayesh_Time_visit=shabake_namayesh_Time_visit.copy()
shabake_namayesh_Time_visit=shabake_namayesh_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_namayesh_Time_visit['میانگین']
del shabake_namayesh_Time_visit['تاریخ']
del shabake_namayesh_Time_visit['ردیف']
del shabake_namayesh_Time_visit['tag']
shabake_namayesh_Time_visit = shabake_namayesh_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه نمایش', 'مدت بازدید': 'مدت بازدید پربازدید شبکه نمایش'})

print("shabake_tamasha")
shabake_tamasha=sima.query("channel == 'تماشا'")
shabake_tamasha_visit=shabake_tamasha['تعداد بازدید'].sum()
shabake_tamasha_duration=shabake_tamasha['مدت بازدید'].sum()
shabake_tamasha_duration=round(shabake_tamasha_duration*60, 0)
shabake_tamasha_content=shabake_tamasha.copy()
shabake_tamasha_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_tamasha_content=len(shabake_tamasha_content)
shabake_tamasha_pivot=shabake_tamasha.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_tamasha_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_tamasha_popular_visit=shabake_tamasha_pivot.iloc[0:10 , [0, 5]]
shabake_tamasha_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_tamasha_popular_duration=shabake_tamasha_pivot.iloc[0:10 , [0, 4]]

shabake_tamasha_popular_visit = shabake_tamasha_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه تماشا', 'نام برنامه': 'محتواهای پربازدید شبکه تماشا'})
shabake_tamasha_popular_duration = shabake_tamasha_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه تماشا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه تماشا'})

shabake_tamasha_Time_visit=all_data_Time.query("channel == 'تماشا'")
shabake_tamasha_Time_visit=shabake_tamasha_Time_visit.copy()
shabake_tamasha_Time_visit=shabake_tamasha_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_tamasha_Time_visit['میانگین']
del shabake_tamasha_Time_visit['تاریخ']
del shabake_tamasha_Time_visit['ردیف']
del shabake_tamasha_Time_visit['tag']
shabake_tamasha_Time_visit = shabake_tamasha_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه تماشا', 'مدت بازدید': 'مدت بازدید پربازدید شبکه تماشا'})

print("shabake_mostanad")
shabake_mostanad=sima.query("channel == 'مستند'")
shabake_mostanad_visit=shabake_mostanad['تعداد بازدید'].sum()
shabake_mostanad_duration=shabake_mostanad['مدت بازدید'].sum()
shabake_mostanad_duration=round(shabake_mostanad_duration*60, 0)
shabake_mostanad_content=shabake_mostanad.copy()
shabake_mostanad_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_mostanad_content=len(shabake_mostanad_content)
shabake_mostanad_pivot=shabake_mostanad.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_mostanad_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_mostanad_popular_visit=shabake_mostanad_pivot.iloc[0:10 , [0, 5]]
shabake_mostanad_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_mostanad_popular_duration=shabake_mostanad_pivot.iloc[0:10 , [0, 4]]

shabake_mostanad_popular_visit = shabake_mostanad_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه مستند', 'نام برنامه': 'محتواهای پربازدید شبکه مستند'})
shabake_mostanad_popular_duration = shabake_mostanad_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه مستند (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه مستند'})

shabake_mostanad_Time_visit=all_data_Time.query("channel == 'مستند'")
shabake_mostanad_Time_visit=shabake_mostanad_Time_visit.copy()
shabake_mostanad_Time_visit=shabake_mostanad_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_mostanad_Time_visit['میانگین']
del shabake_mostanad_Time_visit['تاریخ']
del shabake_mostanad_Time_visit['ردیف']
del shabake_mostanad_Time_visit['tag']
shabake_mostanad_Time_visit = shabake_mostanad_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه مستند', 'مدت بازدید': 'مدت بازدید پربازدید شبکه مستند'})

print("shabake_shoma")
shabake_shoma=sima.query("channel == 'شما'")
shabake_shoma_visit=shabake_shoma['تعداد بازدید'].sum()
shabake_shoma_duration=shabake_shoma['مدت بازدید'].sum()
shabake_shoma_duration=round(shabake_shoma_duration*60, 0)
shabake_shoma_content=shabake_shoma.copy()
shabake_shoma_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_shoma_content=len(shabake_shoma_content)
shabake_shoma_pivot=shabake_shoma.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_shoma_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_shoma_popular_visit=shabake_shoma_pivot.iloc[0:10 , [0, 5]]
shabake_shoma_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_shoma_popular_duration=shabake_shoma_pivot.iloc[0:10 , [0, 4]]

shabake_shoma_popular_visit = shabake_shoma_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه شما', 'نام برنامه': 'محتواهای پربازدید شبکه شما'})
shabake_shoma_popular_duration = shabake_shoma_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه شما (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه شما'})

shabake_shoma_Time_visit=all_data_Time.query("channel == 'شما'")
shabake_shoma_Time_visit=shabake_shoma_Time_visit.copy()
shabake_shoma_Time_visit=shabake_shoma_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_shoma_Time_visit['میانگین']
del shabake_shoma_Time_visit['تاریخ']
del shabake_shoma_Time_visit['ردیف']
del shabake_shoma_Time_visit['tag']
shabake_shoma_Time_visit = shabake_shoma_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه شما', 'مدت بازدید': 'مدت بازدید پربازدید شبکه شما'})

print("shabake_amozesh")
shabake_amozesh=sima.query("channel == 'آموزش'")
shabake_amozesh_visit=shabake_amozesh['تعداد بازدید'].sum()
shabake_amozesh_duration=shabake_amozesh['مدت بازدید'].sum()
shabake_amozesh_duration=round(shabake_amozesh_duration*60, 0)
shabake_amozesh_content=shabake_amozesh.copy()
shabake_amozesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_amozesh_content=len(shabake_amozesh_content)
shabake_amozesh_pivot=shabake_amozesh.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_amozesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_amozesh_popular_visit=shabake_amozesh_pivot.iloc[0:10 , [0, 5]]
shabake_amozesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_amozesh_popular_duration=shabake_amozesh_pivot.iloc[0:10 , [0, 4]]

shabake_amozesh_popular_visit = shabake_amozesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آموزش', 'نام برنامه': 'محتواهای پربازدید شبکه آموزش'})
shabake_amozesh_popular_duration = shabake_amozesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه آموزش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه آموزش'})

shabake_amozesh_Time_visit=all_data_Time.query("channel == 'آموزش'")
shabake_amozesh_Time_visit=shabake_amozesh_Time_visit.copy()
shabake_amozesh_Time_visit=shabake_amozesh_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_amozesh_Time_visit['میانگین']
del shabake_amozesh_Time_visit['تاریخ']
del shabake_amozesh_Time_visit['ردیف']
del shabake_amozesh_Time_visit['tag']
shabake_amozesh_Time_visit = shabake_amozesh_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آموزش', 'مدت بازدید': 'مدت بازدید پربازدید شبکه آموزش'})

print("shabake_varzesh")
shabake_varzesh=sima.query("channel == 'ورزش'")
shabake_varzesh_visit=shabake_varzesh['تعداد بازدید'].sum()
shabake_varzesh_duration=shabake_varzesh['مدت بازدید'].sum()
shabake_varzesh_duration=round(shabake_varzesh_duration*60, 0)
shabake_varzesh_content=shabake_varzesh.copy()
shabake_varzesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_varzesh_content=len(shabake_varzesh_content)
shabake_varzesh_pivot=shabake_varzesh.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_varzesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_varzesh_popular_visit=shabake_varzesh_pivot.iloc[0:10 , [0, 5]]
shabake_varzesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_varzesh_popular_duration=shabake_varzesh_pivot.iloc[0:10 , [0, 4]]

shabake_varzesh_popular_visit = shabake_varzesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه ورزش', 'نام برنامه': 'محتواهای پربازدید شبکه ورزش'})
shabake_varzesh_popular_duration = shabake_varzesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه ورزش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه ورزش'})

shabake_varzesh_Time_visit=all_data_Time.query("channel == 'ورزش'")
shabake_varzesh_Time_visit=shabake_varzesh_Time_visit.copy()
shabake_varzesh_Time_visit=shabake_varzesh_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_varzesh_Time_visit['میانگین']
del shabake_varzesh_Time_visit['تاریخ']
del shabake_varzesh_Time_visit['ردیف']
del shabake_varzesh_Time_visit['tag']
shabake_varzesh_Time_visit = shabake_varzesh_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه ورزش', 'مدت بازدید': 'مدت بازدید پربازدید شبکه ورزش'})

print("shabake_nasim")
shabake_nasim=sima.query("channel == 'نسیم'")
shabake_nasim_visit=shabake_nasim['تعداد بازدید'].sum()
shabake_nasim_duration=shabake_nasim['مدت بازدید'].sum()
shabake_nasim_duration=round(shabake_nasim_duration*60, 0)
shabake_nasim_content=shabake_nasim.copy()
shabake_nasim_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_nasim_content=len(shabake_nasim_content)
shabake_nasim_pivot=shabake_nasim.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_nasim_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_nasim_popular_visit=shabake_nasim_pivot.iloc[0:10 , [0, 5]]
shabake_nasim_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_nasim_popular_duration=shabake_nasim_pivot.iloc[0:10 , [0, 4]]

shabake_nasim_popular_visit = shabake_nasim_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه نسیم', 'نام برنامه': 'محتواهای پربازدید شبکه نسیم'})
shabake_nasim_popular_duration = shabake_nasim_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه نسیم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه نسیم'})

shabake_nasim_Time_visit=all_data_Time.query("channel == 'نسیم'")
shabake_nasim_Time_visit=shabake_nasim_Time_visit.copy()
shabake_nasim_Time_visit=shabake_nasim_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_nasim_Time_visit['میانگین']
del shabake_nasim_Time_visit['تاریخ']
del shabake_nasim_Time_visit['ردیف']
del shabake_nasim_Time_visit['tag']
shabake_nasim_Time_visit = shabake_nasim_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه نسیم', 'مدت بازدید': 'مدت بازدید پربازدید شبکه نسیم'})

print("shabake_qoran")
shabake_qoran=sima.query("channel == 'قرآن'")
shabake_qoran_visit=shabake_qoran['تعداد بازدید'].sum()
shabake_qoran_duration=shabake_qoran['مدت بازدید'].sum()
shabake_qoran_duration=round(shabake_qoran_duration*60, 0)
shabake_qoran_content=shabake_qoran.copy()
shabake_qoran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_qoran_content=len(shabake_qoran_content)
shabake_qoran_pivot=shabake_qoran.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_qoran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_qoran_popular_visit=shabake_qoran_pivot.iloc[0:10 , [0, 5]]
shabake_qoran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_qoran_popular_duration=shabake_qoran_pivot.iloc[0:10 , [0, 4]]

shabake_qoran_popular_visit = shabake_qoran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه قرآن', 'نام برنامه': 'محتواهای پربازدید شبکه قرآن'})
shabake_qoran_popular_duration = shabake_qoran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه قرآن (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه قرآن'})

shabake_qoran_Time_visit=all_data_Time.query("channel == 'قرآن'")
shabake_qoran_Time_visit=shabake_qoran_Time_visit.copy()
shabake_qoran_Time_visit=shabake_qoran_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_qoran_Time_visit['میانگین']
del shabake_qoran_Time_visit['تاریخ']
del shabake_qoran_Time_visit['ردیف']
del shabake_qoran_Time_visit['tag']
shabake_qoran_Time_visit = shabake_qoran_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه قرآن', 'مدت بازدید': 'مدت بازدید پربازدید شبکه قرآن'})

print("shabake_salamat")
shabake_salamat=sima.query("channel == 'سلامت'")
shabake_salamat_visit=shabake_salamat['تعداد بازدید'].sum()
shabake_salamat_duration=shabake_salamat['مدت بازدید'].sum()
shabake_salamat_duration=round(shabake_salamat_duration*60, 0)
shabake_salamat_content=shabake_salamat.copy()
shabake_salamat_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_salamat_content=len(shabake_salamat_content)
shabake_salamat_pivot=shabake_salamat.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_salamat_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_salamat_popular_visit=shabake_salamat_pivot.iloc[0:10 , [0, 5]]
shabake_salamat_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_salamat_popular_duration=shabake_salamat_pivot.iloc[0:10 , [0, 4]]

shabake_salamat_popular_visit = shabake_salamat_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سلامت', 'نام برنامه': 'محتواهای پربازدید شبکه سلامت'})
shabake_salamat_popular_duration = shabake_salamat_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه سلامت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه سلامت'})

shabake_salamat_Time_visit=all_data_Time.query("channel == 'سلامت'")
shabake_salamat_Time_visit=shabake_salamat_Time_visit.copy()
shabake_salamat_Time_visit=shabake_salamat_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_salamat_Time_visit['میانگین']
del shabake_salamat_Time_visit['تاریخ']
del shabake_salamat_Time_visit['ردیف']
del shabake_salamat_Time_visit['tag']
shabake_salamat_Time_visit = shabake_salamat_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سلامت', 'مدت بازدید': 'مدت بازدید پربازدید شبکه سلامت'})

print("shabake_irankala")
shabake_irankala=sima.query("channel == 'ایران کالا'")
shabake_irankala_visit=shabake_irankala['تعداد بازدید'].sum()
shabake_irankala_duration=shabake_irankala['مدت بازدید'].sum()
shabake_irankala_duration=round(shabake_irankala_duration*60, 0)
shabake_irankala_content=shabake_irankala.copy()
shabake_irankala_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_irankala_content=len(shabake_irankala_content)
shabake_irankala_pivot=shabake_irankala.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_irankala_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_irankala_popular_visit=shabake_irankala_pivot.iloc[0:10 , [0, 5]]
shabake_irankala_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_irankala_popular_duration=shabake_irankala_pivot.iloc[0:10 , [0, 4]]

shabake_irankala_popular_visit = shabake_irankala_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه ایران کالا', 'نام برنامه': 'محتواهای پربازدید شبکه ایران کالا'})
shabake_irankala_popular_duration = shabake_irankala_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه ایران کالا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه ایران کالا'})

shabake_irankala_Time_visit=all_data_Time.query("channel == 'ایران کالا'")
shabake_irankala_Time_visit=shabake_irankala_Time_visit.copy()
shabake_irankala_Time_visit=shabake_irankala_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_irankala_Time_visit['میانگین']
del shabake_irankala_Time_visit['تاریخ']
del shabake_irankala_Time_visit['ردیف']
del shabake_irankala_Time_visit['tag']
shabake_irankala_Time_visit = shabake_irankala_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه ایران کالا', 'مدت بازدید': 'مدت بازدید پربازدید شبکه ایران کالا'})

print("shabake_alalam")
shabake_alalam=sima.query("channel == 'العالم'")
shabake_alalam_visit=shabake_alalam['تعداد بازدید'].sum()
shabake_alalam_duration=shabake_alalam['مدت بازدید'].sum()
shabake_alalam_duration=round(shabake_alalam_duration*60, 0)
shabake_alalam_content=shabake_alalam.copy()
shabake_alalam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_alalam_content=len(shabake_alalam_content)
shabake_alalam_pivot=shabake_alalam.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_alalam_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_alalam_popular_visit=shabake_alalam_pivot.iloc[0:10 , [0, 5]]
shabake_alalam_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_alalam_popular_duration=shabake_alalam_pivot.iloc[0:10 , [0, 4]]

shabake_alalam_popular_visit = shabake_alalam_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه العالم', 'نام برنامه': 'محتواهای پربازدید شبکه العالم'})
shabake_alalam_popular_duration = shabake_alalam_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه العالم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه العالم'})

shabake_alalam_Time_visit=all_data_Time.query("channel == 'العالم'")
shabake_alalam_Time_visit=shabake_alalam_Time_visit.copy()
shabake_alalam_Time_visit=shabake_alalam_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_alalam_Time_visit['میانگین']
del shabake_alalam_Time_visit['تاریخ']
del shabake_alalam_Time_visit['ردیف']
del shabake_alalam_Time_visit['tag']
shabake_alalam_Time_visit = shabake_alalam_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه العالم', 'مدت بازدید': 'مدت بازدید پربازدید شبکه العالم'})

print("shabake_alkosar")
shabake_alkosar=sima.query("channel == 'الکوثر'")
shabake_alkosar_visit=shabake_alkosar['تعداد بازدید'].sum()
shabake_alkosar_duration=shabake_alkosar['مدت بازدید'].sum()
shabake_alkosar_duration=round(shabake_alkosar_duration*60, 0)
shabake_alkosar_content=shabake_alkosar.copy()
shabake_alkosar_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_alkosar_content=len(shabake_alkosar_content)
shabake_alkosar_pivot=shabake_alkosar.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_alkosar_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_alkosar_popular_visit=shabake_alkosar_pivot.iloc[0:10 , [0, 5]]
shabake_alkosar_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_alkosar_popular_duration=shabake_alkosar_pivot.iloc[0:10 , [0, 4]]

shabake_alkosar_popular_visit = shabake_alkosar_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه الکوثر', 'نام برنامه': 'محتواهای پربازدید شبکه الکوثر'})
shabake_alkosar_popular_duration = shabake_alkosar_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه الکوثر (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه الکوثر'})

shabake_alkosar_Time_visit=all_data_Time.query("channel == 'الکوثر'")
shabake_alkosar_Time_visit=shabake_alkosar_Time_visit.copy()
shabake_alkosar_Time_visit=shabake_alkosar_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_alkosar_Time_visit['میانگین']
del shabake_alkosar_Time_visit['تاریخ']
del shabake_alkosar_Time_visit['ردیف']
del shabake_alkosar_Time_visit['tag']
shabake_alkosar_Time_visit = shabake_alkosar_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه الکوثر', 'مدت بازدید': 'مدت بازدید پربازدید شبکه الکوثر'})

print("shabake_presstv")
shabake_presstv=sima.query("channel == 'پرس تی وی'")
shabake_presstv_visit=shabake_presstv['تعداد بازدید'].sum()
shabake_presstv_duration=shabake_presstv['مدت بازدید'].sum()
shabake_presstv_duration=round(shabake_presstv_duration*60, 0)
shabake_presstv_content=shabake_presstv.copy()
shabake_presstv_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_presstv_content=len(shabake_presstv_content)
shabake_presstv_pivot=shabake_presstv.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_presstv_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_presstv_popular_visit=shabake_presstv_pivot.iloc[0:10 , [0, 5]]
shabake_presstv_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_presstv_popular_duration=shabake_presstv_pivot.iloc[0:10 , [0, 4]]

shabake_presstv_popular_visit = shabake_presstv_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه پرس تی وی', 'نام برنامه': 'محتواهای پربازدید شبکه پرس تی وی'})
shabake_presstv_popular_duration = shabake_presstv_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه پرس تی وی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه پرس تی وی'})

shabake_presstv_Time_visit=all_data_Time.query("channel == 'پرس تی وی'")
shabake_presstv_Time_visit=shabake_presstv_Time_visit.copy()
shabake_presstv_Time_visit=shabake_presstv_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_presstv_Time_visit['میانگین']
del shabake_presstv_Time_visit['تاریخ']
del shabake_presstv_Time_visit['ردیف']
del shabake_presstv_Time_visit['tag']
shabake_presstv_Time_visit = shabake_presstv_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه پرس تی وی', 'مدت بازدید': 'مدت بازدید پربازدید شبکه پرس تی وی'})

print("shabake_sepehr")
shabake_sepehr=sima.query("channel == 'سپهر'")
shabake_sepehr_visit=shabake_sepehr['تعداد بازدید'].sum()
shabake_sepehr_duration=shabake_sepehr['مدت بازدید'].sum()
shabake_sepehr_duration=round(shabake_sepehr_duration*60, 0)
shabake_sepehr_content=shabake_sepehr.copy()
shabake_sepehr_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_sepehr_content=len(shabake_sepehr_content)
shabake_sepehr_pivot=shabake_sepehr.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_sepehr_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_sepehr_popular_visit=shabake_sepehr_pivot.iloc[0:10 , [0, 5]]
shabake_sepehr_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_sepehr_popular_duration=shabake_sepehr_pivot.iloc[0:10 , [0, 4]]

shabake_sepehr_popular_visit = shabake_sepehr_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سپهر', 'نام برنامه': 'محتواهای پربازدید شبکه سپهر'})
shabake_sepehr_popular_duration = shabake_sepehr_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه سپهر (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه سپهر'})

shabake_sepehr_Time_visit=all_data_Time.query("channel == 'سپهر'")
shabake_sepehr_Time_visit=shabake_sepehr_Time_visit.copy()
shabake_sepehr_Time_visit=shabake_sepehr_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_sepehr_Time_visit['میانگین']
del shabake_sepehr_Time_visit['تاریخ']
del shabake_sepehr_Time_visit['ردیف']
del shabake_sepehr_Time_visit['tag']
shabake_sepehr_Time_visit = shabake_sepehr_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سپهر', 'مدت بازدید': 'مدت بازدید پربازدید شبکه سپهر'})

print("shabake_jamejam")
shabake_jamejam=sima.query("channel == 'جام جم 1'")
shabake_jamejam_visit=shabake_jamejam['تعداد بازدید'].sum()
shabake_jamejam_duration=shabake_jamejam['مدت بازدید'].sum()
shabake_jamejam_duration=round(shabake_jamejam_duration*60, 0)
shabake_jamejam_content=shabake_jamejam.copy()
shabake_jamejam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
shabake_jamejam_content=len(shabake_jamejam_content)
shabake_jamejam_pivot=shabake_jamejam.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
shabake_jamejam_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_jamejam_popular_visit=shabake_jamejam_pivot.iloc[0:10 , [0, 5]]
shabake_jamejam_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
shabake_jamejam_popular_duration=shabake_jamejam_pivot.iloc[0:10 , [0, 4]]

shabake_jamejam_popular_visit = shabake_jamejam_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه جام جم', 'نام برنامه': 'محتواهای پربازدید شبکه جام جم'})
shabake_jamejam_popular_duration = shabake_jamejam_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه جام جم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه جام جم'})

shabake_jamejam_Time_visit=all_data_Time.query("channel == 'جام جم 1'")
shabake_jamejam_Time_visit=shabake_jamejam_Time_visit.copy()
shabake_jamejam_Time_visit=shabake_jamejam_Time_visit.groupby(['ساعت']).sum().reset_index()
del shabake_jamejam_Time_visit['میانگین']
del shabake_jamejam_Time_visit['تاریخ']
del shabake_jamejam_Time_visit['ردیف']
del shabake_jamejam_Time_visit['tag']
shabake_jamejam_Time_visit = shabake_jamejam_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه جام جم', 'مدت بازدید': 'مدت بازدید پربازدید شبکه جام جم'})

print("dataframe sima channels")
sima_channels_statistics={'channel_name': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر', 'شبکه جام جم',],
       'channel_content': [shabake_1_content, shabake_2_content, shabake_3_content, shabake_5_content, shabake_5_content,
                           shabake_khabar_content, shabake_ofogh_content, shabake_pooya_content, shabake_omid_content, shabake_ifilm_content,
                           shabake_namayesh_content, shabake_tamasha_content, shabake_mostanad_content, shabake_shoma_content, shabake_amozesh_content,
                           shabake_varzesh_content, shabake_nasim_content, shabake_qoran_content, shabake_salamat_content, shabake_irankala_content,
                           shabake_alalam_content, shabake_alkosar_content, shabake_presstv_content, shabake_sepehr_content, shabake_jamejam_content,],
       'channel_visit': [shabake_1_visit, shabake_2_visit, shabake_3_visit, shabake_5_visit, shabake_5_visit,
                           shabake_khabar_visit, shabake_ofogh_visit, shabake_pooya_visit, shabake_omid_visit, shabake_ifilm_visit,
                           shabake_namayesh_visit, shabake_tamasha_visit, shabake_mostanad_visit, shabake_shoma_visit, shabake_amozesh_visit,
                           shabake_varzesh_visit, shabake_nasim_visit, shabake_qoran_visit, shabake_salamat_visit, shabake_irankala_visit,
                           shabake_alalam_visit, shabake_alkosar_visit, shabake_presstv_visit, shabake_sepehr_visit, shabake_jamejam_visit,],
       'channel_duration': [shabake_1_duration, shabake_2_duration, shabake_3_duration, shabake_5_duration, shabake_5_duration,
                           shabake_khabar_duration, shabake_ofogh_duration, shabake_pooya_duration, shabake_omid_duration, shabake_ifilm_duration,
                           shabake_namayesh_duration, shabake_tamasha_duration, shabake_mostanad_duration, shabake_shoma_duration, shabake_amozesh_duration,
                           shabake_varzesh_duration, shabake_nasim_duration, shabake_qoran_duration, shabake_salamat_duration, shabake_irankala_duration,
                           shabake_alalam_duration, shabake_alkosar_duration, shabake_presstv_duration, shabake_sepehr_duration, shabake_jamejam_duration,],}
sima_channels_statistics=pd.DataFrame(sima_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
sima_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
sima_channels_statistics=radio_channels_statistics.rename(columns={'channel_content': 'نام شبکه', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})

print("sima channels popular contents")

shabake_1_popular_visit.to_excel('busy/shabake_1_popular_visit.xlsx')
shabake_1_popular_duration.to_excel('busy/shabake_1_popular_duration.xlsx')
shabake_1_popular_visit=pd.read_excel('busy/shabake_1_popular_visit.xlsx')
shabake_1_popular_duration=pd.read_excel('busy/shabake_1_popular_duration.xlsx')
del shabake_1_popular_visit['Unnamed: 0']
del shabake_1_popular_duration['Unnamed: 0']

shabake_2_popular_visit.to_excel('busy/shabake_2_popular_visit.xlsx')
shabake_2_popular_duration.to_excel('busy/shabake_2_popular_duration.xlsx')
shabake_2_popular_visit=pd.read_excel('busy/shabake_2_popular_visit.xlsx')
shabake_2_popular_duration=pd.read_excel('busy/shabake_2_popular_duration.xlsx')
del shabake_2_popular_visit['Unnamed: 0']
del shabake_2_popular_duration['Unnamed: 0']

shabake_3_popular_visit.to_excel('busy/shabake_3_popular_visit.xlsx')
shabake_3_popular_duration.to_excel('busy/shabake_3_popular_duration.xlsx')
shabake_3_popular_visit=pd.read_excel('busy/shabake_3_popular_visit.xlsx')
shabake_3_popular_duration=pd.read_excel('busy/shabake_3_popular_duration.xlsx')
del shabake_3_popular_visit['Unnamed: 0']
del shabake_3_popular_duration['Unnamed: 0']

shabake_4_popular_visit.to_excel('busy/shabake_4_popular_visit.xlsx')
shabake_4_popular_duration.to_excel('busy/shabake_4_popular_duration.xlsx')
shabake_4_popular_visit=pd.read_excel('busy/shabake_4_popular_visit.xlsx')
shabake_4_popular_duration=pd.read_excel('busy/shabake_4_popular_duration.xlsx')
del shabake_4_popular_visit['Unnamed: 0']
del shabake_4_popular_duration['Unnamed: 0']

shabake_5_popular_visit.to_excel('busy/shabake_5_popular_visit.xlsx')
shabake_5_popular_duration.to_excel('busy/shabake_5_popular_duration.xlsx')
shabake_5_popular_visit=pd.read_excel('busy/shabake_5_popular_visit.xlsx')
shabake_5_popular_duration=pd.read_excel('busy/shabake_5_popular_duration.xlsx')
del shabake_5_popular_visit['Unnamed: 0']
del shabake_5_popular_duration['Unnamed: 0']

shabake_khabar_popular_visit.to_excel('busy/shabake_khabar_popular_visit.xlsx')
shabake_khabar_popular_duration.to_excel('busy/shabake_khabar_popular_duration.xlsx')
shabake_khabar_popular_visit=pd.read_excel('busy/shabake_khabar_popular_visit.xlsx')
shabake_khabar_popular_duration=pd.read_excel('busy/shabake_khabar_popular_duration.xlsx')
del shabake_khabar_popular_visit['Unnamed: 0']
del shabake_khabar_popular_duration['Unnamed: 0']

shabake_ofogh_popular_visit.to_excel('busy/shabake_ofogh_popular_visit.xlsx')
shabake_ofogh_popular_duration.to_excel('busy/shabake_ofogh_popular_duration.xlsx')
shabake_ofogh_popular_visit=pd.read_excel('busy/shabake_ofogh_popular_visit.xlsx')
shabake_ofogh_popular_duration=pd.read_excel('busy/shabake_ofogh_popular_duration.xlsx')
del shabake_ofogh_popular_visit['Unnamed: 0']
del shabake_ofogh_popular_duration['Unnamed: 0']

shabake_pooya_popular_visit.to_excel('busy/shabake_pooya_popular_visit.xlsx')
shabake_pooya_popular_duration.to_excel('busy/shabake_pooya_popular_duration.xlsx')
shabake_pooya_popular_visit=pd.read_excel('busy/shabake_pooya_popular_visit.xlsx')
shabake_pooya_popular_duration=pd.read_excel('busy/shabake_pooya_popular_duration.xlsx')
del shabake_pooya_popular_visit['Unnamed: 0']
del shabake_pooya_popular_duration['Unnamed: 0']

shabake_omid_popular_visit.to_excel('busy/shabake_omid_popular_visit.xlsx')
shabake_omid_popular_duration.to_excel('busy/shabake_omid_popular_duration.xlsx')
shabake_omid_popular_visit=pd.read_excel('busy/shabake_omid_popular_visit.xlsx')
shabake_omid_popular_duration=pd.read_excel('busy/shabake_omid_popular_duration.xlsx')
del shabake_omid_popular_visit['Unnamed: 0']
del shabake_omid_popular_duration['Unnamed: 0']

shabake_ifilm_popular_visit.to_excel('busy/shabake_ifilm_popular_visit.xlsx')
shabake_ifilm_popular_duration.to_excel('busy/shabake_ifilm_popular_duration.xlsx')
shabake_ifilm_popular_visit=pd.read_excel('busy/shabake_ifilm_popular_visit.xlsx')
shabake_ifilm_popular_duration=pd.read_excel('busy/shabake_ifilm_popular_duration.xlsx')
del shabake_ifilm_popular_visit['Unnamed: 0']
del shabake_ifilm_popular_duration['Unnamed: 0']

shabake_namayesh_popular_visit.to_excel('busy/shabake_namayesh_popular_visit.xlsx')
shabake_namayesh_popular_duration.to_excel('busy/shabake_namayesh_popular_duration.xlsx')
shabake_namayesh_popular_visit=pd.read_excel('busy/shabake_namayesh_popular_visit.xlsx')
shabake_namayesh_popular_duration=pd.read_excel('busy/shabake_namayesh_popular_duration.xlsx')
del shabake_namayesh_popular_visit['Unnamed: 0']
del shabake_namayesh_popular_duration['Unnamed: 0']

shabake_tamasha_popular_visit.to_excel('busy/shabake_tamasha_popular_visit.xlsx')
shabake_tamasha_popular_duration.to_excel('busy/shabake_tamasha_popular_duration.xlsx')
shabake_tamasha_popular_visit=pd.read_excel('busy/shabake_tamasha_popular_visit.xlsx')
shabake_tamasha_popular_duration=pd.read_excel('busy/shabake_tamasha_popular_duration.xlsx')
del shabake_tamasha_popular_visit['Unnamed: 0']
del shabake_tamasha_popular_duration['Unnamed: 0']

shabake_mostanad_popular_visit.to_excel('busy/shabake_mostanad_popular_visit.xlsx')
shabake_mostanad_popular_duration.to_excel('busy/shabake_mostanad_popular_duration.xlsx')
shabake_mostanad_popular_visit=pd.read_excel('busy/shabake_mostanad_popular_visit.xlsx')
shabake_mostanad_popular_duration=pd.read_excel('busy/shabake_mostanad_popular_duration.xlsx')
del shabake_mostanad_popular_visit['Unnamed: 0']
del shabake_mostanad_popular_duration['Unnamed: 0']

shabake_shoma_popular_visit.to_excel('busy/shabake_shoma_popular_visit.xlsx')
shabake_shoma_popular_duration.to_excel('busy/shabake_shoma_popular_duration.xlsx')
shabake_shoma_popular_visit=pd.read_excel('busy/shabake_shoma_popular_visit.xlsx')
shabake_shoma_popular_duration=pd.read_excel('busy/shabake_shoma_popular_duration.xlsx')
del shabake_shoma_popular_visit['Unnamed: 0']
del shabake_shoma_popular_duration['Unnamed: 0']

shabake_amozesh_popular_visit.to_excel('busy/shabake_amozesh_popular_visit.xlsx')
shabake_amozesh_popular_duration.to_excel('busy/shabake_amozesh_popular_duration.xlsx')
shabake_amozesh_popular_visit=pd.read_excel('busy/shabake_amozesh_popular_visit.xlsx')
shabake_amozesh_popular_duration=pd.read_excel('busy/shabake_amozesh_popular_duration.xlsx')
del shabake_amozesh_popular_visit['Unnamed: 0']
del shabake_amozesh_popular_duration['Unnamed: 0']

shabake_varzesh_popular_visit.to_excel('busy/shabake_varzesh_popular_visit.xlsx')
shabake_varzesh_popular_duration.to_excel('busy/shabake_varzesh_popular_duration.xlsx')
shabake_varzesh_popular_visit=pd.read_excel('busy/shabake_varzesh_popular_visit.xlsx')
shabake_varzesh_popular_duration=pd.read_excel('busy/shabake_varzesh_popular_duration.xlsx')
del shabake_varzesh_popular_visit['Unnamed: 0']
del shabake_varzesh_popular_duration['Unnamed: 0']

shabake_nasim_popular_visit.to_excel('busy/shabake_nasim_popular_visit.xlsx')
shabake_nasim_popular_duration.to_excel('busy/shabake_nasim_popular_duration.xlsx')
shabake_nasim_popular_visit=pd.read_excel('busy/shabake_nasim_popular_visit.xlsx')
shabake_nasim_popular_duration=pd.read_excel('busy/shabake_nasim_popular_duration.xlsx')
del shabake_nasim_popular_visit['Unnamed: 0']
del shabake_nasim_popular_duration['Unnamed: 0']

shabake_qoran_popular_visit.to_excel('busy/shabake_qoran_popular_visit.xlsx')
shabake_qoran_popular_duration.to_excel('busy/shabake_qoran_popular_duration.xlsx')
shabake_qoran_popular_visit=pd.read_excel('busy/shabake_qoran_popular_visit.xlsx')
shabake_qoran_popular_duration=pd.read_excel('busy/shabake_qoran_popular_duration.xlsx')
del shabake_qoran_popular_visit['Unnamed: 0']
del shabake_qoran_popular_duration['Unnamed: 0']

shabake_salamat_popular_visit.to_excel('busy/shabake_salamat_popular_visit.xlsx')
shabake_salamat_popular_duration.to_excel('busy/shabake_salamat_popular_duration.xlsx')
shabake_salamat_popular_visit=pd.read_excel('busy/shabake_salamat_popular_visit.xlsx')
shabake_salamat_popular_duration=pd.read_excel('busy/shabake_salamat_popular_duration.xlsx')
del shabake_salamat_popular_visit['Unnamed: 0']
del shabake_salamat_popular_duration['Unnamed: 0']

shabake_irankala_popular_visit.to_excel('busy/shabake_irankala_popular_visit.xlsx')
shabake_irankala_popular_duration.to_excel('busy/shabake_irankala_popular_duration.xlsx')
shabake_irankala_popular_visit=pd.read_excel('busy/shabake_irankala_popular_visit.xlsx')
shabake_irankala_popular_duration=pd.read_excel('busy/shabake_irankala_popular_duration.xlsx')
del shabake_irankala_popular_visit['Unnamed: 0']
del shabake_irankala_popular_duration['Unnamed: 0']

shabake_alalam_popular_visit.to_excel('busy/shabake_alalam_popular_visit.xlsx')
shabake_alalam_popular_duration.to_excel('busy/shabake_alalam_popular_duration.xlsx')
shabake_alalam_popular_visit=pd.read_excel('busy/shabake_alalam_popular_visit.xlsx')
shabake_alalam_popular_duration=pd.read_excel('busy/shabake_alalam_popular_duration.xlsx')
del shabake_alalam_popular_visit['Unnamed: 0']
del shabake_alalam_popular_duration['Unnamed: 0']

shabake_alkosar_popular_visit.to_excel('busy/shabake_alkosar_popular_visit.xlsx')
shabake_alkosar_popular_duration.to_excel('busy/shabake_alkosar_popular_duration.xlsx')
shabake_alkosar_popular_visit=pd.read_excel('busy/shabake_alkosar_popular_visit.xlsx')
shabake_alkosar_popular_duration=pd.read_excel('busy/shabake_alkosar_popular_duration.xlsx')
del shabake_alkosar_popular_visit['Unnamed: 0']
del shabake_alkosar_popular_duration['Unnamed: 0']

shabake_presstv_popular_visit.to_excel('busy/shabake_presstv_popular_visit.xlsx')
shabake_presstv_popular_duration.to_excel('busy/shabake_presstv_popular_duration.xlsx')
shabake_presstv_popular_visit=pd.read_excel('busy/shabake_presstv_popular_visit.xlsx')
shabake_presstv_popular_duration=pd.read_excel('busy/shabake_presstv_popular_duration.xlsx')
del shabake_presstv_popular_visit['Unnamed: 0']
del shabake_presstv_popular_duration['Unnamed: 0']

shabake_sepehr_popular_visit.to_excel('busy/shabake_sepehr_popular_visit.xlsx')
shabake_sepehr_popular_duration.to_excel('busy/shabake_sepehr_popular_duration.xlsx')
shabake_sepehr_popular_visit=pd.read_excel('busy/shabake_sepehr_popular_visit.xlsx')
shabake_sepehr_popular_duration=pd.read_excel('busy/shabake_sepehr_popular_duration.xlsx')
del shabake_sepehr_popular_visit['Unnamed: 0']
del shabake_sepehr_popular_duration['Unnamed: 0']

shabake_jamejam_popular_visit.to_excel('busy/shabake_jamejam_popular_visit.xlsx')
shabake_jamejam_popular_duration.to_excel('busy/shabake_jamejam_popular_duration.xlsx')
shabake_jamejam_popular_visit=pd.read_excel('busy/shabake_jamejam_popular_visit.xlsx')
shabake_jamejam_popular_duration=pd.read_excel('busy/shabake_jamejam_popular_duration.xlsx')
del shabake_jamejam_popular_visit['Unnamed: 0']
del shabake_jamejam_popular_duration['Unnamed: 0']

sima_channels_popular_content=pd.DataFrame()
sima_channels_popular_content=pd.concat([shabake_1_popular_visit, shabake_1_popular_duration,
               shabake_2_popular_visit, shabake_2_popular_duration,
               shabake_3_popular_visit, shabake_3_popular_duration,
               shabake_4_popular_visit, shabake_4_popular_duration,
               shabake_5_popular_visit, shabake_5_popular_duration,
               shabake_khabar_popular_visit, shabake_khabar_popular_duration,
               shabake_ofogh_popular_visit, shabake_ofogh_popular_duration,
               shabake_pooya_popular_visit, shabake_pooya_popular_duration,
               shabake_omid_popular_visit, shabake_omid_popular_duration,
               shabake_ifilm_popular_visit, shabake_ifilm_popular_duration,
               shabake_namayesh_popular_visit, shabake_namayesh_popular_duration,
               shabake_tamasha_popular_visit, shabake_tamasha_popular_duration,
               shabake_mostanad_popular_visit, shabake_mostanad_popular_duration,
               shabake_shoma_popular_visit, shabake_shoma_popular_duration,
               shabake_amozesh_popular_visit, shabake_amozesh_popular_duration,
               shabake_varzesh_popular_visit, shabake_varzesh_popular_duration,
               shabake_nasim_popular_visit, shabake_nasim_popular_duration,
               shabake_qoran_popular_visit, shabake_qoran_popular_duration,
               shabake_salamat_popular_visit, shabake_salamat_popular_duration,
               shabake_irankala_popular_visit, shabake_irankala_popular_duration,
               shabake_alalam_popular_visit, shabake_alalam_popular_duration,
               shabake_alkosar_popular_visit, shabake_alkosar_popular_duration,
               shabake_presstv_popular_visit, shabake_presstv_popular_duration,
               shabake_sepehr_popular_visit, shabake_sepehr_popular_duration,
               shabake_jamejam_popular_visit, shabake_jamejam_popular_duration],axis=1)

del shabake_2_Time_visit['ساعت']
del shabake_3_Time_visit['ساعت']
del shabake_4_Time_visit['ساعت']
del shabake_5_Time_visit['ساعت']
del shabake_khabar_Time_visit['ساعت']
del shabake_ofogh_Time_visit['ساعت']
del shabake_pooya_Time_visit['ساعت']
del shabake_omid_Time_visit['ساعت']
del shabake_ifilm_Time_visit['ساعت']
del shabake_namayesh_Time_visit['ساعت']
del shabake_tamasha_Time_visit['ساعت']
del shabake_mostanad_Time_visit['ساعت']
del shabake_shoma_Time_visit['ساعت']
del shabake_amozesh_Time_visit['ساعت']
del shabake_varzesh_Time_visit['ساعت']
del shabake_nasim_Time_visit['ساعت']
del shabake_qoran_Time_visit['ساعت']
del shabake_salamat_Time_visit['ساعت']
del shabake_irankala_Time_visit['ساعت']
del shabake_alalam_Time_visit['ساعت']
del shabake_alkosar_Time_visit['ساعت']
del shabake_presstv_Time_visit['ساعت']
del shabake_sepehr_Time_visit['ساعت']
del shabake_jamejam_Time_visit['ساعت']

sima_channels_Time_visit=pd.DataFrame()
sima_channels_Time_visit=pd.concat([shabake_1_Time_visit, shabake_2_Time_visit, shabake_3_Time_visit,shabake_4_Time_visit, shabake_5_Time_visit, 
               shabake_khabar_Time_visit, shabake_ofogh_Time_visit, shabake_pooya_Time_visit,shabake_omid_Time_visit, shabake_ifilm_Time_visit,
               shabake_namayesh_Time_visit, shabake_tamasha_Time_visit, shabake_mostanad_Time_visit,shabake_shoma_Time_visit, shabake_amozesh_Time_visit,
               shabake_varzesh_Time_visit, shabake_nasim_Time_visit, shabake_qoran_Time_visit,shabake_salamat_Time_visit, shabake_irankala_Time_visit,
               shabake_alalam_Time_visit, shabake_alkosar_Time_visit, shabake_presstv_Time_visit,shabake_sepehr_Time_visit, shabake_jamejam_Time_visit,],axis=1)

writer = pd.ExcelWriter('output/آمار سیما.xlsx', engine='xlsxwriter')
sima_channels_statistics.to_excel(writer, 'آمار شبکه های سیما')
sima_channels_popular_content.to_excel(writer, 'محتواهای پربازدید')
sima_channels_Time_visit.to_excel(writer, 'آمار ساعتی')
writer.save()

print("End sima")
################################# RADIO ########################################
print("start radio")

radio_all=radio.copy()
radio_all_pivot=radio_all.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
radio_all_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_all_popular_visit=radio_all_pivot.iloc[0:10 , [0, 4]]
radio_all_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_all_popular_duration=radio_all_pivot.iloc[0:10 , [0, 3]]


print("radio_eghtesad")
radio_eghtesad=radio.query("channel == 'رادیو اقتصاد'")
radio_eghtesad_visit=radio_eghtesad['تعداد بازدید'].sum()
radio_eghtesad_duration=radio_eghtesad['مدت بازدید'].sum()
radio_eghtesad_duration=round(radio_eghtesad_duration*60, 0)
radio_eghtesad_content=radio_eghtesad.copy()
radio_eghtesad_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_eghtesad_content=len(radio_eghtesad_content)
radio_eghtesad_pivot=radio_eghtesad.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_eghtesad_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_eghtesad_popular_visit=radio_eghtesad_pivot.iloc[0:10 , [0, 5]]
radio_eghtesad_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_eghtesad_popular_duration=radio_eghtesad_pivot.iloc[0:10 , [0, 4]]

radio_eghtesad_popular_visit = radio_eghtesad_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو اقتصاد', 'نام برنامه': 'محتواهای پربازدید رادیو اقتصاد'})
radio_eghtesad_popular_duration = radio_eghtesad_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو اقتصاد (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو اقتصاد'})

print("radio_ava")
radio_ava=radio.query("channel == 'رادیو آوا'")
radio_ava_visit=radio_ava['تعداد بازدید'].sum()
radio_ava_duration=radio_ava['مدت بازدید'].sum()
radio_ava_duration=round(radio_ava_duration*60, 0)
radio_ava_content=radio_ava.copy()
radio_ava_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_ava_content=len(radio_ava_content)
radio_ava_pivot=radio_ava.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_ava_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_ava_popular_visit=radio_ava_pivot.iloc[0:10 , [0, 5]]
radio_ava_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_ava_popular_duration=radio_ava_pivot.iloc[0:10 , [0, 4]]

radio_ava_popular_visit = radio_ava_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو آوا', 'نام برنامه': 'محتواهای پربازدید رادیو آوا'})
radio_ava_popular_duration = radio_ava_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو آوا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو آوا'})

print("radio_iran")
radio_iran=radio.query("channel == 'رادیو ایران'")
radio_iran_visit=radio_iran['تعداد بازدید'].sum()
radio_iran_duration=radio_iran['مدت بازدید'].sum()
radio_iran_duration=round(radio_iran_duration*60, 0)
radio_iran_content=radio_iran.copy()
radio_iran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_iran_content=len(radio_iran_content)
radio_iran_pivot=radio_iran.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_iran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_iran_popular_visit=radio_iran_pivot.iloc[0:10 , [0, 5]]
radio_iran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_iran_popular_duration=radio_iran_pivot.iloc[0:10 , [0, 4]]

radio_iran_popular_visit = radio_iran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو ایران', 'نام برنامه': 'محتواهای پربازدید رادیو ایران'})
radio_iran_popular_duration = radio_iran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو ایران (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو ایران'})

print("radio_payam")
radio_payam=radio.query("channel == 'رادیو پیام'")
radio_payam_visit=radio_payam['تعداد بازدید'].sum()
radio_payam_duration=radio_payam['مدت بازدید'].sum()
radio_payam_duration=round(radio_payam_duration*60, 0)
radio_payam_content=radio_payam.copy()
radio_payam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_payam_content=len(radio_payam_content)
radio_payam_pivot=radio_payam.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_payam_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_payam_popular_visit=radio_payam_pivot.iloc[0:10 , [0, 5]]
radio_payam_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_payam_popular_duration=radio_payam_pivot.iloc[0:10 , [0, 4]]

radio_payam_popular_visit = radio_payam_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو پیام', 'نام برنامه': 'محتواهای پربازدید رادیو پیام'})
radio_payam_popular_duration = radio_payam_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو پیام (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو پیام'})

print("radio_javan")
radio_javan=radio.query("channel == 'رادیو جوان'")
radio_javan_visit=radio_javan['تعداد بازدید'].sum()
radio_javan_duration=radio_javan['مدت بازدید'].sum()
radio_javan_duration=round(radio_javan_duration*60, 0)
radio_javan_content=radio_javan.copy()
radio_javan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_javan_content=len(radio_javan_content)
radio_javan_pivot=radio_javan.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_javan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_javan_popular_visit=radio_javan_pivot.iloc[0:10 , [0, 5]]
radio_javan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_javan_popular_duration=radio_javan_pivot.iloc[0:10 , [0, 4]]

radio_javan_popular_visit = radio_javan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو جوان', 'نام برنامه': 'محتواهای پربازدید رادیو جوان'})
radio_javan_popular_duration = radio_javan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو جوان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو جوان'})

print("radio_salamat")
radio_salamat=radio.query("channel == 'رادیو سلامت'")
radio_salamat_visit=radio_salamat['تعداد بازدید'].sum()
radio_salamat_duration=radio_salamat['مدت بازدید'].sum()
radio_salamat_duration=round(radio_salamat_duration*60, 0)
radio_salamat_content=radio_salamat.copy()
radio_salamat_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_salamat_content=len(radio_salamat_content)
radio_salamat_pivot=radio_salamat.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_salamat_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_salamat_popular_visit=radio_salamat_pivot.iloc[0:10 , [0, 5]]
radio_salamat_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_salamat_popular_duration=radio_salamat_pivot.iloc[0:10 , [0, 4]]

radio_salamat_popular_visit = radio_salamat_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو سلامت', 'نام برنامه': 'محتواهای پربازدید رادیو سلامت'})
radio_salamat_popular_duration = radio_salamat_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو سلامت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو سلامت'})

print("radio_saba")
radio_saba=radio.query("channel == 'رادیو صبا'")
radio_saba_visit=radio_saba['تعداد بازدید'].sum()
radio_saba_duration=radio_saba['مدت بازدید'].sum()
radio_saba_duration=round(radio_saba_duration*60, 0)
radio_saba_content=radio_saba.copy()
radio_saba_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_saba_content=len(radio_saba_content)
radio_saba_pivot=radio_saba.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_saba_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_saba_popular_visit=radio_saba_pivot.iloc[0:10 , [0, 5]]
radio_saba_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_saba_popular_duration=radio_saba_pivot.iloc[0:10 , [0, 4]]

radio_saba_popular_visit = radio_saba_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو سبا', 'نام برنامه': 'محتواهای پربازدید رادیو سبا'})
radio_saba_popular_duration = radio_saba_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو سبا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو سبا'})

print("radio_farhang")
radio_farhang=radio.query("channel == 'رادیو فرهنگ'")
radio_farhang_visit=radio_farhang['تعداد بازدید'].sum()
radio_farhang_duration=radio_farhang['مدت بازدید'].sum()
radio_farhang_duration=round(radio_farhang_duration*60, 0)
radio_farhang_content=radio_farhang.copy()
radio_farhang_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_farhang_content=len(radio_farhang_content)
radio_farhang_pivot=radio_farhang.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_farhang_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_farhang_popular_visit=radio_farhang_pivot.iloc[0:10 , [0, 5]]
radio_farhang_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_farhang_popular_duration=radio_farhang_pivot.iloc[0:10 , [0, 4]]

radio_farhang_popular_visit = radio_farhang_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو فرهنگ', 'نام برنامه': 'محتواهای پربازدید رادیو فرهنگ'})
radio_farhang_popular_duration = radio_farhang_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو فرهنگ (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو فرهنگ'})

print("radio_qoran")
radio_qoran=radio.query("channel == 'رادیو قرآن'")
radio_qoran_visit=radio_qoran['تعداد بازدید'].sum()
radio_qoran_duration=radio_qoran['مدت بازدید'].sum()
radio_qoran_duration=round(radio_qoran_duration*60, 0)
radio_qoran_content=radio_qoran.copy()
radio_qoran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_qoran_content=len(radio_qoran_content)
radio_qoran_pivot=radio_qoran.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_qoran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_qoran_popular_visit=radio_qoran_pivot.iloc[0:10 , [0, 5]]
radio_qoran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_qoran_popular_duration=radio_qoran_pivot.iloc[0:10 , [0, 4]]

radio_qoran_popular_visit = radio_qoran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو قرآن', 'نام برنامه': 'محتواهای پربازدید رادیو قرآن'})
radio_qoran_popular_duration = radio_qoran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو قرآن (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو قرآن'})

print("radio_goftego")
radio_goftego=radio.query("channel == 'رادیو گفتگو'")
radio_goftego_visit=radio_goftego['تعداد بازدید'].sum()
radio_goftego_duration=radio_goftego['مدت بازدید'].sum()
radio_goftego_duration=round(radio_goftego_duration*60, 0)
radio_goftego_content=radio_goftego.copy()
radio_goftego_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_goftego_content=len(radio_goftego_content)
radio_goftego_pivot=radio_goftego.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_goftego_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_goftego_popular_visit=radio_goftego_pivot.iloc[0:10 , [0, 5]]
radio_goftego_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_goftego_popular_duration=radio_goftego_pivot.iloc[0:10 , [0, 4]]

radio_goftego_popular_visit = radio_goftego_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو گفتگو', 'نام برنامه': 'محتواهای پربازدید رادیو گفتگو'})
radio_goftego_popular_duration = radio_goftego_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو گفتگو (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو گفتگو'})

print("radio_maaref")
radio_maaref=radio.query("channel == 'رادیو معارف'")
radio_maaref_visit=radio_maaref['تعداد بازدید'].sum()
radio_maaref_duration=radio_maaref['مدت بازدید'].sum()
radio_maaref_duration=round(radio_maaref_duration*60, 0)
radio_maaref_content=radio_maaref.copy()
radio_maaref_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_maaref_content=len(radio_maaref_content)
radio_maaref_pivot=radio_maaref.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_maaref_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_maaref_popular_visit=radio_maaref_pivot.iloc[0:10 , [0, 5]]
radio_maaref_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_maaref_popular_duration=radio_maaref_pivot.iloc[0:10 , [0, 4]]

radio_maaref_popular_visit = radio_maaref_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو معارف', 'نام برنامه': 'محتواهای پربازدید رادیو معارف'})
radio_maaref_popular_duration = radio_maaref_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو معارف (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو معارف'})

print("radio_namayesh")
radio_namayesh=radio.query("channel == 'رادیو نمایش'")
radio_namayesh_visit=radio_namayesh['تعداد بازدید'].sum()
radio_namayesh_duration=radio_namayesh['مدت بازدید'].sum()
radio_namayesh_duration=round(radio_namayesh_duration*60, 0)
radio_namayesh_content=radio_namayesh.copy()
radio_namayesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_namayesh_content=len(radio_namayesh_content)
radio_namayesh_pivot=radio_namayesh.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_namayesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_namayesh_popular_visit=radio_namayesh_pivot.iloc[0:10 , [0, 5]]
radio_namayesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_namayesh_popular_duration=radio_namayesh_pivot.iloc[0:10 , [0, 4]]

radio_namayesh_popular_visit = radio_namayesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو نمایش', 'نام برنامه': 'محتواهای پربازدید رادیو نمایش'})
radio_namayesh_popular_duration = radio_namayesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو نمایش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو نمایش'})

print("radio_varzesh")
radio_varzesh=radio.query("channel == 'رادیو ورزش'")
radio_varzesh_visit=radio_varzesh['تعداد بازدید'].sum()
radio_varzesh_duration=radio_varzesh['مدت بازدید'].sum()
radio_varzesh_duration=round(radio_varzesh_duration*60, 0)
radio_varzesh_content=radio_varzesh.copy()
radio_varzesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
radio_varzesh_content=len(radio_varzesh_content)
radio_varzesh_pivot=radio_varzesh.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
radio_varzesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_varzesh_popular_visit=radio_varzesh_pivot.iloc[0:10 , [0, 5]]
radio_varzesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_varzesh_popular_duration=radio_varzesh_pivot.iloc[0:10 , [0, 4]]

radio_varzesh_popular_visit = radio_varzesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو ورزش', 'نام برنامه': 'محتواهای پربازدید رادیو ورزش'})
radio_varzesh_popular_duration = radio_varzesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو ورزش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو ورزش'})

print("dataframe radio channels")
radio_channels_statistics={'channel_name': ['رادیو اقتصاد', 'رادیو آوا', 'رادیو ایران', 'رادیو پیام', 'رادیو جوان',
                                     'رادیو سلامت', 'رادیو صبا', 'رادیو فرهنگ', 'رادیو قرآن', 'رادیو گفتگو',
                                     'رادیو معارف', 'رادیو نمایش', 'رادیو ورزش',],
       'channel_content': [radio_eghtesad_content, radio_ava_content, radio_iran_content, radio_payam_content, radio_javan_content,
                           radio_salamat_content, radio_saba_content, radio_farhang_content, radio_qoran_content, radio_goftego_content,
                           radio_maaref_content, radio_namayesh_content, radio_varzesh_content,],
       'channel_visit': [radio_eghtesad_visit, radio_ava_visit, radio_iran_visit, radio_payam_visit, radio_javan_visit,
                           radio_salamat_visit, radio_saba_visit, radio_farhang_visit, radio_qoran_visit, radio_goftego_visit,
                           radio_maaref_visit, radio_namayesh_visit, radio_varzesh_visit,],
       'channel_duration': [radio_eghtesad_duration, radio_ava_duration, radio_iran_duration, radio_payam_duration, radio_javan_duration,
                           radio_salamat_duration, radio_saba_duration, radio_farhang_duration, radio_qoran_duration, radio_goftego_duration,
                           radio_maaref_duration, radio_namayesh_duration, radio_varzesh_duration,],}
radio_channels_statistics=pd.DataFrame(radio_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
radio_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
radio_channels_statistics=radio_channels_statistics.rename(columns={'channel_content': 'نام شبکه', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})

radio_eghtesad_popular_visit.to_excel('busy/radio_eghtesad_popular_visit.xlsx')
radio_eghtesad_popular_duration.to_excel('busy/radio_eghtesad_popular_duration.xlsx')
radio_eghtesad_popular_visit=pd.read_excel('busy/radio_eghtesad_popular_visit.xlsx')
radio_eghtesad_popular_duration=pd.read_excel('busy/radio_eghtesad_popular_duration.xlsx')
del radio_eghtesad_popular_visit['Unnamed: 0']
del radio_eghtesad_popular_duration['Unnamed: 0']

radio_ava_popular_visit.to_excel('busy/radio_ava_popular_visit.xlsx')
radio_ava_popular_duration.to_excel('busy/radio_ava_popular_duration.xlsx')
radio_ava_popular_visit=pd.read_excel('busy/radio_ava_popular_visit.xlsx')
radio_ava_popular_duration=pd.read_excel('busy/radio_ava_popular_duration.xlsx')
del radio_ava_popular_visit['Unnamed: 0']
del radio_ava_popular_duration['Unnamed: 0']

radio_iran_popular_visit.to_excel('busy/radio_iran_popular_visit.xlsx')
radio_iran_popular_duration.to_excel('busy/radio_iran_popular_duration.xlsx')
radio_iran_popular_visit=pd.read_excel('busy/radio_iran_popular_visit.xlsx')
radio_iran_popular_duration=pd.read_excel('busy/radio_iran_popular_duration.xlsx')
del radio_iran_popular_visit['Unnamed: 0']
del radio_iran_popular_duration['Unnamed: 0']

radio_payam_popular_visit.to_excel('busy/radio_payam_popular_visit.xlsx')
radio_payam_popular_duration.to_excel('busy/radio_payam_popular_duration.xlsx')
radio_payam_popular_visit=pd.read_excel('busy/radio_payam_popular_visit.xlsx')
radio_payam_popular_duration=pd.read_excel('busy/radio_payam_popular_duration.xlsx')
del radio_payam_popular_visit['Unnamed: 0']
del radio_payam_popular_duration['Unnamed: 0']

radio_javan_popular_visit.to_excel('busy/radio_javan_popular_visit.xlsx')
radio_javan_popular_duration.to_excel('busy/radio_javan_popular_duration.xlsx')
radio_javan_popular_visit=pd.read_excel('busy/radio_javan_popular_visit.xlsx')
radio_javan_popular_duration=pd.read_excel('busy/radio_javan_popular_duration.xlsx')
del radio_javan_popular_visit['Unnamed: 0']
del radio_javan_popular_duration['Unnamed: 0']

radio_salamat_popular_visit.to_excel('busy/radio_salamat_popular_visit.xlsx')
radio_salamat_popular_duration.to_excel('busy/radio_salamat_popular_duration.xlsx')
radio_salamat_popular_visit=pd.read_excel('busy/radio_salamat_popular_visit.xlsx')
radio_salamat_popular_duration=pd.read_excel('busy/radio_salamat_popular_duration.xlsx')
del radio_salamat_popular_visit['Unnamed: 0']
del radio_salamat_popular_duration['Unnamed: 0']

radio_saba_popular_visit.to_excel('busy/radio_saba_popular_visit.xlsx')
radio_saba_popular_duration.to_excel('busy/radio_saba_popular_duration.xlsx')
radio_saba_popular_visit=pd.read_excel('busy/radio_saba_popular_visit.xlsx')
radio_saba_popular_duration=pd.read_excel('busy/radio_saba_popular_duration.xlsx')
del radio_saba_popular_visit['Unnamed: 0']
del radio_saba_popular_duration['Unnamed: 0']

radio_farhang_popular_visit.to_excel('busy/radio_farhang_popular_visit.xlsx')
radio_farhang_popular_duration.to_excel('busy/radio_farhang_popular_duration.xlsx')
radio_farhang_popular_visit=pd.read_excel('busy/radio_farhang_popular_visit.xlsx')
radio_farhang_popular_duration=pd.read_excel('busy/radio_farhang_popular_duration.xlsx')
del radio_farhang_popular_visit['Unnamed: 0']
del radio_farhang_popular_duration['Unnamed: 0']

radio_qoran_popular_visit.to_excel('busy/radio_qoran_popular_visit.xlsx')
radio_qoran_popular_duration.to_excel('busy/radio_qoran_popular_duration.xlsx')
radio_qoran_popular_visit=pd.read_excel('busy/radio_qoran_popular_visit.xlsx')
radio_qoran_popular_duration=pd.read_excel('busy/radio_qoran_popular_duration.xlsx')
del radio_qoran_popular_visit['Unnamed: 0']
del radio_qoran_popular_duration['Unnamed: 0']

radio_goftego_popular_visit.to_excel('busy/radio_goftego_popular_visit.xlsx')
radio_goftego_popular_duration.to_excel('busy/radio_goftego_popular_duration.xlsx')
radio_goftego_popular_visit=pd.read_excel('busy/radio_goftego_popular_visit.xlsx')
radio_goftego_popular_duration=pd.read_excel('busy/radio_goftego_popular_duration.xlsx')
del radio_goftego_popular_visit['Unnamed: 0']
del radio_goftego_popular_duration['Unnamed: 0']

radio_maaref_popular_visit.to_excel('busy/radio_maaref_popular_visit.xlsx')
radio_maaref_popular_duration.to_excel('busy/radio_maaref_popular_duration.xlsx')
radio_maaref_popular_visit=pd.read_excel('busy/radio_maaref_popular_visit.xlsx')
radio_maaref_popular_duration=pd.read_excel('busy/radio_maaref_popular_duration.xlsx')
del radio_maaref_popular_visit['Unnamed: 0']
del radio_maaref_popular_duration['Unnamed: 0']

radio_namayesh_popular_visit.to_excel('busy/radio_namayesh_popular_visit.xlsx')
radio_namayesh_popular_duration.to_excel('busy/radio_namayesh_popular_duration.xlsx')
radio_namayesh_popular_visit=pd.read_excel('busy/radio_namayesh_popular_visit.xlsx')
radio_namayesh_popular_duration=pd.read_excel('busy/radio_namayesh_popular_duration.xlsx')
del radio_namayesh_popular_visit['Unnamed: 0']
del radio_namayesh_popular_duration['Unnamed: 0']

radio_varzesh_popular_visit.to_excel('busy/radio_varzesh_popular_visit.xlsx')
radio_varzesh_popular_duration.to_excel('busy/radio_varzesh_popular_duration.xlsx')
radio_varzesh_popular_visit=pd.read_excel('busy/radio_varzesh_popular_visit.xlsx')
radio_varzesh_popular_duration=pd.read_excel('busy/radio_varzesh_popular_duration.xlsx')
del radio_varzesh_popular_visit['Unnamed: 0']
del radio_varzesh_popular_duration['Unnamed: 0']

radio_channels_popular_content=pd.DataFrame()
radio_channels_popular_content=pd.concat([radio_eghtesad_popular_visit, radio_eghtesad_popular_duration,
                                          radio_ava_popular_visit, radio_ava_popular_duration,
                                          radio_iran_popular_visit, radio_iran_popular_duration,
                                          radio_payam_popular_visit, radio_payam_popular_duration,
                                          radio_javan_popular_visit, radio_javan_popular_duration,
                                          radio_salamat_popular_visit, radio_salamat_popular_duration,
                                          radio_saba_popular_visit, radio_saba_popular_duration,
                                          radio_farhang_popular_visit, radio_farhang_popular_duration,
                                          radio_qoran_popular_visit, radio_qoran_popular_duration,
                                          radio_goftego_popular_visit, radio_goftego_popular_duration,
                                          radio_maaref_popular_visit, radio_maaref_popular_duration,
                                          radio_namayesh_popular_visit, radio_namayesh_popular_duration,
                                          radio_varzesh_popular_visit, radio_varzesh_popular_duration,],axis=1)

writer = pd.ExcelWriter('output/آمار رادیو.xlsx', engine='xlsxwriter')
radio_channels_statistics.to_excel(writer, 'آمار شبکه های رادیو')
radio_channels_popular_content.to_excel(writer, 'محتواهای پربازدید')
writer.save()

print("End radio")
################################# OSTANI ########################################
print("start ostani")

ostani_all=ostani.copy()
ostani_all_pivot=ostani_all.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
ostani_all_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_all_popular_visit=ostani_all_pivot.iloc[0:10 , [0, 4]]
ostani_all_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_all_popular_duration=ostani_all_pivot.iloc[0:10 , [0, 3]]


print("ostani_abadan")
ostani_abadan=ostani.query("channel == 'استانی آبادان'")
ostani_abadan_visit=ostani_abadan['تعداد بازدید'].sum()
ostani_abadan_duration=ostani_abadan['مدت بازدید'].sum()
ostani_abadan_duration=round(ostani_abadan_duration*60, 0)
ostani_abadan_content=ostani_abadan.copy()
ostani_abadan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_abadan_content=len(ostani_abadan_content)
ostani_abadan_pivot=ostani_abadan.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_abadan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_abadan_popular_visit=ostani_abadan_pivot.iloc[0:10 , [0, 5]]
ostani_abadan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_abadan_popular_duration=ostani_abadan_pivot.iloc[0:10 , [0, 4]]

ostani_abadan_popular_visit = ostani_abadan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی آبادان', 'نام برنامه': 'محتواهای پربازدید استانی آبادان'})
ostani_abadan_popular_duration = ostani_abadan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی آبادان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی آبادان'})

print("ostani_azarbayjan_gharbi")
ostani_azarbayjan_gharbi=ostani.query("channel == 'استانی آذربایجان غربی'")
ostani_azarbayjan_gharbi_visit=ostani_azarbayjan_gharbi['تعداد بازدید'].sum()
ostani_azarbayjan_gharbi_duration=ostani_azarbayjan_gharbi['مدت بازدید'].sum()
ostani_azarbayjan_gharbi_duration=round(ostani_azarbayjan_gharbi_duration*60, 0)
ostani_azarbayjan_gharbi_content=ostani_azarbayjan_gharbi.copy()
ostani_azarbayjan_gharbi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_azarbayjan_gharbi_content=len(ostani_azarbayjan_gharbi_content)
ostani_azarbayjan_gharbi_pivot=ostani_azarbayjan_gharbi.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_azarbayjan_gharbi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_azarbayjan_gharbi_popular_visit=ostani_azarbayjan_gharbi_pivot.iloc[0:10 , [0, 5]]
ostani_azarbayjan_gharbi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_azarbayjan_gharbi_popular_duration=ostani_azarbayjan_gharbi_pivot.iloc[0:10 , [0, 4]]

ostani_azarbayjan_gharbi_popular_visit = ostani_azarbayjan_gharbi_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی آذربایجان غربی', 'نام برنامه': 'محتواهای پربازدید استانی آذربایجان غربی'})
ostani_azarbayjan_gharbi_popular_duration = ostani_azarbayjan_gharbi_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی آذربایجان غربی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی آذربایجان غربی'})

print("ostani_esfahan")
ostani_esfahan=ostani.query("channel == 'استانی اصفهان'")
ostani_esfahan_visit=ostani_esfahan['تعداد بازدید'].sum()
ostani_esfahan_duration=ostani_esfahan['مدت بازدید'].sum()
ostani_esfahan_duration=round(ostani_esfahan_duration*60, 0)
ostani_esfahan_content=ostani_esfahan.copy()
ostani_esfahan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_esfahan_content=len(ostani_esfahan_content)
ostani_esfahan_pivot=ostani_esfahan.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_esfahan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_esfahan_popular_visit=ostani_esfahan_pivot.iloc[0:10 , [0, 5]]
ostani_esfahan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_esfahan_popular_duration=ostani_esfahan_pivot.iloc[0:10 , [0, 4]]

ostani_esfahan_popular_visit = ostani_esfahan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی اصفهان', 'نام برنامه': 'محتواهای پربازدید استانی اصفهان'})
ostani_esfahan_popular_duration = ostani_esfahan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی اصفهان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی اصفهان'})

print("ostani_aflak")
ostani_aflak=ostani.query("channel == 'استانی افلاک'")
ostani_aflak_visit=ostani_aflak['تعداد بازدید'].sum()
ostani_aflak_duration=ostani_aflak['مدت بازدید'].sum()
ostani_aflak_duration=round(ostani_aflak_duration*60, 0)
ostani_aflak_content=ostani_aflak.copy()
ostani_aflak_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_aflak_content=len(ostani_aflak_content)
ostani_aflak_pivot=ostani_aflak.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_aflak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_aflak_popular_visit=ostani_aflak_pivot.iloc[0:10 , [0, 5]]
ostani_aflak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_aflak_popular_duration=ostani_aflak_pivot.iloc[0:10 , [0, 4]]

ostani_aflak_popular_visit = ostani_aflak_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی افلاک', 'نام برنامه': 'محتواهای پربازدید استانی افلاک'})
ostani_aflak_popular_duration = ostani_aflak_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی افلاک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی افلاک'})

print("ostani_alborz")
ostani_alborz=ostani.query("channel == 'استانی البرز'")
ostani_alborz_visit=ostani_alborz['تعداد بازدید'].sum()
ostani_alborz_duration=ostani_alborz['مدت بازدید'].sum()
ostani_alborz_duration=round(ostani_alborz_duration*60, 0)
ostani_alborz_content=ostani_alborz.copy()
ostani_alborz_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_alborz_content=len(ostani_alborz_content)
ostani_alborz_pivot=ostani_alborz.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_alborz_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_alborz_popular_visit=ostani_alborz_pivot.iloc[0:10 , [0, 5]]
ostani_alborz_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_alborz_popular_duration=ostani_alborz_pivot.iloc[0:10 , [0, 4]]

ostani_alborz_popular_visit = ostani_alborz_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی البرز', 'نام برنامه': 'محتواهای پربازدید استانی البرز'})
ostani_alborz_popular_duration = ostani_alborz_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی البرز (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی البرز'})

print("ostani_ilam")
ostani_ilam=ostani.query("channel == 'استانی ایلام'")
ostani_ilam_visit=ostani_ilam['تعداد بازدید'].sum()
ostani_ilam_duration=ostani_ilam['مدت بازدید'].sum()
ostani_ilam_duration=round(ostani_ilam_duration*60, 0)
ostani_ilam_content=ostani_ilam.copy()
ostani_ilam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_ilam_content=len(ostani_ilam_content)
ostani_ilam_pivot=ostani_ilam.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_ilam_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_ilam_popular_visit=ostani_ilam_pivot.iloc[0:10 , [0, 5]]
ostani_ilam_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_ilam_popular_duration=ostani_ilam_pivot.iloc[0:10 , [0, 4]]

ostani_ilam_popular_visit = ostani_ilam_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی ایلام', 'نام برنامه': 'محتواهای پربازدید استانی ایلام'})
ostani_ilam_popular_duration = ostani_ilam_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی ایلام (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی ایلام'})

print("ostani_baran")
ostani_baran=ostani.query("channel == 'استانی باران'")
ostani_baran_visit=ostani_baran['تعداد بازدید'].sum()
ostani_baran_duration=ostani_baran['مدت بازدید'].sum()
ostani_baran_duration=round(ostani_baran_duration*60, 0)
ostani_baran_content=ostani_baran.copy()
ostani_baran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_baran_content=len(ostani_baran_content)
ostani_baran_pivot=ostani_baran.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_baran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_baran_popular_visit=ostani_baran_pivot.iloc[0:10 , [0, 5]]
ostani_baran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_baran_popular_duration=ostani_baran_pivot.iloc[0:10 , [0, 4]]

ostani_baran_popular_visit = ostani_baran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی باران', 'نام برنامه': 'محتواهای پربازدید استانی باران'})
ostani_baran_popular_duration = ostani_baran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی باران (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی باران'})

print("ostani_boshehr")
ostani_boshehr=ostani.query("channel == 'استانی بوشهر'")
ostani_boshehr_visit=ostani_boshehr['تعداد بازدید'].sum()
ostani_boshehr_duration=ostani_boshehr['مدت بازدید'].sum()
ostani_boshehr_duration=round(ostani_boshehr_duration*60, 0)
ostani_boshehr_content=ostani_boshehr.copy()
ostani_boshehr_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_boshehr_content=len(ostani_boshehr_content)
ostani_boshehr_pivot=ostani_boshehr.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_boshehr_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_boshehr_popular_visit=ostani_boshehr_pivot.iloc[0:10 , [0, 5]]
ostani_boshehr_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_boshehr_popular_duration=ostani_boshehr_pivot.iloc[0:10 , [0, 4]]

ostani_boshehr_popular_visit = ostani_boshehr_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی بوشهر', 'نام برنامه': 'محتواهای پربازدید استانی بوشهر'})
ostani_boshehr_popular_duration = ostani_boshehr_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی بوشهر (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی بوشهر'})

print("ostani_taban")
ostani_taban=ostani.query("channel == 'استانی تابان'")
ostani_taban_visit=ostani_taban['تعداد بازدید'].sum()
ostani_taban_duration=ostani_taban['مدت بازدید'].sum()
ostani_taban_duration=round(ostani_taban_duration*60, 0)
ostani_taban_content=ostani_taban.copy()
ostani_taban_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_taban_content=len(ostani_taban_content)
ostani_taban_pivot=ostani_taban.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_taban_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_taban_popular_visit=ostani_taban_pivot.iloc[0:10 , [0, 5]]
ostani_taban_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_taban_popular_duration=ostani_taban_pivot.iloc[0:10 , [0, 4]]

ostani_taban_popular_visit = ostani_taban_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی تابان', 'نام برنامه': 'محتواهای پربازدید استانی تابان'})
ostani_taban_popular_duration = ostani_taban_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی تابان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی تابان'})

print("ostani_khorasan_razavi")
ostani_khorasan_razavi=ostani.query("channel == 'استانی خراسان رضوی'")
ostani_khorasan_razavi_visit=ostani_khorasan_razavi['تعداد بازدید'].sum()
ostani_khorasan_razavi_duration=ostani_khorasan_razavi['مدت بازدید'].sum()
ostani_khorasan_razavi_duration=round(ostani_khorasan_razavi_duration*60, 0)
ostani_khorasan_razavi_content=ostani_khorasan_razavi.copy()
ostani_khorasan_razavi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_khorasan_razavi_content=len(ostani_khorasan_razavi_content)
ostani_khorasan_razavi_pivot=ostani_khorasan_razavi.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_khorasan_razavi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_khorasan_razavi_popular_visit=ostani_khorasan_razavi_pivot.iloc[0:10 , [0, 5]]
ostani_khorasan_razavi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_khorasan_razavi_popular_duration=ostani_khorasan_razavi_pivot.iloc[0:10 , [0, 4]]

ostani_khorasan_razavi_popular_visit = ostani_khorasan_razavi_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی خراسان رضوی', 'نام برنامه': 'محتواهای پربازدید استانی خراسان رضوی'})
ostani_khorasan_razavi_popular_duration = ostani_khorasan_razavi_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی خراسان رضوی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی خراسان رضوی'})

print("ostani_khozestan")
ostani_khozestan=ostani.query("channel == 'استانی خوزستان'")
ostani_khozestan_visit=ostani_khozestan['تعداد بازدید'].sum()
ostani_khozestan_duration=ostani_khozestan['مدت بازدید'].sum()
ostani_khozestan_duration=round(ostani_khozestan_duration*60, 0)
ostani_khozestan_content=ostani_khozestan.copy()
ostani_khozestan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_khozestan_content=len(ostani_khozestan_content)
ostani_khozestan_pivot=ostani_khozestan.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_khozestan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_khozestan_popular_visit=ostani_khozestan_pivot.iloc[0:10 , [0, 5]]
ostani_khozestan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_khozestan_popular_duration=ostani_khozestan_pivot.iloc[0:10 , [0, 4]]

ostani_khozestan_popular_visit = ostani_khozestan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی خوزستان', 'نام برنامه': 'محتواهای پربازدید استانی خوزستان'})
ostani_khozestan_popular_duration = ostani_khozestan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی خوزستان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی خوزستان'})

print("ostani_dena")
ostani_dena=ostani.query("channel == 'استانی دنا'")
ostani_dena_visit=ostani_dena['تعداد بازدید'].sum()
ostani_dena_duration=ostani_dena['مدت بازدید'].sum()
ostani_dena_duration=round(ostani_dena_duration*60, 0)
ostani_dena_content=ostani_dena.copy()
ostani_dena_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_dena_content=len(ostani_dena_content)
ostani_dena_pivot=ostani_dena.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_dena_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_dena_popular_visit=ostani_dena_pivot.iloc[0:10 , [0, 5]]
ostani_dena_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_dena_popular_duration=ostani_dena_pivot.iloc[0:10 , [0, 4]]

ostani_dena_popular_visit = ostani_dena_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی دنا', 'نام برنامه': 'محتواهای پربازدید استانی دنا'})
ostani_dena_popular_duration = ostani_dena_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی دنا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی دنا'})

print("ostani_sabalan")
ostani_sabalan=ostani.query("channel == 'استانی سبلان'")
ostani_sabalan_visit=ostani_sabalan['تعداد بازدید'].sum()
ostani_sabalan_duration=ostani_sabalan['مدت بازدید'].sum()
ostani_sabalan_duration=round(ostani_sabalan_duration*60, 0)
ostani_sabalan_content=ostani_sabalan.copy()
ostani_sabalan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_sabalan_content=len(ostani_sabalan_content)
ostani_sabalan_pivot=ostani_sabalan.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_sabalan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_sabalan_popular_visit=ostani_sabalan_pivot.iloc[0:10 , [0, 5]]
ostani_sabalan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_sabalan_popular_duration=ostani_sabalan_pivot.iloc[0:10 , [0, 4]]

ostani_sabalan_popular_visit = ostani_sabalan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی سبلان', 'نام برنامه': 'محتواهای پربازدید استانی سبلان'})
ostani_sabalan_popular_duration = ostani_sabalan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی سبلان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی سبلان'})

print("ostani_sahand")
ostani_sahand=ostani.query("channel == 'استانی سهند'")
ostani_sahand_visit=ostani_sahand['تعداد بازدید'].sum()
ostani_sahand_duration=ostani_sahand['مدت بازدید'].sum()
ostani_sahand_duration=round(ostani_sahand_duration*60, 0)
ostani_sahand_content=ostani_sahand.copy()
ostani_sahand_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_sahand_content=len(ostani_sahand_content)
ostani_sahand_pivot=ostani_sahand.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_sahand_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_sahand_popular_visit=ostani_sahand_pivot.iloc[0:10 , [0, 5]]
ostani_sahand_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_sahand_popular_duration=ostani_sahand_pivot.iloc[0:10 , [0, 4]]

ostani_sahand_popular_visit = ostani_sahand_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی سهند', 'نام برنامه': 'محتواهای پربازدید استانی سهند'})
ostani_sahand_popular_duration = ostani_sahand_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی سهند (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی سهند'})

print("ostani_fars")
ostani_fars=ostani.query("channel == 'استانی فارس'")
ostani_fars_visit=ostani_fars['تعداد بازدید'].sum()
ostani_fars_duration=ostani_fars['مدت بازدید'].sum()
ostani_fars_duration=round(ostani_fars_duration*60, 0)
ostani_fars_content=ostani_fars.copy()
ostani_fars_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_fars_content=len(ostani_fars_content)
ostani_fars_pivot=ostani_fars.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_fars_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_fars_popular_visit=ostani_fars_pivot.iloc[0:10 , [0, 5]]
ostani_fars_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_fars_popular_duration=ostani_fars_pivot.iloc[0:10 , [0, 4]]

ostani_fars_popular_visit = ostani_fars_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی فارس', 'نام برنامه': 'محتواهای پربازدید استانی فارس'})
ostani_fars_popular_duration = ostani_fars_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی فارس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی فارس'})

print("ostani_ghazvin")
ostani_ghazvin=ostani.query("channel == 'استانی قزوین'")
ostani_ghazvin_visit=ostani_ghazvin['تعداد بازدید'].sum()
ostani_ghazvin_duration=ostani_ghazvin['مدت بازدید'].sum()
ostani_ghazvin_duration=round(ostani_ghazvin_duration*60, 0)
ostani_ghazvin_content=ostani_ghazvin.copy()
ostani_ghazvin_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_ghazvin_content=len(ostani_ghazvin_content)
ostani_ghazvin_pivot=ostani_ghazvin.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_ghazvin_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_ghazvin_popular_visit=ostani_ghazvin_pivot.iloc[0:10 , [0, 5]]
ostani_ghazvin_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_ghazvin_popular_duration=ostani_ghazvin_pivot.iloc[0:10 , [0, 4]]

ostani_ghazvin_popular_visit = ostani_ghazvin_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی قزوین', 'نام برنامه': 'محتواهای پربازدید استانی قزوین'})
ostani_ghazvin_popular_duration = ostani_ghazvin_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی قزوین (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی قزوین'})

print("ostani_kordestan")
ostani_kordestan=ostani.query("channel == 'استانی کردستان'")
ostani_kordestan_visit=ostani_kordestan['تعداد بازدید'].sum()
ostani_kordestan_duration=ostani_kordestan['مدت بازدید'].sum()
ostani_kordestan_duration=round(ostani_kordestan_duration*60, 0)
ostani_kordestan_content=ostani_kordestan.copy()
ostani_kordestan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_kordestan_content=len(ostani_kordestan_content)
ostani_kordestan_pivot=ostani_kordestan.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_kordestan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_kordestan_popular_visit=ostani_kordestan_pivot.iloc[0:10 , [0, 5]]
ostani_kordestan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_kordestan_popular_duration=ostani_kordestan_pivot.iloc[0:10 , [0, 4]]

ostani_kordestan_popular_visit = ostani_kordestan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی کردستان', 'نام برنامه': 'محتواهای پربازدید استانی کردستان'})
ostani_kordestan_popular_duration = ostani_kordestan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی کردستان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی کردستان'})

print("ostani_kermanshah")
ostani_kermanshah=ostani.query("channel == 'استانی کرمانشاه'")
ostani_kermanshah_visit=ostani_kermanshah['تعداد بازدید'].sum()
ostani_kermanshah_duration=ostani_kermanshah['مدت بازدید'].sum()
ostani_kermanshah_duration=round(ostani_kermanshah_duration*60, 0)
ostani_kermanshah_content=ostani_kermanshah.copy()
ostani_kermanshah_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_kermanshah_content=len(ostani_kermanshah_content)
ostani_kermanshah_pivot=ostani_kermanshah.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_kermanshah_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_kermanshah_popular_visit=ostani_kermanshah_pivot.iloc[0:10 , [0, 5]]
ostani_kermanshah_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_kermanshah_popular_duration=ostani_kermanshah_pivot.iloc[0:10 , [0, 4]]

ostani_kermanshah_popular_visit = ostani_kermanshah_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی کرمانشاه', 'نام برنامه': 'محتواهای پربازدید استانی کرمانشاه'})
ostani_kermanshah_popular_duration = ostani_kermanshah_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی کرمانشاه (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی کرمانشاه'})

print("ostani_kish")
ostani_kish=ostani.query("channel == 'استانی کیش'")
ostani_kish_visit=ostani_kish['تعداد بازدید'].sum()
ostani_kish_duration=ostani_kish['مدت بازدید'].sum()
ostani_kish_duration=round(ostani_kish_duration*60, 0)
ostani_kish_content=ostani_kish.copy()
ostani_kish_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_kish_content=len(ostani_kish_content)
ostani_kish_pivot=ostani_kish.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_kish_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_kish_popular_visit=ostani_kish_pivot.iloc[0:10 , [0, 5]]
ostani_kish_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_kish_popular_duration=ostani_kish_pivot.iloc[0:10 , [0, 4]]

ostani_kish_popular_visit = ostani_kish_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی کیش', 'نام برنامه': 'محتواهای پربازدید استانی کیش'})
ostani_kish_popular_duration = ostani_kish_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی کیش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی کیش'})

print("ostani_mazandaran")
ostani_mazandaran=ostani.query("channel == 'استانی مازندران'")
ostani_mazandaran_visit=ostani_mazandaran['تعداد بازدید'].sum()
ostani_mazandaran_duration=ostani_mazandaran['مدت بازدید'].sum()
ostani_mazandaran_duration=round(ostani_mazandaran_duration*60, 0)
ostani_mazandaran_content=ostani_mazandaran.copy()
ostani_mazandaran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_mazandaran_content=len(ostani_mazandaran_content)
ostani_mazandaran_pivot=ostani_mazandaran.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_mazandaran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_mazandaran_popular_visit=ostani_mazandaran_pivot.iloc[0:10 , [0, 5]]
ostani_mazandaran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_mazandaran_popular_duration=ostani_mazandaran_pivot.iloc[0:10 , [0, 4]]

ostani_mazandaran_popular_visit = ostani_mazandaran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی مازندران', 'نام برنامه': 'محتواهای پربازدید استانی مازندران'})
ostani_mazandaran_popular_duration = ostani_mazandaran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی مازندران (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی مازندران'})

print("ostani_hamedan")
ostani_hamedan=ostani.query("channel == 'استانی همدان'")
ostani_hamedan_visit=ostani_hamedan['تعداد بازدید'].sum()
ostani_hamedan_duration=ostani_hamedan['مدت بازدید'].sum()
ostani_hamedan_duration=round(ostani_hamedan_duration*60, 0)
ostani_hamedan_content=ostani_hamedan.copy()
ostani_hamedan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ostani_hamedan_content=len(ostani_hamedan_content)
ostani_hamedan_pivot=ostani_hamedan.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ostani_hamedan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_hamedan_popular_visit=ostani_hamedan_pivot.iloc[0:10 , [0, 5]]
ostani_hamedan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_hamedan_popular_duration=ostani_hamedan_pivot.iloc[0:10 , [0, 4]]

ostani_hamedan_popular_visit = ostani_hamedan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی همدان', 'نام برنامه': 'محتواهای پربازدید استانی همدان'})
ostani_hamedan_popular_duration = ostani_hamedan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی همدان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی همدان'})

print("dataframe ostani channels")
ostani_channels_statistics={'channel_name': ['استانی آبادان', 'استانی آذربایجان غربی', 'استانی اصفهان', 'استانی افلاک', 'استانی البرز',
                                     'استانی ایلام', 'استانی باران', 'استانی بوشهر', 'استانی تابان', 'استانی خراسان رضوی', 'استانی خوزستان',
                                     'استانی دنا', 'استانی سبلان', 'استانی سهند','استانی فارس', 'استانی قزوین', 'استانی کردستان',
                                     'استانی کرمانشاه', 'استانی کیش', 'استانی مازندران', 'استانی همدان'],
       'channel_content': [ostani_abadan_content, ostani_azarbayjan_gharbi_content, ostani_esfahan_content, ostani_aflak_content, ostani_alborz_content, ostani_ilam_content,
                           ostani_baran_content, ostani_boshehr_content, ostani_taban_content, ostani_khorasan_razavi_content, ostani_khozestan_content, ostani_dena_content,
                           ostani_sabalan_content, ostani_sahand_content, ostani_fars_content,ostani_ghazvin_content,ostani_kordestan_content,
                           ostani_kermanshah_content, ostani_kish_content, ostani_mazandaran_content, ostani_hamedan_content],
       'channel_visit': [ostani_abadan_visit, ostani_azarbayjan_gharbi_visit, ostani_esfahan_visit, ostani_aflak_visit, ostani_alborz_visit, ostani_ilam_visit,
                           ostani_baran_visit, ostani_boshehr_visit, ostani_taban_visit, ostani_khorasan_razavi_visit, ostani_khozestan_visit, ostani_dena_visit,
                           ostani_sabalan_visit, ostani_sahand_visit, ostani_fars_visit,ostani_ghazvin_visit,ostani_kordestan_visit,
                           ostani_kermanshah_visit, ostani_kish_visit, ostani_mazandaran_visit, ostani_hamedan_visit],
       'channel_duration': [ostani_abadan_duration, ostani_azarbayjan_gharbi_duration, ostani_esfahan_duration, ostani_aflak_duration, ostani_alborz_duration, ostani_ilam_duration,
                           ostani_baran_duration, ostani_boshehr_duration, ostani_taban_duration, ostani_khorasan_razavi_duration, ostani_khozestan_duration, ostani_dena_duration,
                           ostani_sabalan_duration, ostani_sahand_duration, ostani_fars_duration,ostani_ghazvin_duration,ostani_kordestan_duration,
                           ostani_kermanshah_duration, ostani_kish_duration, ostani_mazandaran_duration, ostani_hamedan_duration],}
ostani_channels_statistics=pd.DataFrame(ostani_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
ostani_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
ostani_channels_statistics=ostani_channels_statistics.rename(columns={'channel_content': 'نام شبکه', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})

ostani_abadan_popular_visit.to_excel('busy/ostani_abadan_popular_visit.xlsx')
ostani_abadan_popular_duration.to_excel('busy/ostani_abadan_popular_duration.xlsx')
ostani_abadan_popular_visit=pd.read_excel('busy/ostani_abadan_popular_visit.xlsx')
ostani_abadan_popular_duration=pd.read_excel('busy/ostani_abadan_popular_duration.xlsx')
del ostani_abadan_popular_visit['Unnamed: 0']
del ostani_abadan_popular_duration['Unnamed: 0']

ostani_azarbayjan_gharbi_popular_visit.to_excel('busy/ostani_azarbayjan_gharbi_popular_visit.xlsx')
ostani_azarbayjan_gharbi_popular_duration.to_excel('busy/ostani_azarbayjan_gharbi_popular_duration.xlsx')
ostani_azarbayjan_gharbi_popular_visit=pd.read_excel('busy/ostani_azarbayjan_gharbi_popular_visit.xlsx')
ostani_azarbayjan_gharbi_popular_duration=pd.read_excel('busy/ostani_azarbayjan_gharbi_popular_duration.xlsx')
del ostani_azarbayjan_gharbi_popular_visit['Unnamed: 0']
del ostani_azarbayjan_gharbi_popular_duration['Unnamed: 0']

ostani_esfahan_popular_visit.to_excel('busy/ostani_esfahan_popular_visit.xlsx')
ostani_esfahan_popular_duration.to_excel('busy/ostani_esfahan_popular_duration.xlsx')
ostani_esfahan_popular_visit=pd.read_excel('busy/ostani_esfahan_popular_visit.xlsx')
ostani_esfahan_popular_duration=pd.read_excel('busy/ostani_esfahan_popular_duration.xlsx')
del ostani_esfahan_popular_visit['Unnamed: 0']
del ostani_esfahan_popular_duration['Unnamed: 0']

ostani_aflak_popular_visit.to_excel('busy/ostani_aflak_popular_visit.xlsx')
ostani_aflak_popular_duration.to_excel('busy/ostani_aflak_popular_duration.xlsx')
ostani_aflak_popular_visit=pd.read_excel('busy/ostani_aflak_popular_visit.xlsx')
ostani_aflak_popular_duration=pd.read_excel('busy/ostani_aflak_popular_duration.xlsx')
del ostani_aflak_popular_visit['Unnamed: 0']
del ostani_aflak_popular_duration['Unnamed: 0']

ostani_alborz_popular_visit.to_excel('busy/ostani_alborz_popular_visit.xlsx')
ostani_alborz_popular_duration.to_excel('busy/ostani_alborz_popular_duration.xlsx')
ostani_alborz_popular_visit=pd.read_excel('busy/ostani_alborz_popular_visit.xlsx')
ostani_alborz_popular_duration=pd.read_excel('busy/ostani_alborz_popular_duration.xlsx')
del ostani_alborz_popular_visit['Unnamed: 0']
del ostani_alborz_popular_duration['Unnamed: 0']

ostani_ilam_popular_visit.to_excel('busy/ostani_ilam_popular_visit.xlsx')
ostani_ilam_popular_duration.to_excel('busy/ostani_ilam_popular_duration.xlsx')
ostani_ilam_popular_visit=pd.read_excel('busy/ostani_ilam_popular_visit.xlsx')
ostani_ilam_popular_duration=pd.read_excel('busy/ostani_ilam_popular_duration.xlsx')
del ostani_ilam_popular_visit['Unnamed: 0']
del ostani_ilam_popular_duration['Unnamed: 0']

ostani_baran_popular_visit.to_excel('busy/ostani_baran_popular_visit.xlsx')
ostani_baran_popular_duration.to_excel('busy/ostani_baran_popular_duration.xlsx')
ostani_baran_popular_visit=pd.read_excel('busy/ostani_baran_popular_visit.xlsx')
ostani_baran_popular_duration=pd.read_excel('busy/ostani_baran_popular_duration.xlsx')
del ostani_baran_popular_visit['Unnamed: 0']
del ostani_baran_popular_duration['Unnamed: 0']

ostani_boshehr_popular_visit.to_excel('busy/ostani_boshehr_popular_visit.xlsx')
ostani_boshehr_popular_duration.to_excel('busy/ostani_boshehr_popular_duration.xlsx')
ostani_boshehr_popular_visit=pd.read_excel('busy/ostani_boshehr_popular_visit.xlsx')
ostani_boshehr_popular_duration=pd.read_excel('busy/ostani_boshehr_popular_duration.xlsx')
del ostani_boshehr_popular_visit['Unnamed: 0']
del ostani_boshehr_popular_duration['Unnamed: 0']

ostani_taban_popular_visit.to_excel('busy/ostani_taban_popular_visit.xlsx')
ostani_taban_popular_duration.to_excel('busy/ostani_taban_popular_duration.xlsx')
ostani_taban_popular_visit=pd.read_excel('busy/ostani_taban_popular_visit.xlsx')
ostani_taban_popular_duration=pd.read_excel('busy/ostani_taban_popular_duration.xlsx')
del ostani_taban_popular_visit['Unnamed: 0']
del ostani_taban_popular_duration['Unnamed: 0']

ostani_khorasan_razavi_popular_visit.to_excel('busy/ostani_khorasan_razavi_popular_visit.xlsx')
ostani_khorasan_razavi_popular_duration.to_excel('busy/ostani_khorasan_razavi_popular_duration.xlsx')
ostani_khorasan_razavi_popular_visit=pd.read_excel('busy/ostani_khorasan_razavi_popular_visit.xlsx')
ostani_khorasan_razavi_popular_duration=pd.read_excel('busy/ostani_khorasan_razavi_popular_duration.xlsx')
del ostani_khorasan_razavi_popular_visit['Unnamed: 0']
del ostani_khorasan_razavi_popular_duration['Unnamed: 0']

ostani_khozestan_popular_visit.to_excel('busy/ostani_khozestan_popular_visit.xlsx')
ostani_khozestan_popular_duration.to_excel('busy/ostani_khozestan_popular_duration.xlsx')
ostani_khozestan_popular_visit=pd.read_excel('busy/ostani_khozestan_popular_visit.xlsx')
ostani_khozestan_popular_duration=pd.read_excel('busy/ostani_khozestan_popular_duration.xlsx')
del ostani_khozestan_popular_visit['Unnamed: 0']
del ostani_khozestan_popular_duration['Unnamed: 0']

ostani_dena_popular_visit.to_excel('busy/ostani_dena_popular_visit.xlsx')
ostani_dena_popular_duration.to_excel('busy/ostani_dena_popular_duration.xlsx')
ostani_dena_popular_visit=pd.read_excel('busy/ostani_dena_popular_visit.xlsx')
ostani_dena_popular_duration=pd.read_excel('busy/ostani_dena_popular_duration.xlsx')
del ostani_dena_popular_visit['Unnamed: 0']
del ostani_dena_popular_duration['Unnamed: 0']

ostani_sabalan_popular_visit.to_excel('busy/ostani_sabalan_popular_visit.xlsx')
ostani_sabalan_popular_duration.to_excel('busy/ostani_sabalan_popular_duration.xlsx')
ostani_sabalan_popular_visit=pd.read_excel('busy/ostani_sabalan_popular_visit.xlsx')
ostani_sabalan_popular_duration=pd.read_excel('busy/ostani_sabalan_popular_duration.xlsx')
del ostani_sabalan_popular_visit['Unnamed: 0']
del ostani_sabalan_popular_duration['Unnamed: 0']

ostani_sahand_popular_visit.to_excel('busy/ostani_sahand_popular_visit.xlsx')
ostani_sahand_popular_duration.to_excel('busy/ostani_sahand_popular_duration.xlsx')
ostani_sahand_popular_visit=pd.read_excel('busy/ostani_sahand_popular_visit.xlsx')
ostani_sahand_popular_duration=pd.read_excel('busy/ostani_sahand_popular_duration.xlsx')
del ostani_sahand_popular_visit['Unnamed: 0']
del ostani_sahand_popular_duration['Unnamed: 0']

ostani_fars_popular_visit.to_excel('busy/ostani_fars_popular_visit.xlsx')
ostani_fars_popular_duration.to_excel('busy/ostani_fars_popular_duration.xlsx')
ostani_fars_popular_visit=pd.read_excel('busy/ostani_fars_popular_visit.xlsx')
ostani_fars_popular_duration=pd.read_excel('busy/ostani_fars_popular_duration.xlsx')
del ostani_fars_popular_visit['Unnamed: 0']
del ostani_fars_popular_duration['Unnamed: 0']

ostani_ghazvin_popular_visit.to_excel('busy/ostani_ghazvin_popular_visit.xlsx')
ostani_ghazvin_popular_duration.to_excel('busy/ostani_ghazvin_popular_duration.xlsx')
ostani_ghazvin_popular_visit=pd.read_excel('busy/ostani_ghazvin_popular_visit.xlsx')
ostani_ghazvin_popular_duration=pd.read_excel('busy/ostani_ghazvin_popular_duration.xlsx')
del ostani_ghazvin_popular_visit['Unnamed: 0']
del ostani_ghazvin_popular_duration['Unnamed: 0']

ostani_kordestan_popular_visit.to_excel('busy/ostani_kordestan_popular_visit.xlsx')
ostani_kordestan_popular_duration.to_excel('busy/ostani_kordestan_popular_duration.xlsx')
ostani_kordestan_popular_visit=pd.read_excel('busy/ostani_kordestan_popular_visit.xlsx')
ostani_kordestan_popular_duration=pd.read_excel('busy/ostani_kordestan_popular_duration.xlsx')
del ostani_kordestan_popular_visit['Unnamed: 0']
del ostani_kordestan_popular_duration['Unnamed: 0']

ostani_kermanshah_popular_visit.to_excel('busy/ostani_kermanshah_popular_visit.xlsx')
ostani_kermanshah_popular_duration.to_excel('busy/ostani_kermanshah_popular_duration.xlsx')
ostani_kermanshah_popular_visit=pd.read_excel('busy/ostani_kermanshah_popular_visit.xlsx')
ostani_kermanshah_popular_duration=pd.read_excel('busy/ostani_kermanshah_popular_duration.xlsx')
del ostani_kermanshah_popular_visit['Unnamed: 0']
del ostani_kermanshah_popular_duration['Unnamed: 0']

ostani_kish_popular_visit.to_excel('busy/ostani_kish_popular_visit.xlsx')
ostani_kish_popular_duration.to_excel('busy/ostani_kish_popular_duration.xlsx')
ostani_kish_popular_visit=pd.read_excel('busy/ostani_kish_popular_visit.xlsx')
ostani_kish_popular_duration=pd.read_excel('busy/ostani_kish_popular_duration.xlsx')
del ostani_kish_popular_visit['Unnamed: 0']
del ostani_kish_popular_duration['Unnamed: 0']

ostani_mazandaran_popular_visit.to_excel('busy/ostani_mazandaran_popular_visit.xlsx')
ostani_mazandaran_popular_duration.to_excel('busy/ostani_mazandaran_popular_duration.xlsx')
ostani_mazandaran_popular_visit=pd.read_excel('busy/ostani_mazandaran_popular_visit.xlsx')
ostani_mazandaran_popular_duration=pd.read_excel('busy/ostani_mazandaran_popular_duration.xlsx')
del ostani_mazandaran_popular_visit['Unnamed: 0']
del ostani_mazandaran_popular_duration['Unnamed: 0']

ostani_hamedan_popular_visit.to_excel('busy/ostani_hamedan_popular_visit.xlsx')
ostani_hamedan_popular_duration.to_excel('busy/ostani_hamedan_popular_duration.xlsx')
ostani_hamedan_popular_visit=pd.read_excel('busy/ostani_hamedan_popular_visit.xlsx')
ostani_hamedan_popular_duration=pd.read_excel('busy/ostani_hamedan_popular_duration.xlsx')
del ostani_hamedan_popular_visit['Unnamed: 0']
del ostani_hamedan_popular_duration['Unnamed: 0']

ostani_channels_popular_content=pd.DataFrame()
ostani_channels_popular_content=pd.concat([ostani_abadan_popular_visit, ostani_abadan_popular_duration,
                                          ostani_azarbayjan_gharbi_popular_visit, ostani_azarbayjan_gharbi_popular_duration,
                                          ostani_esfahan_popular_visit, ostani_esfahan_popular_duration,
                                          ostani_aflak_popular_visit, ostani_aflak_popular_duration,
                                          ostani_alborz_popular_visit, ostani_alborz_popular_duration,
                                          ostani_ilam_popular_visit, ostani_ilam_popular_duration,
                                          ostani_baran_popular_visit, ostani_baran_popular_duration,
                                          ostani_boshehr_popular_visit, ostani_boshehr_popular_duration,
                                          ostani_taban_popular_visit, ostani_taban_popular_duration,
                                          ostani_khorasan_razavi_popular_visit, ostani_khorasan_razavi_popular_duration,
                                          ostani_khozestan_popular_visit, ostani_khozestan_popular_duration,
                                          ostani_dena_popular_visit, ostani_dena_popular_duration,
                                          ostani_sabalan_popular_visit, ostani_sabalan_popular_duration,
                                          ostani_sahand_popular_visit, ostani_sahand_popular_duration,
                                          ostani_fars_popular_visit, ostani_fars_popular_duration,
                                          ostani_ghazvin_popular_visit, ostani_ghazvin_popular_duration,
                                          ostani_kordestan_popular_visit, ostani_kordestan_popular_duration,
                                          ostani_kermanshah_popular_visit, ostani_kermanshah_popular_duration,
                                          ostani_kish_popular_visit, ostani_kish_popular_duration,
                                          ostani_mazandaran_popular_visit, ostani_mazandaran_popular_duration,
                                          ostani_hamedan_popular_visit, ostani_hamedan_popular_duration,],axis=1)

writer = pd.ExcelWriter('output/آمار استانی.xlsx', engine='xlsxwriter')
ostani_channels_statistics.to_excel(writer, 'آمار شبکه های استانی')
ostani_channels_popular_content.to_excel(writer, 'محتواهای پربازدید')
writer.save()

print("End ostani")
################################# EKHTESASI ########################################
print("start EKHTESASI")

print("ekhtesasi_perspolis")
ekhtesasi_perspolis=ekhtesasi.query("channel == 'پرسپولیس'")
ekhtesasi_perspolis_visit=ekhtesasi_perspolis['تعداد بازدید'].sum()
ekhtesasi_perspolis_duration=ekhtesasi_perspolis['مدت بازدید'].sum()
ekhtesasi_perspolis_duration=round(ekhtesasi_perspolis_duration*60, 0)
ekhtesasi_perspolis_content=ekhtesasi_perspolis.copy()
ekhtesasi_perspolis_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_perspolis_content=len(ekhtesasi_perspolis_content)
ekhtesasi_perspolis_pivot=ekhtesasi_perspolis.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_perspolis_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_perspolis_popular_visit=ekhtesasi_perspolis_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_perspolis_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_perspolis_popular_duration=ekhtesasi_perspolis_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_perspolis_popular_visit = ekhtesasi_perspolis_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی پرسپولیس', 'نام برنامه': 'محتواهای پربازدید اختصاصی پرسپولیس'})
ekhtesasi_perspolis_popular_duration = ekhtesasi_perspolis_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی پرسپولیس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی پرسپولیس'})

print("ekhtesasi_esteghlal")
ekhtesasi_esteghlal=ekhtesasi.query("channel == 'استقلال'")
ekhtesasi_esteghlal_visit=ekhtesasi_esteghlal['تعداد بازدید'].sum()
ekhtesasi_esteghlal_duration=ekhtesasi_esteghlal['مدت بازدید'].sum()
ekhtesasi_esteghlal_duration=round(ekhtesasi_esteghlal_duration*60, 0)
ekhtesasi_esteghlal_content=ekhtesasi_esteghlal.copy()
ekhtesasi_esteghlal_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_esteghlal_content=len(ekhtesasi_esteghlal_content)
ekhtesasi_esteghlal_pivot=ekhtesasi_esteghlal.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_esteghlal_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_esteghlal_popular_visit=ekhtesasi_esteghlal_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_esteghlal_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_esteghlal_popular_duration=ekhtesasi_esteghlal_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_esteghlal_popular_visit = ekhtesasi_esteghlal_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی استقلال', 'نام برنامه': 'محتواهای پربازدید اختصاصی استقلال'})
ekhtesasi_esteghlal_popular_duration = ekhtesasi_esteghlal_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی استقلال (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی استقلال'})

print("ekhtesasi_shaparak")
ekhtesasi_shaparak=ekhtesasi.query("channel == 'شاپرک'")
ekhtesasi_shaparak_visit=ekhtesasi_shaparak['تعداد بازدید'].sum()
ekhtesasi_shaparak_duration=ekhtesasi_shaparak['مدت بازدید'].sum()
ekhtesasi_shaparak_duration=round(ekhtesasi_shaparak_duration*60, 0)
ekhtesasi_shaparak_content=ekhtesasi_shaparak.copy()
ekhtesasi_shaparak_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_shaparak_content=len(ekhtesasi_shaparak_content)
ekhtesasi_shaparak_pivot=ekhtesasi_shaparak.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_shaparak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_shaparak_popular_visit=ekhtesasi_shaparak_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_shaparak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_shaparak_popular_duration=ekhtesasi_shaparak_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_shaparak_popular_visit = ekhtesasi_shaparak_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی شاپرک', 'نام برنامه': 'محتواهای پربازدید اختصاصی شاپرک'})
ekhtesasi_shaparak_popular_duration = ekhtesasi_shaparak_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی شاپرک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی شاپرک'})

print("ekhtesasi_shetab")
ekhtesasi_shetab=ekhtesasi.query("channel == 'شتاب'")
ekhtesasi_shetab_visit=ekhtesasi_shetab['تعداد بازدید'].sum()
ekhtesasi_shetab_duration=ekhtesasi_shetab['مدت بازدید'].sum()
ekhtesasi_shetab_duration=round(ekhtesasi_shetab_duration*60, 0)
ekhtesasi_shetab_content=ekhtesasi_shetab.copy()
ekhtesasi_shetab_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_shetab_content=len(ekhtesasi_shetab_content)
ekhtesasi_shetab_pivot=ekhtesasi_shetab.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_shetab_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_shetab_popular_visit=ekhtesasi_shetab_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_shetab_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_shetab_popular_duration=ekhtesasi_shetab_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_shetab_popular_visit = ekhtesasi_shetab_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی شتاب', 'نام برنامه': 'محتواهای پربازدید اختصاصی شتاب'})
ekhtesasi_shetab_popular_duration = ekhtesasi_shetab_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی شتاب (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی شتاب'})

print("ekhtesasi_kodak_digiton")
ekhtesasi_kodak_digiton=ekhtesasi.query("channel == 'کودک دیجیتون'")
ekhtesasi_kodak_digiton_visit=ekhtesasi_kodak_digiton['تعداد بازدید'].sum()
ekhtesasi_kodak_digiton_duration=ekhtesasi_kodak_digiton['مدت بازدید'].sum()
ekhtesasi_kodak_digiton_duration=round(ekhtesasi_kodak_digiton_duration*60, 0)
ekhtesasi_kodak_digiton_content=ekhtesasi_kodak_digiton.copy()
ekhtesasi_kodak_digiton_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_kodak_digiton_content=len(ekhtesasi_kodak_digiton_content)
ekhtesasi_kodak_digiton_pivot=ekhtesasi_kodak_digiton.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_kodak_digiton_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_kodak_digiton_popular_visit=ekhtesasi_kodak_digiton_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_kodak_digiton_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_kodak_digiton_popular_duration=ekhtesasi_kodak_digiton_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_kodak_digiton_popular_visit = ekhtesasi_kodak_digiton_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی کودک دیجیتون', 'نام برنامه': 'محتواهای پربازدید اختصاصی کودک دیجیتون'})
ekhtesasi_kodak_digiton_popular_duration = ekhtesasi_kodak_digiton_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی کودک دیجیتون (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی کودک دیجیتون'})

print("ekhtesasi_lenz_sport")
ekhtesasi_lenz_sport=ekhtesasi.query("channel == 'لنزاسپورت'")
ekhtesasi_lenz_sport_visit=ekhtesasi_lenz_sport['تعداد بازدید'].sum()
ekhtesasi_lenz_sport_duration=ekhtesasi_lenz_sport['مدت بازدید'].sum()
ekhtesasi_lenz_sport_duration=round(ekhtesasi_lenz_sport_duration*60, 0)
ekhtesasi_lenz_sport_content=ekhtesasi_lenz_sport.copy()
ekhtesasi_lenz_sport_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_lenz_sport_content=len(ekhtesasi_lenz_sport_content)
ekhtesasi_lenz_sport_pivot=ekhtesasi_lenz_sport.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_lenz_sport_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_lenz_sport_popular_visit=ekhtesasi_lenz_sport_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_lenz_sport_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_lenz_sport_popular_duration=ekhtesasi_lenz_sport_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_lenz_sport_popular_visit = ekhtesasi_lenz_sport_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی لنز اسپرت', 'نام برنامه': 'محتواهای پربازدید اختصاصی لنز اسپرت'})
ekhtesasi_lenz_sport_popular_duration = ekhtesasi_lenz_sport_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی لنز اسپرت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی لنز اسپرت'})

print("ekhtesasi_lenz_sport_plus")
ekhtesasi_lenz_sport_plus=ekhtesasi.query("channel == 'لنز اسپورت پلاس'")
ekhtesasi_lenz_sport_plus_visit=ekhtesasi_lenz_sport_plus['تعداد بازدید'].sum()
ekhtesasi_lenz_sport_plus_duration=ekhtesasi_lenz_sport_plus['مدت بازدید'].sum()
ekhtesasi_lenz_sport_plus_duration=round(ekhtesasi_lenz_sport_plus_duration*60, 0)
ekhtesasi_lenz_sport_plus_content=ekhtesasi_lenz_sport_plus.copy()
ekhtesasi_lenz_sport_plus_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_lenz_sport_plus_content=len(ekhtesasi_lenz_sport_plus_content)
ekhtesasi_lenz_sport_plus_pivot=ekhtesasi_lenz_sport_plus.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_lenz_sport_plus_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_lenz_sport_plus_popular_visit=ekhtesasi_lenz_sport_plus_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_lenz_sport_plus_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_lenz_sport_plus_popular_duration=ekhtesasi_lenz_sport_plus_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_lenz_sport_plus_popular_visit = ekhtesasi_lenz_sport_plus_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی لنز اسپرت پلاس', 'نام برنامه': 'محتواهای پربازدید اختصاصی لنز اسپرت پلاس'})
ekhtesasi_lenz_sport_plus_popular_duration = ekhtesasi_lenz_sport_plus_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی لنز اسپرت پلاس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی لنز اسپرت پلاس'})

print("ekhtesasi_tva_sport")
ekhtesasi_tva_sport=ekhtesasi.query("channel == 'تیوا اسپورت'")
ekhtesasi_tva_sport_visit=ekhtesasi_tva_sport['تعداد بازدید'].sum()
ekhtesasi_tva_sport_duration=ekhtesasi_tva_sport['مدت بازدید'].sum()
ekhtesasi_tva_sport_duration=round(ekhtesasi_tva_sport_duration*60, 0)
ekhtesasi_tva_sport_content=ekhtesasi_tva_sport.copy()
ekhtesasi_tva_sport_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_sport_content=len(ekhtesasi_tva_sport_content)
ekhtesasi_tva_sport_pivot=ekhtesasi_tva_sport.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_sport_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_sport_popular_visit=ekhtesasi_tva_sport_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_sport_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_sport_popular_duration=ekhtesasi_tva_sport_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_sport_popular_visit = ekhtesasi_tva_sport_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا اسپرت', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا اسپرت'})
ekhtesasi_tva_sport_popular_duration = ekhtesasi_tva_sport_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا اسپرت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا اسپرت'})

print("ekhtesasi_tva_sport_two")
ekhtesasi_tva_sport_two=ekhtesasi.query("channel == 'تیوا اسپورت دو'")
ekhtesasi_tva_sport_two_visit=ekhtesasi_tva_sport_two['تعداد بازدید'].sum()
ekhtesasi_tva_sport_two_duration=ekhtesasi_tva_sport_two['مدت بازدید'].sum()
ekhtesasi_tva_sport_two_duration=round(ekhtesasi_tva_sport_two_duration*60, 0)
ekhtesasi_tva_sport_two_content=ekhtesasi_tva_sport_two.copy()
ekhtesasi_tva_sport_two_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_sport_two_content=len(ekhtesasi_tva_sport_two_content)
ekhtesasi_tva_sport_two_pivot=ekhtesasi_tva_sport_two.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_sport_two_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_sport_two_popular_visit=ekhtesasi_tva_sport_two_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_sport_two_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_sport_two_popular_duration=ekhtesasi_tva_sport_two_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_sport_two_popular_visit = ekhtesasi_tva_sport_two_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا اسپرت دو', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا اسپرت دو'})
ekhtesasi_tva_sport_two_popular_duration = ekhtesasi_tva_sport_two_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا اسپرت دو (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا اسپرت دو'})

print("ekhtesasi_tva_boors")
ekhtesasi_tva_boors=ekhtesasi.query("channel == 'تیوا بورس'")
ekhtesasi_tva_boors_visit=ekhtesasi_tva_boors['تعداد بازدید'].sum()
ekhtesasi_tva_boors_duration=ekhtesasi_tva_boors['مدت بازدید'].sum()
ekhtesasi_tva_boors_duration=round(ekhtesasi_tva_boors_duration*60, 0)
ekhtesasi_tva_boors_content=ekhtesasi_tva_boors.copy()
ekhtesasi_tva_boors_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_boors_content=len(ekhtesasi_tva_boors_content)
ekhtesasi_tva_boors_pivot=ekhtesasi_tva_boors.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_boors_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_boors_popular_visit=ekhtesasi_tva_boors_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_boors_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_boors_popular_duration=ekhtesasi_tva_boors_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_boors_popular_visit = ekhtesasi_tva_boors_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا بورس', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا بورس'})
ekhtesasi_tva_boors_popular_duration = ekhtesasi_tva_boors_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا بورس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا بورس'})

print("ekhtesasi_tva_two")
ekhtesasi_tva_two=ekhtesasi.query("channel == 'تیوا دو'")
ekhtesasi_tva_two_visit=ekhtesasi_tva_two['تعداد بازدید'].sum()
ekhtesasi_tva_two_duration=ekhtesasi_tva_two['مدت بازدید'].sum()
ekhtesasi_tva_two_duration=round(ekhtesasi_tva_two_duration*60, 0)
ekhtesasi_tva_two_content=ekhtesasi_tva_two.copy()
ekhtesasi_tva_two_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_two_content=len(ekhtesasi_tva_two_content)
ekhtesasi_tva_two_pivot=ekhtesasi_tva_two.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_two_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_two_popular_visit=ekhtesasi_tva_two_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_two_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_two_popular_duration=ekhtesasi_tva_two_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_two_popular_visit = ekhtesasi_tva_two_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا دو', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا دو'})
ekhtesasi_tva_two_popular_duration = ekhtesasi_tva_two_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا دو (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا دو'})

print("ekhtesasi_tva_film")
ekhtesasi_tva_film=ekhtesasi.query("channel == 'تیوا فیلم'")
ekhtesasi_tva_film_visit=ekhtesasi_tva_film['تعداد بازدید'].sum()
ekhtesasi_tva_film_duration=ekhtesasi_tva_film['مدت بازدید'].sum()
ekhtesasi_tva_film_duration=round(ekhtesasi_tva_film_duration*60, 0)
ekhtesasi_tva_film_content=ekhtesasi_tva_film.copy()
ekhtesasi_tva_film_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_film_content=len(ekhtesasi_tva_film_content)
ekhtesasi_tva_film_pivot=ekhtesasi_tva_film.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_film_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_film_popular_visit=ekhtesasi_tva_film_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_film_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_film_popular_duration=ekhtesasi_tva_film_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_film_popular_visit = ekhtesasi_tva_film_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا فیلم', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا فیلم'})
ekhtesasi_tva_film_popular_duration = ekhtesasi_tva_film_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا فیلم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا فیلم'})

print("ekhtesasi_tva_kodak")
ekhtesasi_tva_kodak=ekhtesasi.query("channel == 'تیوا کودک'")
ekhtesasi_tva_kodak_visit=ekhtesasi_tva_kodak['تعداد بازدید'].sum()
ekhtesasi_tva_kodak_duration=ekhtesasi_tva_kodak['مدت بازدید'].sum()
ekhtesasi_tva_kodak_duration=round(ekhtesasi_tva_kodak_duration*60, 0)
ekhtesasi_tva_kodak_content=ekhtesasi_tva_kodak.copy()
ekhtesasi_tva_kodak_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_kodak_content=len(ekhtesasi_tva_kodak_content)
ekhtesasi_tva_kodak_pivot=ekhtesasi_tva_kodak.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_kodak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_kodak_popular_visit=ekhtesasi_tva_kodak_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_kodak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_kodak_popular_duration=ekhtesasi_tva_kodak_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_kodak_popular_visit = ekhtesasi_tva_kodak_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا کودک', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا کودک'})
ekhtesasi_tva_kodak_popular_duration = ekhtesasi_tva_kodak_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا کودک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا کودک'})

print("ekhtesasi_tva_nava")
ekhtesasi_tva_nava=ekhtesasi.query("channel == 'تیوا نوا'")
ekhtesasi_tva_nava_visit=ekhtesasi_tva_nava['تعداد بازدید'].sum()
ekhtesasi_tva_nava_duration=ekhtesasi_tva_nava['مدت بازدید'].sum()
ekhtesasi_tva_nava_duration=round(ekhtesasi_tva_nava_duration*60, 0)
ekhtesasi_tva_nava_content=ekhtesasi_tva_nava.copy()
ekhtesasi_tva_nava_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_nava_content=len(ekhtesasi_tva_nava_content)
ekhtesasi_tva_nava_pivot=ekhtesasi_tva_nava.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_nava_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_nava_popular_visit=ekhtesasi_tva_nava_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_nava_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_nava_popular_duration=ekhtesasi_tva_nava_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_nava_popular_visit = ekhtesasi_tva_nava_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا نوا', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا نوا'})
ekhtesasi_tva_nava_popular_duration = ekhtesasi_tva_nava_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا نوا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا نوا'})

print("ekhtesasi_tva_one")
ekhtesasi_tva_one=ekhtesasi.query("channel == 'تیوا یک'")
ekhtesasi_tva_one_visit=ekhtesasi_tva_one['تعداد بازدید'].sum()
ekhtesasi_tva_one_duration=ekhtesasi_tva_one['مدت بازدید'].sum()
ekhtesasi_tva_one_duration=round(ekhtesasi_tva_one_duration*60, 0)
ekhtesasi_tva_one_content=ekhtesasi_tva_one.copy()
ekhtesasi_tva_one_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_one_content=len(ekhtesasi_tva_one_content)
ekhtesasi_tva_one_pivot=ekhtesasi_tva_one.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_one_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_one_popular_visit=ekhtesasi_tva_one_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_one_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_one_popular_duration=ekhtesasi_tva_one_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_one_popular_visit = ekhtesasi_tva_one_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا یک', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا یک'})
ekhtesasi_tva_one_popular_duration = ekhtesasi_tva_one_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا یک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا یک'})

print("ekhtesasi_mahfel")
ekhtesasi_mahfel=ekhtesasi.query("channel == 'محفل'")
ekhtesasi_mahfel_visit=ekhtesasi_mahfel['تعداد بازدید'].sum()
ekhtesasi_mahfel_duration=ekhtesasi_mahfel['مدت بازدید'].sum()
ekhtesasi_mahfel_duration=round(ekhtesasi_mahfel_duration*60, 0)
ekhtesasi_mahfel_content=ekhtesasi_mahfel.copy()
ekhtesasi_mahfel_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_mahfel_content=len(ekhtesasi_mahfel_content)
ekhtesasi_mahfel_pivot=ekhtesasi_mahfel.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_mahfel_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_mahfel_popular_visit=ekhtesasi_mahfel_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_mahfel_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_mahfel_popular_duration=ekhtesasi_mahfel_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_mahfel_popular_visit = ekhtesasi_mahfel_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی محفل', 'نام برنامه': 'محتواهای پربازدید اختصاصی محفل'})
ekhtesasi_mahfel_popular_duration = ekhtesasi_mahfel_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی محفل (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی محفل'})

print("ekhtesasi_tva_avand")
ekhtesasi_tva_avand=ekhtesasi.query("channel == 'تیوا آوند'")
ekhtesasi_tva_avand_visit=ekhtesasi_tva_avand['تعداد بازدید'].sum()
ekhtesasi_tva_avand_duration=ekhtesasi_tva_avand['مدت بازدید'].sum()
ekhtesasi_tva_avand_duration=round(ekhtesasi_tva_avand_duration*60, 0)
ekhtesasi_tva_avand_content=ekhtesasi_tva_avand.copy()
ekhtesasi_tva_avand_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
ekhtesasi_tva_avand_content=len(ekhtesasi_tva_avand_content)
ekhtesasi_tva_avand_pivot=ekhtesasi_tva_avand.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
ekhtesasi_tva_avand_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_avand_popular_visit=ekhtesasi_tva_avand_pivot.iloc[0:10 , [0, 5]]
ekhtesasi_tva_avand_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_tva_avand_popular_duration=ekhtesasi_tva_avand_pivot.iloc[0:10 , [0, 4]]

ekhtesasi_tva_avand_popular_visit = ekhtesasi_tva_avand_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا آوند', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا آوند'})
ekhtesasi_tva_avand_popular_duration = ekhtesasi_tva_avand_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا آوند (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا آوند'})

print("ekhtesasi_baghimande")
total_ekhtesasi_malom=pd.DataFrame()
total_ekhtesasi_majhol=pd.DataFrame()
       
total_ekhtesasi_malom=pd.concat([ekhtesasi_mahfel,
                                 ekhtesasi_tva_one,
                                 ekhtesasi_tva_nava,
                                 ekhtesasi_tva_kodak,
                                 ekhtesasi_tva_film,
                                 ekhtesasi_tva_two,
                                 ekhtesasi_tva_boors,
                                 ekhtesasi_tva_sport_two,
                                 ekhtesasi_tva_sport,
                                 ekhtesasi_lenz_sport_plus,
                                 ekhtesasi_lenz_sport,
                                 ekhtesasi_kodak_digiton,
                                 ekhtesasi_shetab,
                                 ekhtesasi_shaparak,
                                 ekhtesasi_esteghlal,
                                 ekhtesasi_perspolis,
                                 ekhtesasi_tva_avand,], axis=0)

total_ekhtesasi_majhol=pd.concat([ekhtesasi, total_ekhtesasi_malom]).drop_duplicates(keep=False)

print("dataframe ekhtesasi channels")
ekhtesasi_channels_statistics={'channel_name': ['اختصاصی محفل', 'اختصاصی تیوا یک', 'اختصاصی تیوا نوا', 'اختصاصی تیوا کودک',
                                     'اختصاصی تیوا فیلم', 'اختصاصی تیوا دو', 'اختصاصی تیوا بورس', 'اختصاصی تیوا اسپرت دو',
                                     'اختصاصی تیوا اسپرت', 'اختصاصی لنز اسپرت پلاس', 'اختصاصی لنز اسپرت','اختصاصی کودک دیجیتون',
                                     'اختصاصی شتاب', 'اختصاصی شاپرک', 'اختصاصی استقلال', 'اختصاصی پرسپولیس',
                                     'اختصاصی تیوا آوند'],
       'channel_content': [ekhtesasi_mahfel_content, ekhtesasi_tva_one_content, ekhtesasi_tva_nava_content, ekhtesasi_tva_kodak_content, 
                           ekhtesasi_tva_film_content, ekhtesasi_tva_two_content, ekhtesasi_tva_boors_content, ekhtesasi_tva_sport_two_content, 
                           ekhtesasi_tva_sport_content, ekhtesasi_lenz_sport_plus_content, ekhtesasi_lenz_sport_content, ekhtesasi_kodak_digiton_content,
                           ekhtesasi_shetab_content, ekhtesasi_shaparak_content, ekhtesasi_esteghlal_content, ekhtesasi_perspolis_content,
                           ekhtesasi_tva_avand_content],
       'channel_visit': [ekhtesasi_mahfel_visit, ekhtesasi_tva_one_visit, ekhtesasi_tva_nava_visit, ekhtesasi_tva_kodak_visit,
                         ekhtesasi_tva_film_visit, ekhtesasi_tva_two_visit,ekhtesasi_tva_boors_visit, ekhtesasi_tva_sport_two_visit, 
                         ekhtesasi_tva_sport_visit, ekhtesasi_lenz_sport_plus_visit, ekhtesasi_lenz_sport_visit, ekhtesasi_kodak_digiton_visit,
                           ekhtesasi_shetab_visit, ekhtesasi_shaparak_visit, ekhtesasi_esteghlal_visit,ekhtesasi_perspolis_visit,
                           ekhtesasi_tva_avand_visit],
       'channel_duration': [ekhtesasi_mahfel_duration, ekhtesasi_tva_one_duration, ekhtesasi_tva_nava_duration, ekhtesasi_tva_kodak_duration,
                            ekhtesasi_tva_film_duration, ekhtesasi_tva_two_duration,ekhtesasi_tva_boors_duration, ekhtesasi_tva_sport_two_duration, 
                            ekhtesasi_tva_sport_duration, ekhtesasi_lenz_sport_plus_duration, ekhtesasi_lenz_sport_duration, ekhtesasi_kodak_digiton_duration,
                            ekhtesasi_shetab_duration, ekhtesasi_shaparak_duration, ekhtesasi_esteghlal_duration,ekhtesasi_perspolis_duration,
                            ekhtesasi_tva_avand_duration],}
ekhtesasi_channels_statistics=pd.DataFrame(ekhtesasi_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
ekhtesasi_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
ekhtesasi_channels_statistics=ekhtesasi_channels_statistics.rename(columns={'channel_content': 'نام شبکه', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})


ekhtesasi_mahfel_popular_visit.to_excel('busy/ekhtesasi_mahfel_popular_visit.xlsx')
ekhtesasi_mahfel_popular_duration.to_excel('busy/ekhtesasi_mahfel_popular_duration.xlsx')
ekhtesasi_mahfel_popular_visit=pd.read_excel('busy/ekhtesasi_mahfel_popular_visit.xlsx')
ekhtesasi_mahfel_popular_duration=pd.read_excel('busy/ekhtesasi_mahfel_popular_duration.xlsx')
del ekhtesasi_mahfel_popular_visit['Unnamed: 0']
del ekhtesasi_mahfel_popular_duration['Unnamed: 0']

ekhtesasi_tva_one_popular_visit.to_excel('busy/ekhtesasi_tva_one_popular_visit.xlsx')
ekhtesasi_tva_one_popular_duration.to_excel('busy/ekhtesasi_tva_one_popular_duration.xlsx')
ekhtesasi_tva_one_popular_visit=pd.read_excel('busy/ekhtesasi_tva_one_popular_visit.xlsx')
ekhtesasi_tva_one_popular_duration=pd.read_excel('busy/ekhtesasi_tva_one_popular_duration.xlsx')
del ekhtesasi_tva_one_popular_visit['Unnamed: 0']
del ekhtesasi_tva_one_popular_duration['Unnamed: 0']

ekhtesasi_tva_nava_popular_visit.to_excel('busy/ekhtesasi_tva_nava_popular_visit.xlsx')
ekhtesasi_tva_nava_popular_duration.to_excel('busy/ekhtesasi_tva_nava_popular_duration.xlsx')
ekhtesasi_tva_nava_popular_visit=pd.read_excel('busy/ekhtesasi_tva_nava_popular_visit.xlsx')
ekhtesasi_tva_nava_popular_duration=pd.read_excel('busy/ekhtesasi_tva_nava_popular_duration.xlsx')
del ekhtesasi_tva_nava_popular_visit['Unnamed: 0']
del ekhtesasi_tva_nava_popular_duration['Unnamed: 0']

ekhtesasi_tva_kodak_popular_visit.to_excel('busy/ekhtesasi_tva_kodak_popular_visit.xlsx')
ekhtesasi_tva_kodak_popular_duration.to_excel('busy/ekhtesasi_tva_kodak_popular_duration.xlsx')
ekhtesasi_tva_kodak_popular_visit=pd.read_excel('busy/ekhtesasi_tva_kodak_popular_visit.xlsx')
ekhtesasi_tva_kodak_popular_duration=pd.read_excel('busy/ekhtesasi_tva_kodak_popular_duration.xlsx')
del ekhtesasi_tva_kodak_popular_visit['Unnamed: 0']
del ekhtesasi_tva_kodak_popular_duration['Unnamed: 0']

ekhtesasi_tva_film_popular_visit.to_excel('busy/ekhtesasi_tva_film_popular_visit.xlsx')
ekhtesasi_tva_film_popular_duration.to_excel('busy/ekhtesasi_tva_film_popular_duration.xlsx')
ekhtesasi_tva_film_popular_visit=pd.read_excel('busy/ekhtesasi_tva_film_popular_visit.xlsx')
ekhtesasi_tva_film_popular_duration=pd.read_excel('busy/ekhtesasi_tva_film_popular_duration.xlsx')
del ekhtesasi_tva_film_popular_visit['Unnamed: 0']
del ekhtesasi_tva_film_popular_duration['Unnamed: 0']

ekhtesasi_tva_two_popular_visit.to_excel('busy/ekhtesasi_tva_two_popular_visit.xlsx')
ekhtesasi_tva_two_popular_duration.to_excel('busy/ekhtesasi_tva_two_popular_duration.xlsx')
ekhtesasi_tva_two_popular_visit=pd.read_excel('busy/ekhtesasi_tva_two_popular_visit.xlsx')
ekhtesasi_tva_two_popular_duration=pd.read_excel('busy/ekhtesasi_tva_two_popular_duration.xlsx')
del ekhtesasi_tva_two_popular_visit['Unnamed: 0']
del ekhtesasi_tva_two_popular_duration['Unnamed: 0']

ekhtesasi_tva_boors_popular_visit.to_excel('busy/ekhtesasi_tva_boors_popular_visit.xlsx')
ekhtesasi_tva_boors_popular_duration.to_excel('busy/ekhtesasi_tva_boors_popular_duration.xlsx')
ekhtesasi_tva_boors_popular_visit=pd.read_excel('busy/ekhtesasi_tva_boors_popular_visit.xlsx')
ekhtesasi_tva_boors_popular_duration=pd.read_excel('busy/ekhtesasi_tva_boors_popular_duration.xlsx')
del ekhtesasi_tva_boors_popular_visit['Unnamed: 0']
del ekhtesasi_tva_boors_popular_duration['Unnamed: 0']

ekhtesasi_tva_sport_two_popular_visit.to_excel('busy/ekhtesasi_tva_sport_two_popular_visit.xlsx')
ekhtesasi_tva_sport_two_popular_duration.to_excel('busy/ekhtesasi_tva_sport_two_popular_duration.xlsx')
ekhtesasi_tva_sport_two_popular_visit=pd.read_excel('busy/ekhtesasi_tva_sport_two_popular_visit.xlsx')
ekhtesasi_tva_sport_two_popular_duration=pd.read_excel('busy/ekhtesasi_tva_sport_two_popular_duration.xlsx')
del ekhtesasi_tva_sport_two_popular_visit['Unnamed: 0']
del ekhtesasi_tva_sport_two_popular_duration['Unnamed: 0']

ekhtesasi_tva_sport_popular_visit.to_excel('busy/ekhtesasi_tva_sport_popular_visit.xlsx')
ekhtesasi_tva_sport_popular_duration.to_excel('busy/ekhtesasi_tva_sport_popular_duration.xlsx')
ekhtesasi_tva_sport_popular_visit=pd.read_excel('busy/ekhtesasi_tva_sport_popular_visit.xlsx')
ekhtesasi_tva_sport_popular_duration=pd.read_excel('busy/ekhtesasi_tva_sport_popular_duration.xlsx')
del ekhtesasi_tva_sport_popular_visit['Unnamed: 0']
del ekhtesasi_tva_sport_popular_duration['Unnamed: 0']

ekhtesasi_lenz_sport_plus_popular_visit.to_excel('busy/ekhtesasi_lenz_sport_plus_popular_visit.xlsx')
ekhtesasi_lenz_sport_plus_popular_duration.to_excel('busy/ekhtesasi_lenz_sport_plus_popular_duration.xlsx')
ekhtesasi_lenz_sport_plus_popular_visit=pd.read_excel('busy/ekhtesasi_lenz_sport_plus_popular_visit.xlsx')
ekhtesasi_lenz_sport_plus_popular_duration=pd.read_excel('busy/ekhtesasi_lenz_sport_plus_popular_duration.xlsx')
del ekhtesasi_lenz_sport_plus_popular_visit['Unnamed: 0']
del ekhtesasi_lenz_sport_plus_popular_duration['Unnamed: 0']

ekhtesasi_lenz_sport_popular_visit.to_excel('busy/ekhtesasi_lenz_sport_popular_visit.xlsx')
ekhtesasi_lenz_sport_popular_duration.to_excel('busy/ekhtesasi_lenz_sport_popular_duration.xlsx')
ekhtesasi_lenz_sport_popular_visit=pd.read_excel('busy/ekhtesasi_lenz_sport_popular_visit.xlsx')
ekhtesasi_lenz_sport_popular_duration=pd.read_excel('busy/ekhtesasi_lenz_sport_popular_duration.xlsx')
del ekhtesasi_lenz_sport_popular_visit['Unnamed: 0']
del ekhtesasi_lenz_sport_popular_duration['Unnamed: 0']

ekhtesasi_kodak_digiton_popular_visit.to_excel('busy/ekhtesasi_kodak_digiton_popular_visit.xlsx')
ekhtesasi_kodak_digiton_popular_duration.to_excel('busy/ekhtesasi_kodak_digiton_popular_duration.xlsx')
ekhtesasi_kodak_digiton_popular_visit=pd.read_excel('busy/ekhtesasi_kodak_digiton_popular_visit.xlsx')
ekhtesasi_kodak_digiton_popular_duration=pd.read_excel('busy/ekhtesasi_kodak_digiton_popular_duration.xlsx')
del ekhtesasi_kodak_digiton_popular_visit['Unnamed: 0']
del ekhtesasi_kodak_digiton_popular_duration['Unnamed: 0']

ekhtesasi_shetab_popular_visit.to_excel('busy/ekhtesasi_shetab_popular_visit.xlsx')
ekhtesasi_shetab_popular_duration.to_excel('busy/ekhtesasi_shetab_popular_duration.xlsx')
ekhtesasi_shetab_popular_visit=pd.read_excel('busy/ekhtesasi_shetab_popular_visit.xlsx')
ekhtesasi_shetab_popular_duration=pd.read_excel('busy/ekhtesasi_shetab_popular_duration.xlsx')
del ekhtesasi_shetab_popular_visit['Unnamed: 0']
del ekhtesasi_shetab_popular_duration['Unnamed: 0']

ekhtesasi_shaparak_popular_visit.to_excel('busy/ekhtesasi_shaparak_popular_visit.xlsx')
ekhtesasi_shaparak_popular_duration.to_excel('busy/ekhtesasi_shaparak_popular_duration.xlsx')
ekhtesasi_shaparak_popular_visit=pd.read_excel('busy/ekhtesasi_shaparak_popular_visit.xlsx')
ekhtesasi_shaparak_popular_duration=pd.read_excel('busy/ekhtesasi_shaparak_popular_duration.xlsx')
del ekhtesasi_shaparak_popular_visit['Unnamed: 0']
del ekhtesasi_shaparak_popular_duration['Unnamed: 0']

ekhtesasi_esteghlal_popular_visit.to_excel('busy/ekhtesasi_esteghlal_popular_visit.xlsx')
ekhtesasi_esteghlal_popular_duration.to_excel('busy/ekhtesasi_esteghlal_popular_duration.xlsx')
ekhtesasi_esteghlal_popular_visit=pd.read_excel('busy/ekhtesasi_esteghlal_popular_visit.xlsx')
ekhtesasi_esteghlal_popular_duration=pd.read_excel('busy/ekhtesasi_esteghlal_popular_duration.xlsx')
del ekhtesasi_esteghlal_popular_visit['Unnamed: 0']
del ekhtesasi_esteghlal_popular_duration['Unnamed: 0']

ekhtesasi_perspolis_popular_visit.to_excel('busy/ekhtesasi_perspolis_popular_visit.xlsx')
ekhtesasi_perspolis_popular_duration.to_excel('busy/ekhtesasi_perspolis_popular_duration.xlsx')
ekhtesasi_perspolis_popular_visit=pd.read_excel('busy/ekhtesasi_perspolis_popular_visit.xlsx')
ekhtesasi_perspolis_popular_duration=pd.read_excel('busy/ekhtesasi_perspolis_popular_duration.xlsx')
del ekhtesasi_perspolis_popular_visit['Unnamed: 0']
del ekhtesasi_perspolis_popular_duration['Unnamed: 0']

ekhtesasi_tva_avand_popular_visit.to_excel('busy/ekhtesasi_tva_avand_popular_visit.xlsx')
ekhtesasi_tva_avand_popular_duration.to_excel('busy/ekhtesasi_tva_avand_popular_duration.xlsx')
ekhtesasi_tva_avand_popular_visit=pd.read_excel('busy/ekhtesasi_tva_avand_popular_visit.xlsx')
ekhtesasi_tva_avand_popular_duration=pd.read_excel('busy/ekhtesasi_tva_avand_popular_duration.xlsx')
del ekhtesasi_tva_avand_popular_visit['Unnamed: 0']
del ekhtesasi_tva_avand_popular_duration['Unnamed: 0']

ekhtesasi_channels_popular_content=pd.DataFrame()
ekhtesasi_channels_popular_content=pd.concat([ekhtesasi_tva_kodak_popular_visit, ekhtesasi_tva_kodak_popular_duration,
                                           ekhtesasi_tva_nava_popular_visit, ekhtesasi_tva_kodak_popular_duration,
                                           ekhtesasi_tva_one_popular_visit, ekhtesasi_tva_one_popular_duration,
                                           ekhtesasi_mahfel_popular_visit, ekhtesasi_mahfel_popular_duration,
                                           ekhtesasi_tva_film_popular_visit, ekhtesasi_tva_film_popular_duration,
                                           ekhtesasi_tva_two_popular_visit, ekhtesasi_tva_two_popular_duration,
                                           ekhtesasi_tva_boors_popular_visit, ekhtesasi_tva_boors_popular_duration,
                                           ekhtesasi_tva_sport_two_popular_visit, ekhtesasi_tva_sport_two_popular_duration,
                                           ekhtesasi_tva_sport_popular_visit, ekhtesasi_tva_sport_popular_duration,
                                           ekhtesasi_lenz_sport_plus_popular_visit, ekhtesasi_lenz_sport_plus_popular_duration,
                                           ekhtesasi_lenz_sport_popular_visit, ekhtesasi_lenz_sport_popular_duration,
                                           ekhtesasi_kodak_digiton_popular_visit, ekhtesasi_kodak_digiton_popular_duration,
                                           ekhtesasi_shetab_popular_visit, ekhtesasi_shetab_popular_duration,
                                           ekhtesasi_shaparak_popular_visit, ekhtesasi_shaparak_popular_duration,
                                           ekhtesasi_esteghlal_popular_visit, ekhtesasi_esteghlal_popular_duration,
                                           ekhtesasi_perspolis_popular_visit, ekhtesasi_perspolis_popular_duration,
                                           ekhtesasi_tva_avand_popular_visit, ekhtesasi_tva_avand_popular_duration,],axis=1)

writer = pd.ExcelWriter('output/آمار اختصاصی.xlsx', engine='xlsxwriter')
ekhtesasi_channels_statistics.to_excel(writer, 'آمار شبکه های اختصاصی')
ekhtesasi_channels_popular_content.to_excel(writer, 'محتواهای پربازدید')
writer.save()

print("END EKHTESASI")
################################# VOD ########################################
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
tva_vod_serial_popular_visit_middle=tva_vod_serial_popular_visit_middle.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ متوسط بازدید هر محتوا- تیوا', 'visit middle': 'تعداد بازدید'})

tva_vod_serial['duration middle']=round(tva_vod_serial['Avg. Duration (sec)']/tva_vod_serial['episode'], 0)
tva_vod_serial.sort_values('duration middle', axis = 0, ascending = False, inplace = True, na_position ='last')
tva_vod_serial_popular_duration_middle=tva_vod_serial.iloc[0:10 , [0, 8]]
tva_vod_serial_popular_duration_middle=tva_vod_serial_popular_duration_middle.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ متوسط زمان بازدید هر محتوا- تیوا', 'duration middle': 'مدت زمان بازدید (به دقیقه)'})

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
vod_lenz['content name'].replace('', 'nan', inplace=True)
vod_lenz.dropna(subset=['content name'], inplace=True)

vod_lenz.insert(6, 'Name content summary', '')
vod_lenz_content_name=vod_lenz['content name']
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
lenz_vod_serial_popular_visit_middle=lenz_vod_serial_popular_visit_middle.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ متوسط بازدید هر محتوا- لنز', 'visit middle': 'تعداد بازدید'})

lenz_vod_serial['duration middle']=round(lenz_vod_serial['access duration (hour)']/lenz_vod_serial['episode'], 0)
lenz_vod_serial.sort_values('duration middle', axis = 0, ascending = False, inplace = True, na_position ='last')
lenz_vod_serial_popular_duration_middle=lenz_vod_serial.iloc[0:10 , [0, 7]]
lenz_vod_serial_popular_duration_middle=lenz_vod_serial_popular_duration_middle.rename(columns={'Name content summary': 'نام محتواهای پربازدید سریالی به لحاظ متوسط زمان بازدید هر محتوا- لنز', 'duration middle': 'مدت زمان بازدید (به دقیقه)'})

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

#lenz_vod_film_popular_visit=lenz_vod_film_popular_visit.rename(columns={'access times': 'Sessions'})
#vod_film_popular_total = lenz_vod_film_popular_visit.append(tva_vod_film_popular_visit, ignore_index=True)
#vod_film_popular_total.sort_values('Sessions', axis = 0, ascending = False, inplace = True, na_position ='last')
#vod_film_popular_total=vod_film_popular_total.iloc[0:10 , [0, 1]]

#lenz_vod_serial_popular_visit=lenz_vod_serial_popular_visit.rename(columns={'access times': 'Sessions'})
#vod_serial_popular_total = lenz_vod_serial_popular_visit.append(tva_vod_serial_popular_visit, ignore_index=True)
#vod_serial_popular_total.sort_values('Sessions', axis = 0, ascending = False, inplace = True, na_position ='last')
#vod_serial_popular_total=vod_serial_popular_total.iloc[0:10 , [0, 1]]

#vod_serial_popular_total_middle = lenz_vod_serial_popular_visit_middle.append(tva_vod_serial_popular_visit_middle, ignore_index=True)
#vod_serial_popular_total_middle.sort_values('visit middle', axis = 0, ascending = False, inplace = True, na_position ='last')
#vod_serial_popular_total_middle=vod_serial_popular_total_middle.iloc[0:10 , [0, 1]]

lenz_vod_content=lenz_vod_film_content+lenz_vod_serial_content
lenz_vod_visit=lenz_vod_serial_visit+lenz_vod_film_visit
lenz_vod_duration=lenz_vod_serial_duration+lenz_vod_film_duration
print("End lenz vod")

print("vod statistics summary")

vod_statistics_summary={'operators': ['تیوا', 'لنز'],
       'all_content': [tva_vod_content, lenz_vod_content],
       'total_visit': [tva_vod_visit, lenz_vod_visit],
       'total_duration': [tva_vod_duration, lenz_vod_duration],
       'vod_serial_content': [tva_vod_serial_content, lenz_vod_serial_content],
       'vod_film_content': [tva_vod_film_content, lenz_vod_film_content],
       'vod_serial_visit': [tva_vod_serial_visit, lenz_vod_serial_visit],
       'vod_film_visit': [tva_vod_film_visit, lenz_vod_film_visit],
       'vod_serial_duration': [tva_vod_serial_duration, lenz_vod_serial_duration],
       'vod_film_duration': [tva_vod_film_duration, lenz_vod_film_duration],}
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

print("dataframe of vod popular content")

vod_popular_content=pd.DataFrame()
vod_popular_content=pd.concat([tva_vod_serial_popular_visit, tva_vod_serial_popular_duration,
                                              tva_vod_serial_popular_visit_middle, tva_vod_serial_popular_duration_middle,
                                              tva_vod_film_popular_visit, tva_vod_film_popular_duration,
                                              lenz_vod_serial_popular_visit, lenz_vod_serial_popular_duration,
                                              lenz_vod_serial_popular_visit_middle, lenz_vod_serial_popular_duration_middle,
                                              lenz_vod_film_popular_visit, lenz_vod_film_popular_duration],axis=1)

writer = pd.ExcelWriter('output/آمار VOD.xlsx', engine='xlsxwriter')
vod_statistics_summary.to_excel(writer, 'خلاصه آمار VOD')
vod_popular_content.to_excel(writer, 'محتواهای پربازدید')
writer.save()

################################# old Data ########################################

        ########################### فروردین #############################
print("EPG Farvardin 1399")
EPG_Farvardin_1399=pd.read_excel('EPG/EPG 1399/EPG Farvardin 1399.xlsx', sheet_name='آمار')
EPG_Farvardin_1399.fillna(0, inplace=True)
sima_1_visit_Farvardin_1399=EPG_Farvardin_1399.iat[1, 4]
sima_2_visit_Farvardin_1399=EPG_Farvardin_1399.iat[2, 4]
sima_3_visit_Farvardin_1399=EPG_Farvardin_1399.iat[3, 4]
sima_4_visit_Farvardin_1399=EPG_Farvardin_1399.iat[4, 4]
sima_5_visit_Farvardin_1399=EPG_Farvardin_1399.iat[5, 4]
sima_khabar_visit_Farvardin_1399=EPG_Farvardin_1399.iat[6, 4]
sima_ofogh_visit_Farvardin_1399=EPG_Farvardin_1399.iat[7, 4]
sima_pooya_visit_Farvardin_1399=EPG_Farvardin_1399.iat[8, 4]
sima_omid_visit_Farvardin_1399=EPG_Farvardin_1399.iat[9, 4]
sima_ifilm_visit_Farvardin_1399=EPG_Farvardin_1399.iat[10, 4]
sima_namayesh_visit_Farvardin_1399=EPG_Farvardin_1399.iat[11, 4]
sima_tamasha_visit_Farvardin_1399=EPG_Farvardin_1399.iat[12, 4]
sima_mostanad_visit_Farvardin_1399=EPG_Farvardin_1399.iat[13, 4]
sima_shoma_visit_Farvardin_1399=EPG_Farvardin_1399.iat[14, 4]
sima_amozesh_visit_Farvardin_1399=EPG_Farvardin_1399.iat[15, 4]
sima_varzesh_visit_Farvardin_1399=EPG_Farvardin_1399.iat[16, 4]
sima_nasim_visit_Farvardin_1399=EPG_Farvardin_1399.iat[17, 4]
sima_qoran_visit_Farvardin_1399=EPG_Farvardin_1399.iat[18, 4]
sima_salamat_visit_Farvardin_1399=EPG_Farvardin_1399.iat[19, 4]
sima_irankala_visit_Farvardin_1399=EPG_Farvardin_1399.iat[20, 4]
sima_alalam_visit_Farvardin_1399=EPG_Farvardin_1399.iat[21, 4]
sima_alkosar_visit_Farvardin_1399=EPG_Farvardin_1399.iat[22, 4]
sima_presstv_visit_Farvardin_1399=EPG_Farvardin_1399.iat[23, 4]
sima_sepehr_visit_Farvardin_1399=EPG_Farvardin_1399.iat[24, 4]

sima_1_duration_Farvardin_1399=EPG_Farvardin_1399.iat[1, 6]
sima_2_duration_Farvardin_1399=EPG_Farvardin_1399.iat[2, 6]
sima_3_duration_Farvardin_1399=EPG_Farvardin_1399.iat[3, 6]
sima_4_duration_Farvardin_1399=EPG_Farvardin_1399.iat[4, 6]
sima_5_duration_Farvardin_1399=EPG_Farvardin_1399.iat[5, 6]
sima_khabar_duration_Farvardin_1399=EPG_Farvardin_1399.iat[6, 6]
sima_ofogh_duration_Farvardin_1399=EPG_Farvardin_1399.iat[7, 6]
sima_pooya_duration_Farvardin_1399=EPG_Farvardin_1399.iat[8, 6]
sima_omid_duration_Farvardin_1399=EPG_Farvardin_1399.iat[9, 6]
sima_ifilm_duration_Farvardin_1399=EPG_Farvardin_1399.iat[10, 6]
sima_namayesh_duration_Farvardin_1399=EPG_Farvardin_1399.iat[11, 6]
sima_tamasha_duration_Farvardin_1399=EPG_Farvardin_1399.iat[12, 6]
sima_mostanad_duration_Farvardin_1399=EPG_Farvardin_1399.iat[13, 6]
sima_shoma_duration_Farvardin_1399=EPG_Farvardin_1399.iat[14, 6]
sima_amozesh_duration_Farvardin_1399=EPG_Farvardin_1399.iat[15, 6]
sima_varzesh_duration_Farvardin_1399=EPG_Farvardin_1399.iat[16, 6]
sima_nasim_duration_Farvardin_1399=EPG_Farvardin_1399.iat[17, 6]
sima_qoran_duration_Farvardin_1399=EPG_Farvardin_1399.iat[18, 6]
sima_salamat_duration_Farvardin_1399=EPG_Farvardin_1399.iat[19, 6]
sima_irankala_duration_Farvardin_1399=EPG_Farvardin_1399.iat[20, 6]
sima_alalam_duration_Farvardin_1399=EPG_Farvardin_1399.iat[21, 6]
sima_alkosar_duration_Farvardin_1399=EPG_Farvardin_1399.iat[22, 6]
sima_presstv_duration_Farvardin_1399=EPG_Farvardin_1399.iat[23, 6]
sima_sepehr_duration_Farvardin_1399=EPG_Farvardin_1399.iat[24, 6]

sima_lenz_visit_Farvardin_1399=EPG_Farvardin_1399.iat[33, 2]
sima_aio_visit_Farvardin_1399=EPG_Farvardin_1399.iat[34, 2]
sima_anten_visit_Farvardin_1399=EPG_Farvardin_1399.iat[35, 2]
sima_tva_visit_Farvardin_1399=EPG_Farvardin_1399.iat[36, 2]
sima_fam_visit_Farvardin_1399=EPG_Farvardin_1399.iat[37, 2]
sima_televebion_visit_Farvardin_1399=EPG_Farvardin_1399.iat[38, 2]
sima_sepehr_Farvardin_1399=EPG_Farvardin_1399.iat[39, 2]
sima_shima_visit_Farvardin_1399=EPG_Farvardin_1399.iat[40, 2]
sima_site_visit_Farvardin_1399=EPG_Farvardin_1399.iat[41, 2]

register_user_lenz_Farvardin_1399=EPG_Farvardin_1399.iat[36, 4]
register_user_aio_Farvardin_1399=EPG_Farvardin_1399.iat[37, 4]
register_user_anten_Farvardin_1399=EPG_Farvardin_1399.iat[38, 4]
register_user_tva_Farvardin_1399=EPG_Farvardin_1399.iat[39, 4]
register_user_fam_Farvardin_1399=EPG_Farvardin_1399.iat[40, 4]
register_user_televebion_Farvardin_1399=EPG_Farvardin_1399.iat[41, 4]
register_user_sepehr_Farvardin_1399=EPG_Farvardin_1399.iat[42, 4]
register_user_shima_Farvardin_1399=EPG_Farvardin_1399.iat[43, 4]
register_user_site_Farvardin_1399=EPG_Farvardin_1399.iat[44, 4]

active_user_lenz_Farvardin_1399=EPG_Farvardin_1399.iat[36, 10]
active_user_aio_Farvardin_1399=EPG_Farvardin_1399.iat[37, 10]
active_user_anten_Farvardin_1399=EPG_Farvardin_1399.iat[38, 10]
active_user_tva_Farvardin_1399=EPG_Farvardin_1399.iat[39, 10]
active_user_fam_Farvardin_1399=EPG_Farvardin_1399.iat[40, 10]
active_user_televebion_Farvardin_1399=EPG_Farvardin_1399.iat[41, 10]
active_user_sepehr_Farvardin_1399=EPG_Farvardin_1399.iat[42, 10]
active_user_shima_Farvardin_1399=EPG_Farvardin_1399.iat[43, 10]
active_user_site_Farvardin_1399=EPG_Farvardin_1399.iat[44, 10]

all_visit_Farvardin_1399=EPG_Farvardin_1399.iat[25, 4]
all_duration_Farvardin_1399=EPG_Farvardin_1399.iat[25, 6]
all_content_sima_Farvardin_1399=EPG_Farvardin_1399.iat[25, 2]
all_register_user_Farvardin_1399=sum(EPG_Farvardin_1399.iloc[36:44, 4])
all_active_user_Farvardin_1399=sum(EPG_Farvardin_1399.iloc[36:44, 10])

Farvardin_1399_sima_visit_channels=pd.DataFrame()
Farvardin_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Farvardin_1399, sima_2_visit_Farvardin_1399, sima_3_visit_Farvardin_1399,
                 sima_4_visit_Farvardin_1399, sima_5_visit_Farvardin_1399, sima_khabar_visit_Farvardin_1399,
                 sima_ofogh_visit_Farvardin_1399, sima_pooya_visit_Farvardin_1399, sima_omid_visit_Farvardin_1399,
                 sima_ifilm_visit_Farvardin_1399, sima_namayesh_visit_Farvardin_1399, sima_tamasha_visit_Farvardin_1399,
                 sima_mostanad_visit_Farvardin_1399, sima_shoma_visit_Farvardin_1399, sima_amozesh_visit_Farvardin_1399,
                 sima_varzesh_visit_Farvardin_1399, sima_nasim_visit_Farvardin_1399, sima_qoran_visit_Farvardin_1399,
                 sima_salamat_visit_Farvardin_1399, sima_irankala_visit_Farvardin_1399, sima_alalam_visit_Farvardin_1399,
                 sima_alkosar_visit_Farvardin_1399, sima_presstv_visit_Farvardin_1399, sima_sepehr_visit_Farvardin_1399,],
        'duration': [sima_1_duration_Farvardin_1399, sima_2_duration_Farvardin_1399, sima_3_duration_Farvardin_1399,
                 sima_4_duration_Farvardin_1399, sima_5_duration_Farvardin_1399, sima_khabar_duration_Farvardin_1399,
                 sima_ofogh_duration_Farvardin_1399, sima_pooya_duration_Farvardin_1399, sima_omid_duration_Farvardin_1399,
                 sima_ifilm_duration_Farvardin_1399, sima_namayesh_duration_Farvardin_1399, sima_tamasha_duration_Farvardin_1399,
                 sima_mostanad_duration_Farvardin_1399, sima_shoma_duration_Farvardin_1399, sima_amozesh_duration_Farvardin_1399,
                 sima_varzesh_duration_Farvardin_1399, sima_nasim_duration_Farvardin_1399, sima_qoran_duration_Farvardin_1399,
                 sima_salamat_duration_Farvardin_1399, sima_irankala_duration_Farvardin_1399, sima_alalam_duration_Farvardin_1399,
                 sima_alkosar_duration_Farvardin_1399, sima_presstv_duration_Farvardin_1399, sima_sepehr_duration_Farvardin_1399,],}
Farvardin_1399_sima_visit_channels=pd.DataFrame(Farvardin_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Farvardin_1399_sima_visit_channels=Farvardin_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Farvardin_1399_operator_data=pd.DataFrame()
Farvardin_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Farvardin_1399, sima_aio_visit_Farvardin_1399, sima_anten_visit_Farvardin_1399,
                 sima_tva_visit_Farvardin_1399, sima_fam_visit_Farvardin_1399, sima_televebion_visit_Farvardin_1399,
                 sima_sepehr_visit_Farvardin_1399, sima_shima_visit_Farvardin_1399, sima_site_visit_Farvardin_1399,],
       'register': [register_user_lenz_Farvardin_1399, register_user_aio_Farvardin_1399, register_user_anten_Farvardin_1399,
                 register_user_tva_Farvardin_1399, register_user_fam_Farvardin_1399, register_user_televebion_Farvardin_1399,
                 register_user_sepehr_Farvardin_1399, register_user_shima_Farvardin_1399, register_user_site_Farvardin_1399,],
       'active': [active_user_lenz_Farvardin_1399, active_user_aio_Farvardin_1399, active_user_anten_Farvardin_1399,
                 active_user_tva_Farvardin_1399, active_user_fam_Farvardin_1399, active_user_televebion_Farvardin_1399,
                 active_user_sepehr_Farvardin_1399, active_user_shima_Farvardin_1399, active_user_site_Farvardin_1399,],}

Farvardin_1399_operator_data=pd.DataFrame(Farvardin_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Farvardin_1399_operator_data=Farvardin_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Farvardin_1399_all_data_summary=pd.DataFrame()
Farvardin_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Farvardin_1399, all_duration_Farvardin_1399,all_content_sima_Farvardin_1399,
                      all_register_user_Farvardin_1399, all_active_user_Farvardin_1399,],}

Farvardin_1399_all_data_summary=pd.DataFrame(Farvardin_1399_all_data_summary, columns=['parameters', 'statistics'])

Farvardin_1399_all_data_summary=Farvardin_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه فروردین 1399.xlsx', engine='xlsxwriter')
Farvardin_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Farvardin_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Farvardin_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه فروردین')
writer.save()

        ########################### اردیبهشت #############################
print("EPG Ordibehesht 1399")
EPG_Ordibehesht_1399=pd.read_excel('EPG/EPG 1399/EPG Ordibehesht 1399.xlsx', sheet_name='آمار')
EPG_Ordibehesht_1399.fillna(0, inplace=True)
sima_1_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[1, 4]
sima_2_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[2, 4]
sima_3_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[3, 4]
sima_4_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[4, 4]
sima_5_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[5, 4]
sima_khabar_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[6, 4]
sima_ofogh_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[7, 4]
sima_pooya_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[8, 4]
sima_omid_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[9, 4]
sima_ifilm_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[10, 4]
sima_namayesh_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[11, 4]
sima_tamasha_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[12, 4]
sima_mostanad_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[13, 4]
sima_shoma_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[14, 4]
sima_amozesh_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[15, 4]
sima_varzesh_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[16, 4]
sima_nasim_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[17, 4]
sima_qoran_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[18, 4]
sima_salamat_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[19, 4]
sima_irankala_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[20, 4]
sima_alalam_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[21, 4]
sima_alkosar_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[22, 4]
sima_presstv_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[23, 4]
sima_sepehr_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[24, 4]

sima_1_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[1, 6]
sima_2_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[2, 6]
sima_3_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[3, 6]
sima_4_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[4, 6]
sima_5_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[5, 6]
sima_khabar_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[6, 6]
sima_ofogh_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[7, 6]
sima_pooya_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[8, 6]
sima_omid_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[9, 6]
sima_ifilm_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[10, 6]
sima_namayesh_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[11, 6]
sima_tamasha_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[12, 6]
sima_mostanad_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[13, 6]
sima_shoma_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[14, 6]
sima_amozesh_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[15, 6]
sima_varzesh_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[16, 6]
sima_nasim_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[17, 6]
sima_qoran_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[18, 6]
sima_salamat_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[19, 6]
sima_irankala_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[20, 6]
sima_alalam_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[21, 6]
sima_alkosar_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[22, 6]
sima_presstv_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[23, 6]
sima_sepehr_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[24, 6]

sima_lenz_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[33, 2]
sima_aio_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[34, 2]
sima_anten_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[35, 2]
sima_tva_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[36, 2]
sima_fam_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[37, 2]
sima_televebion_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[38, 2]
sima_sepehr_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[39, 2]
sima_shima_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[40, 2]
sima_site_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[41, 2]

register_user_lenz_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[36, 4]
register_user_aio_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[37, 4]
register_user_anten_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[38, 4]
register_user_tva_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[39, 4]
register_user_fam_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[40, 4]
register_user_televebion_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[41, 4]
register_user_sepehr_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[42, 4]
register_user_shima_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[43, 4]
register_user_site_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[44, 4]

active_user_lenz_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[36, 10]
active_user_aio_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[37, 10]
active_user_anten_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[38, 10]
active_user_tva_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[39, 10]
active_user_fam_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[40, 10]
active_user_televebion_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[41, 10]
active_user_sepehr_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[42, 10]
active_user_shima_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[43, 10]
active_user_site_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[44, 10]

all_visit_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[25, 4]
all_duration_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[25, 6]
all_content_sima_Ordibehesht_1399=EPG_Ordibehesht_1399.iat[25, 2]
all_register_user_Ordibehesht_1399=sum(EPG_Ordibehesht_1399.iloc[36:44, 4])
all_active_user_Ordibehesht_1399=sum(EPG_Ordibehesht_1399.iloc[36:44, 10])

Ordibehesht_1399_sima_visit_channels=pd.DataFrame()
Ordibehesht_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Ordibehesht_1399, sima_2_visit_Ordibehesht_1399, sima_3_visit_Ordibehesht_1399,
                 sima_4_visit_Ordibehesht_1399, sima_5_visit_Ordibehesht_1399, sima_khabar_visit_Ordibehesht_1399,
                 sima_ofogh_visit_Ordibehesht_1399, sima_pooya_visit_Ordibehesht_1399, sima_omid_visit_Ordibehesht_1399,
                 sima_ifilm_visit_Ordibehesht_1399, sima_namayesh_visit_Ordibehesht_1399, sima_tamasha_visit_Ordibehesht_1399,
                 sima_mostanad_visit_Ordibehesht_1399, sima_shoma_visit_Ordibehesht_1399, sima_amozesh_visit_Ordibehesht_1399,
                 sima_varzesh_visit_Ordibehesht_1399, sima_nasim_visit_Ordibehesht_1399, sima_qoran_visit_Ordibehesht_1399,
                 sima_salamat_visit_Ordibehesht_1399, sima_irankala_visit_Ordibehesht_1399, sima_alalam_visit_Ordibehesht_1399,
                 sima_alkosar_visit_Ordibehesht_1399, sima_presstv_visit_Ordibehesht_1399, sima_sepehr_visit_Ordibehesht_1399,],
        'duration': [sima_1_duration_Ordibehesht_1399, sima_2_duration_Ordibehesht_1399, sima_3_duration_Ordibehesht_1399,
                 sima_4_duration_Ordibehesht_1399, sima_5_duration_Ordibehesht_1399, sima_khabar_duration_Ordibehesht_1399,
                 sima_ofogh_duration_Ordibehesht_1399, sima_pooya_duration_Ordibehesht_1399, sima_omid_duration_Ordibehesht_1399,
                 sima_ifilm_duration_Ordibehesht_1399, sima_namayesh_duration_Ordibehesht_1399, sima_tamasha_duration_Ordibehesht_1399,
                 sima_mostanad_duration_Ordibehesht_1399, sima_shoma_duration_Ordibehesht_1399, sima_amozesh_duration_Ordibehesht_1399,
                 sima_varzesh_duration_Ordibehesht_1399, sima_nasim_duration_Ordibehesht_1399, sima_qoran_duration_Ordibehesht_1399,
                 sima_salamat_duration_Ordibehesht_1399, sima_irankala_duration_Ordibehesht_1399, sima_alalam_duration_Ordibehesht_1399,
                 sima_alkosar_duration_Ordibehesht_1399, sima_presstv_duration_Ordibehesht_1399, sima_sepehr_duration_Ordibehesht_1399,],}
Ordibehesht_1399_sima_visit_channels=pd.DataFrame(Ordibehesht_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Ordibehesht_1399_sima_visit_channels=Ordibehesht_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Ordibehesht_1399_operator_data=pd.DataFrame()
Ordibehesht_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Ordibehesht_1399, sima_aio_visit_Ordibehesht_1399, sima_anten_visit_Ordibehesht_1399,
                 sima_tva_visit_Ordibehesht_1399, sima_fam_visit_Ordibehesht_1399, sima_televebion_visit_Ordibehesht_1399,
                 sima_sepehr_visit_Ordibehesht_1399, sima_shima_visit_Ordibehesht_1399, sima_site_visit_Ordibehesht_1399,],
       'register': [register_user_lenz_Ordibehesht_1399, register_user_aio_Ordibehesht_1399, register_user_anten_Ordibehesht_1399,
                 register_user_tva_Ordibehesht_1399, register_user_fam_Ordibehesht_1399, register_user_televebion_Ordibehesht_1399,
                 register_user_sepehr_Ordibehesht_1399, register_user_shima_Ordibehesht_1399, register_user_site_Ordibehesht_1399,],
       'active': [active_user_lenz_Ordibehesht_1399, active_user_aio_Ordibehesht_1399, active_user_anten_Ordibehesht_1399,
                 active_user_tva_Ordibehesht_1399, active_user_fam_Ordibehesht_1399, active_user_televebion_Ordibehesht_1399,
                 active_user_sepehr_Ordibehesht_1399, active_user_shima_Ordibehesht_1399, active_user_site_Ordibehesht_1399,],}

Ordibehesht_1399_operator_data=pd.DataFrame(Ordibehesht_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Ordibehesht_1399_operator_data=Ordibehesht_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Ordibehesht_1399_all_data_summary=pd.DataFrame()
Ordibehesht_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Ordibehesht_1399, all_duration_Ordibehesht_1399,all_content_sima_Ordibehesht_1399,
                      all_register_user_Ordibehesht_1399, all_active_user_Ordibehesht_1399,],}

Ordibehesht_1399_all_data_summary=pd.DataFrame(Ordibehesht_1399_all_data_summary, columns=['parameters', 'statistics'])

Ordibehesht_1399_all_data_summary=Ordibehesht_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه اردیبهشت 1399.xlsx', engine='xlsxwriter')
Ordibehesht_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Ordibehesht_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Ordibehesht_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه اردیبهشت')
writer.save()

        ########################### خرداد #############################
print("EPG Khordad 1399")
EPG_Khordad_1399=pd.read_excel('EPG/EPG 1399/EPG Khordad 1399.xlsx', sheet_name='آمار')
EPG_Khordad_1399.fillna(0, inplace=True)
sima_1_visit_Khordad_1399=EPG_Khordad_1399.iat[1, 4]
sima_2_visit_Khordad_1399=EPG_Khordad_1399.iat[2, 4]
sima_3_visit_Khordad_1399=EPG_Khordad_1399.iat[3, 4]
sima_4_visit_Khordad_1399=EPG_Khordad_1399.iat[4, 4]
sima_5_visit_Khordad_1399=EPG_Khordad_1399.iat[5, 4]
sima_khabar_visit_Khordad_1399=EPG_Khordad_1399.iat[6, 4]
sima_ofogh_visit_Khordad_1399=EPG_Khordad_1399.iat[7, 4]
sima_pooya_visit_Khordad_1399=EPG_Khordad_1399.iat[8, 4]
sima_omid_visit_Khordad_1399=EPG_Khordad_1399.iat[9, 4]
sima_ifilm_visit_Khordad_1399=EPG_Khordad_1399.iat[10, 4]
sima_namayesh_visit_Khordad_1399=EPG_Khordad_1399.iat[11, 4]
sima_tamasha_visit_Khordad_1399=EPG_Khordad_1399.iat[12, 4]
sima_mostanad_visit_Khordad_1399=EPG_Khordad_1399.iat[13, 4]
sima_shoma_visit_Khordad_1399=EPG_Khordad_1399.iat[14, 4]
sima_amozesh_visit_Khordad_1399=EPG_Khordad_1399.iat[15, 4]
sima_varzesh_visit_Khordad_1399=EPG_Khordad_1399.iat[16, 4]
sima_nasim_visit_Khordad_1399=EPG_Khordad_1399.iat[17, 4]
sima_qoran_visit_Khordad_1399=EPG_Khordad_1399.iat[18, 4]
sima_salamat_visit_Khordad_1399=EPG_Khordad_1399.iat[19, 4]
sima_irankala_visit_Khordad_1399=EPG_Khordad_1399.iat[20, 4]
sima_alalam_visit_Khordad_1399=EPG_Khordad_1399.iat[21, 4]
sima_alkosar_visit_Khordad_1399=EPG_Khordad_1399.iat[22, 4]
sima_presstv_visit_Khordad_1399=EPG_Khordad_1399.iat[23, 4]
sima_sepehr_visit_Khordad_1399=EPG_Khordad_1399.iat[24, 4]

sima_1_duration_Khordad_1399=EPG_Khordad_1399.iat[1, 6]
sima_2_duration_Khordad_1399=EPG_Khordad_1399.iat[2, 6]
sima_3_duration_Khordad_1399=EPG_Khordad_1399.iat[3, 6]
sima_4_duration_Khordad_1399=EPG_Khordad_1399.iat[4, 6]
sima_5_duration_Khordad_1399=EPG_Khordad_1399.iat[5, 6]
sima_khabar_duration_Khordad_1399=EPG_Khordad_1399.iat[6, 6]
sima_ofogh_duration_Khordad_1399=EPG_Khordad_1399.iat[7, 6]
sima_pooya_duration_Khordad_1399=EPG_Khordad_1399.iat[8, 6]
sima_omid_duration_Khordad_1399=EPG_Khordad_1399.iat[9, 6]
sima_ifilm_duration_Khordad_1399=EPG_Khordad_1399.iat[10, 6]
sima_namayesh_duration_Khordad_1399=EPG_Khordad_1399.iat[11, 6]
sima_tamasha_duration_Khordad_1399=EPG_Khordad_1399.iat[12, 6]
sima_mostanad_duration_Khordad_1399=EPG_Khordad_1399.iat[13, 6]
sima_shoma_duration_Khordad_1399=EPG_Khordad_1399.iat[14, 6]
sima_amozesh_duration_Khordad_1399=EPG_Khordad_1399.iat[15, 6]
sima_varzesh_duration_Khordad_1399=EPG_Khordad_1399.iat[16, 6]
sima_nasim_duration_Khordad_1399=EPG_Khordad_1399.iat[17, 6]
sima_qoran_duration_Khordad_1399=EPG_Khordad_1399.iat[18, 6]
sima_salamat_duration_Khordad_1399=EPG_Khordad_1399.iat[19, 6]
sima_irankala_duration_Khordad_1399=EPG_Khordad_1399.iat[20, 6]
sima_alalam_duration_Khordad_1399=EPG_Khordad_1399.iat[21, 6]
sima_alkosar_duration_Khordad_1399=EPG_Khordad_1399.iat[22, 6]
sima_presstv_duration_Khordad_1399=EPG_Khordad_1399.iat[23, 6]
sima_sepehr_duration_Khordad_1399=EPG_Khordad_1399.iat[24, 6]

sima_lenz_visit_Khordad_1399=EPG_Khordad_1399.iat[33, 2]
sima_aio_visit_Khordad_1399=EPG_Khordad_1399.iat[34, 2]
sima_anten_visit_Khordad_1399=EPG_Khordad_1399.iat[35, 2]
sima_tva_visit_Khordad_1399=EPG_Khordad_1399.iat[36, 2]
sima_fam_visit_Khordad_1399=EPG_Khordad_1399.iat[37, 2]
sima_televebion_visit_Khordad_1399=EPG_Khordad_1399.iat[38, 2]
sima_sepehr_Khordad_1399=EPG_Khordad_1399.iat[39, 2]
sima_shima_visit_Khordad_1399=EPG_Khordad_1399.iat[40, 2]
sima_site_visit_Khordad_1399=EPG_Khordad_1399.iat[41, 2]

register_user_lenz_Khordad_1399=EPG_Khordad_1399.iat[36, 4]
register_user_aio_Khordad_1399=EPG_Khordad_1399.iat[37, 4]
register_user_anten_Khordad_1399=EPG_Khordad_1399.iat[38, 4]
register_user_tva_Khordad_1399=EPG_Khordad_1399.iat[39, 4]
register_user_fam_Khordad_1399=EPG_Khordad_1399.iat[40, 4]
register_user_televebion_Khordad_1399=EPG_Khordad_1399.iat[41, 4]
register_user_sepehr_Khordad_1399=EPG_Khordad_1399.iat[42, 4]
register_user_shima_Khordad_1399=EPG_Khordad_1399.iat[43, 4]
register_user_site_Khordad_1399=EPG_Khordad_1399.iat[44, 4]

active_user_lenz_Khordad_1399=EPG_Khordad_1399.iat[36, 10]
active_user_aio_Khordad_1399=EPG_Khordad_1399.iat[37, 10]
active_user_anten_Khordad_1399=EPG_Khordad_1399.iat[38, 10]
active_user_tva_Khordad_1399=EPG_Khordad_1399.iat[39, 10]
active_user_fam_Khordad_1399=EPG_Khordad_1399.iat[40, 10]
active_user_televebion_Khordad_1399=EPG_Khordad_1399.iat[41, 10]
active_user_sepehr_Khordad_1399=EPG_Khordad_1399.iat[42, 10]
active_user_shima_Khordad_1399=EPG_Khordad_1399.iat[43, 10]
active_user_site_Khordad_1399=EPG_Khordad_1399.iat[44, 10]

all_visit_Khordad_1399=EPG_Khordad_1399.iat[25, 4]
all_duration_Khordad_1399=EPG_Khordad_1399.iat[25, 6]
all_content_sima_Khordad_1399=EPG_Khordad_1399.iat[25, 2]
all_register_user_Khordad_1399=sum(EPG_Khordad_1399.iloc[36:44, 4])
all_active_user_Khordad_1399=sum(EPG_Khordad_1399.iloc[36:44, 10])

Khordad_1399_sima_visit_channels=pd.DataFrame()
Khordad_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Khordad_1399, sima_2_visit_Khordad_1399, sima_3_visit_Khordad_1399,
                 sima_4_visit_Khordad_1399, sima_5_visit_Khordad_1399, sima_khabar_visit_Khordad_1399,
                 sima_ofogh_visit_Khordad_1399, sima_pooya_visit_Khordad_1399, sima_omid_visit_Khordad_1399,
                 sima_ifilm_visit_Khordad_1399, sima_namayesh_visit_Khordad_1399, sima_tamasha_visit_Khordad_1399,
                 sima_mostanad_visit_Khordad_1399, sima_shoma_visit_Khordad_1399, sima_amozesh_visit_Khordad_1399,
                 sima_varzesh_visit_Khordad_1399, sima_nasim_visit_Khordad_1399, sima_qoran_visit_Khordad_1399,
                 sima_salamat_visit_Khordad_1399, sima_irankala_visit_Khordad_1399, sima_alalam_visit_Khordad_1399,
                 sima_alkosar_visit_Khordad_1399, sima_presstv_visit_Khordad_1399, sima_sepehr_visit_Khordad_1399,],
        'duration': [sima_1_duration_Khordad_1399, sima_2_duration_Khordad_1399, sima_3_duration_Khordad_1399,
                 sima_4_duration_Khordad_1399, sima_5_duration_Khordad_1399, sima_khabar_duration_Khordad_1399,
                 sima_ofogh_duration_Khordad_1399, sima_pooya_duration_Khordad_1399, sima_omid_duration_Khordad_1399,
                 sima_ifilm_duration_Khordad_1399, sima_namayesh_duration_Khordad_1399, sima_tamasha_duration_Khordad_1399,
                 sima_mostanad_duration_Khordad_1399, sima_shoma_duration_Khordad_1399, sima_amozesh_duration_Khordad_1399,
                 sima_varzesh_duration_Khordad_1399, sima_nasim_duration_Khordad_1399, sima_qoran_duration_Khordad_1399,
                 sima_salamat_duration_Khordad_1399, sima_irankala_duration_Khordad_1399, sima_alalam_duration_Khordad_1399,
                 sima_alkosar_duration_Khordad_1399, sima_presstv_duration_Khordad_1399, sima_sepehr_duration_Khordad_1399,],}
Khordad_1399_sima_visit_channels=pd.DataFrame(Khordad_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Khordad_1399_sima_visit_channels=Khordad_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Khordad_1399_operator_data=pd.DataFrame()
Khordad_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Khordad_1399, sima_aio_visit_Khordad_1399, sima_anten_visit_Khordad_1399,
                 sima_tva_visit_Khordad_1399, sima_fam_visit_Khordad_1399, sima_televebion_visit_Khordad_1399,
                 sima_sepehr_visit_Khordad_1399, sima_shima_visit_Khordad_1399, sima_site_visit_Khordad_1399,],
       'register': [register_user_lenz_Khordad_1399, register_user_aio_Khordad_1399, register_user_anten_Khordad_1399,
                 register_user_tva_Khordad_1399, register_user_fam_Khordad_1399, register_user_televebion_Khordad_1399,
                 register_user_sepehr_Khordad_1399, register_user_shima_Khordad_1399, register_user_site_Khordad_1399,],
       'active': [active_user_lenz_Khordad_1399, active_user_aio_Khordad_1399, active_user_anten_Khordad_1399,
                 active_user_tva_Khordad_1399, active_user_fam_Khordad_1399, active_user_televebion_Khordad_1399,
                 active_user_sepehr_Khordad_1399, active_user_shima_Khordad_1399, active_user_site_Khordad_1399,],}

Khordad_1399_operator_data=pd.DataFrame(Khordad_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Khordad_1399_operator_data=Khordad_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Khordad_1399_all_data_summary=pd.DataFrame()
Khordad_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Khordad_1399, all_duration_Khordad_1399,all_content_sima_Khordad_1399,
                      all_register_user_Khordad_1399, all_active_user_Khordad_1399,],}

Khordad_1399_all_data_summary=pd.DataFrame(Khordad_1399_all_data_summary, columns=['parameters', 'statistics'])

Khordad_1399_all_data_summary=Khordad_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه خرداد 1399.xlsx', engine='xlsxwriter')
Khordad_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Khordad_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Khordad_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه خرداد')
writer.save()

        ########################### تیر #############################
print("EPG Tir 1399")
EPG_Tir_1399=pd.read_excel('EPG/EPG 1399/EPG Tir 1399.xlsx', sheet_name='آمار')
EPG_Tir_1399.fillna(0, inplace=True)
sima_1_visit_Tir_1399=EPG_Tir_1399.iat[1, 4]
sima_2_visit_Tir_1399=EPG_Tir_1399.iat[2, 4]
sima_3_visit_Tir_1399=EPG_Tir_1399.iat[3, 4]
sima_4_visit_Tir_1399=EPG_Tir_1399.iat[4, 4]
sima_5_visit_Tir_1399=EPG_Tir_1399.iat[5, 4]
sima_khabar_visit_Tir_1399=EPG_Tir_1399.iat[6, 4]
sima_ofogh_visit_Tir_1399=EPG_Tir_1399.iat[7, 4]
sima_pooya_visit_Tir_1399=EPG_Tir_1399.iat[8, 4]
sima_omid_visit_Tir_1399=EPG_Tir_1399.iat[9, 4]
sima_ifilm_visit_Tir_1399=EPG_Tir_1399.iat[10, 4]
sima_namayesh_visit_Tir_1399=EPG_Tir_1399.iat[11, 4]
sima_tamasha_visit_Tir_1399=EPG_Tir_1399.iat[12, 4]
sima_mostanad_visit_Tir_1399=EPG_Tir_1399.iat[13, 4]
sima_shoma_visit_Tir_1399=EPG_Tir_1399.iat[14, 4]
sima_amozesh_visit_Tir_1399=EPG_Tir_1399.iat[15, 4]
sima_varzesh_visit_Tir_1399=EPG_Tir_1399.iat[16, 4]
sima_nasim_visit_Tir_1399=EPG_Tir_1399.iat[17, 4]
sima_qoran_visit_Tir_1399=EPG_Tir_1399.iat[18, 4]
sima_salamat_visit_Tir_1399=EPG_Tir_1399.iat[19, 4]
sima_irankala_visit_Tir_1399=EPG_Tir_1399.iat[20, 4]
sima_alalam_visit_Tir_1399=EPG_Tir_1399.iat[21, 4]
sima_alkosar_visit_Tir_1399=EPG_Tir_1399.iat[22, 4]
sima_presstv_visit_Tir_1399=EPG_Tir_1399.iat[23, 4]
sima_sepehr_visit_Tir_1399=EPG_Tir_1399.iat[24, 4]

sima_1_duration_Tir_1399=EPG_Tir_1399.iat[1, 6]
sima_2_duration_Tir_1399=EPG_Tir_1399.iat[2, 6]
sima_3_duration_Tir_1399=EPG_Tir_1399.iat[3, 6]
sima_4_duration_Tir_1399=EPG_Tir_1399.iat[4, 6]
sima_5_duration_Tir_1399=EPG_Tir_1399.iat[5, 6]
sima_khabar_duration_Tir_1399=EPG_Tir_1399.iat[6, 6]
sima_ofogh_duration_Tir_1399=EPG_Tir_1399.iat[7, 6]
sima_pooya_duration_Tir_1399=EPG_Tir_1399.iat[8, 6]
sima_omid_duration_Tir_1399=EPG_Tir_1399.iat[9, 6]
sima_ifilm_duration_Tir_1399=EPG_Tir_1399.iat[10, 6]
sima_namayesh_duration_Tir_1399=EPG_Tir_1399.iat[11, 6]
sima_tamasha_duration_Tir_1399=EPG_Tir_1399.iat[12, 6]
sima_mostanad_duration_Tir_1399=EPG_Tir_1399.iat[13, 6]
sima_shoma_duration_Tir_1399=EPG_Tir_1399.iat[14, 6]
sima_amozesh_duration_Tir_1399=EPG_Tir_1399.iat[15, 6]
sima_varzesh_duration_Tir_1399=EPG_Tir_1399.iat[16, 6]
sima_nasim_duration_Tir_1399=EPG_Tir_1399.iat[17, 6]
sima_qoran_duration_Tir_1399=EPG_Tir_1399.iat[18, 6]
sima_salamat_duration_Tir_1399=EPG_Tir_1399.iat[19, 6]
sima_irankala_duration_Tir_1399=EPG_Tir_1399.iat[20, 6]
sima_alalam_duration_Tir_1399=EPG_Tir_1399.iat[21, 6]
sima_alkosar_duration_Tir_1399=EPG_Tir_1399.iat[22, 6]
sima_presstv_duration_Tir_1399=EPG_Tir_1399.iat[23, 6]
sima_sepehr_duration_Tir_1399=EPG_Tir_1399.iat[24, 6]

sima_lenz_visit_Tir_1399=EPG_Tir_1399.iat[33, 2]
sima_aio_visit_Tir_1399=EPG_Tir_1399.iat[34, 2]
sima_anten_visit_Tir_1399=EPG_Tir_1399.iat[35, 2]
sima_tva_visit_Tir_1399=EPG_Tir_1399.iat[36, 2]
sima_fam_visit_Tir_1399=EPG_Tir_1399.iat[37, 2]
sima_televebion_visit_Tir_1399=EPG_Tir_1399.iat[38, 2]
sima_sepehr_Tir_1399=EPG_Tir_1399.iat[39, 2]
sima_shima_visit_Tir_1399=EPG_Tir_1399.iat[40, 2]
sima_site_visit_Tir_1399=EPG_Tir_1399.iat[41, 2]

register_user_lenz_Tir_1399=EPG_Tir_1399.iat[36, 4]
register_user_aio_Tir_1399=EPG_Tir_1399.iat[37, 4]
register_user_anten_Tir_1399=EPG_Tir_1399.iat[38, 4]
register_user_tva_Tir_1399=EPG_Tir_1399.iat[39, 4]
register_user_fam_Tir_1399=EPG_Tir_1399.iat[40, 4]
register_user_televebion_Tir_1399=EPG_Tir_1399.iat[41, 4]
register_user_sepehr_Tir_1399=EPG_Tir_1399.iat[42, 4]
register_user_shima_Tir_1399=EPG_Tir_1399.iat[43, 4]
register_user_site_Tir_1399=EPG_Tir_1399.iat[44, 4]

active_user_lenz_Tir_1399=EPG_Tir_1399.iat[36, 10]
active_user_aio_Tir_1399=EPG_Tir_1399.iat[37, 10]
active_user_anten_Tir_1399=EPG_Tir_1399.iat[38, 10]
active_user_tva_Tir_1399=EPG_Tir_1399.iat[39, 10]
active_user_fam_Tir_1399=EPG_Tir_1399.iat[40, 10]
active_user_televebion_Tir_1399=EPG_Tir_1399.iat[41, 10]
active_user_sepehr_Tir_1399=EPG_Tir_1399.iat[42, 10]
active_user_shima_Tir_1399=EPG_Tir_1399.iat[43, 10]
active_user_site_Tir_1399=EPG_Tir_1399.iat[44, 10]

all_visit_Tir_1399=EPG_Tir_1399.iat[25, 4]
all_duration_Tir_1399=EPG_Tir_1399.iat[25, 6]
all_content_sima_Tir_1399=EPG_Tir_1399.iat[25, 2]
all_register_user_Tir_1399=sum(EPG_Tir_1399.iloc[36:44, 4])
all_active_user_Tir_1399=sum(EPG_Tir_1399.iloc[36:44, 10])

Tir_1399_sima_visit_channels=pd.DataFrame()
Tir_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Tir_1399, sima_2_visit_Tir_1399, sima_3_visit_Tir_1399,
                 sima_4_visit_Tir_1399, sima_5_visit_Tir_1399, sima_khabar_visit_Tir_1399,
                 sima_ofogh_visit_Tir_1399, sima_pooya_visit_Tir_1399, sima_omid_visit_Tir_1399,
                 sima_ifilm_visit_Tir_1399, sima_namayesh_visit_Tir_1399, sima_tamasha_visit_Tir_1399,
                 sima_mostanad_visit_Tir_1399, sima_shoma_visit_Tir_1399, sima_amozesh_visit_Tir_1399,
                 sima_varzesh_visit_Tir_1399, sima_nasim_visit_Tir_1399, sima_qoran_visit_Tir_1399,
                 sima_salamat_visit_Tir_1399, sima_irankala_visit_Tir_1399, sima_alalam_visit_Tir_1399,
                 sima_alkosar_visit_Tir_1399, sima_presstv_visit_Tir_1399, sima_sepehr_visit_Tir_1399,],
        'duration': [sima_1_duration_Tir_1399, sima_2_duration_Tir_1399, sima_3_duration_Tir_1399,
                 sima_4_duration_Tir_1399, sima_5_duration_Tir_1399, sima_khabar_duration_Tir_1399,
                 sima_ofogh_duration_Tir_1399, sima_pooya_duration_Tir_1399, sima_omid_duration_Tir_1399,
                 sima_ifilm_duration_Tir_1399, sima_namayesh_duration_Tir_1399, sima_tamasha_duration_Tir_1399,
                 sima_mostanad_duration_Tir_1399, sima_shoma_duration_Tir_1399, sima_amozesh_duration_Tir_1399,
                 sima_varzesh_duration_Tir_1399, sima_nasim_duration_Tir_1399, sima_qoran_duration_Tir_1399,
                 sima_salamat_duration_Tir_1399, sima_irankala_duration_Tir_1399, sima_alalam_duration_Tir_1399,
                 sima_alkosar_duration_Tir_1399, sima_presstv_duration_Tir_1399, sima_sepehr_duration_Tir_1399,],}
Tir_1399_sima_visit_channels=pd.DataFrame(Tir_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Tir_1399_sima_visit_channels=Tir_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Tir_1399_operator_data=pd.DataFrame()
Tir_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Tir_1399, sima_aio_visit_Tir_1399, sima_anten_visit_Tir_1399,
                 sima_tva_visit_Tir_1399, sima_fam_visit_Tir_1399, sima_televebion_visit_Tir_1399,
                 sima_sepehr_visit_Tir_1399, sima_shima_visit_Tir_1399, sima_site_visit_Tir_1399,],
       'register': [register_user_lenz_Tir_1399, register_user_aio_Tir_1399, register_user_anten_Tir_1399,
                 register_user_tva_Tir_1399, register_user_fam_Tir_1399, register_user_televebion_Tir_1399,
                 register_user_sepehr_Tir_1399, register_user_shima_Tir_1399, register_user_site_Tir_1399,],
       'active': [active_user_lenz_Tir_1399, active_user_aio_Tir_1399, active_user_anten_Tir_1399,
                 active_user_tva_Tir_1399, active_user_fam_Tir_1399, active_user_televebion_Tir_1399,
                 active_user_sepehr_Tir_1399, active_user_shima_Tir_1399, active_user_site_Tir_1399,],}

Tir_1399_operator_data=pd.DataFrame(Tir_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Tir_1399_operator_data=Tir_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Tir_1399_all_data_summary=pd.DataFrame()
Tir_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Tir_1399, all_duration_Tir_1399,all_content_sima_Tir_1399,
                      all_register_user_Tir_1399, all_active_user_Tir_1399,],}

Tir_1399_all_data_summary=pd.DataFrame(Tir_1399_all_data_summary, columns=['parameters', 'statistics'])

Tir_1399_all_data_summary=Tir_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه تیر 1399.xlsx', engine='xlsxwriter')
Tir_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Tir_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Tir_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه تیر')
writer.save()

        ########################### مرداد #############################
print("EPG Mordad 1399")
EPG_Mordad_1399=pd.read_excel('EPG/EPG 1399/EPG Mordad 1399.xlsx', sheet_name='آمار')
EPG_Mordad_1399.fillna(0, inplace=True)
sima_1_visit_Mordad_1399=EPG_Mordad_1399.iat[1, 4]
sima_2_visit_Mordad_1399=EPG_Mordad_1399.iat[2, 4]
sima_3_visit_Mordad_1399=EPG_Mordad_1399.iat[3, 4]
sima_4_visit_Mordad_1399=EPG_Mordad_1399.iat[4, 4]
sima_5_visit_Mordad_1399=EPG_Mordad_1399.iat[5, 4]
sima_khabar_visit_Mordad_1399=EPG_Mordad_1399.iat[6, 4]
sima_ofogh_visit_Mordad_1399=EPG_Mordad_1399.iat[7, 4]
sima_pooya_visit_Mordad_1399=EPG_Mordad_1399.iat[8, 4]
sima_omid_visit_Mordad_1399=EPG_Mordad_1399.iat[9, 4]
sima_ifilm_visit_Mordad_1399=EPG_Mordad_1399.iat[10, 4]
sima_namayesh_visit_Mordad_1399=EPG_Mordad_1399.iat[11, 4]
sima_tamasha_visit_Mordad_1399=EPG_Mordad_1399.iat[12, 4]
sima_mostanad_visit_Mordad_1399=EPG_Mordad_1399.iat[13, 4]
sima_shoma_visit_Mordad_1399=EPG_Mordad_1399.iat[14, 4]
sima_amozesh_visit_Mordad_1399=EPG_Mordad_1399.iat[15, 4]
sima_varzesh_visit_Mordad_1399=EPG_Mordad_1399.iat[16, 4]
sima_nasim_visit_Mordad_1399=EPG_Mordad_1399.iat[17, 4]
sima_qoran_visit_Mordad_1399=EPG_Mordad_1399.iat[18, 4]
sima_salamat_visit_Mordad_1399=EPG_Mordad_1399.iat[19, 4]
sima_irankala_visit_Mordad_1399=EPG_Mordad_1399.iat[20, 4]
sima_alalam_visit_Mordad_1399=EPG_Mordad_1399.iat[21, 4]
sima_alkosar_visit_Mordad_1399=EPG_Mordad_1399.iat[22, 4]
sima_presstv_visit_Mordad_1399=EPG_Mordad_1399.iat[23, 4]
sima_sepehr_visit_Mordad_1399=EPG_Mordad_1399.iat[24, 4]

sima_1_duration_Mordad_1399=EPG_Mordad_1399.iat[1, 6]
sima_2_duration_Mordad_1399=EPG_Mordad_1399.iat[2, 6]
sima_3_duration_Mordad_1399=EPG_Mordad_1399.iat[3, 6]
sima_4_duration_Mordad_1399=EPG_Mordad_1399.iat[4, 6]
sima_5_duration_Mordad_1399=EPG_Mordad_1399.iat[5, 6]
sima_khabar_duration_Mordad_1399=EPG_Mordad_1399.iat[6, 6]
sima_ofogh_duration_Mordad_1399=EPG_Mordad_1399.iat[7, 6]
sima_pooya_duration_Mordad_1399=EPG_Mordad_1399.iat[8, 6]
sima_omid_duration_Mordad_1399=EPG_Mordad_1399.iat[9, 6]
sima_ifilm_duration_Mordad_1399=EPG_Mordad_1399.iat[10, 6]
sima_namayesh_duration_Mordad_1399=EPG_Mordad_1399.iat[11, 6]
sima_tamasha_duration_Mordad_1399=EPG_Mordad_1399.iat[12, 6]
sima_mostanad_duration_Mordad_1399=EPG_Mordad_1399.iat[13, 6]
sima_shoma_duration_Mordad_1399=EPG_Mordad_1399.iat[14, 6]
sima_amozesh_duration_Mordad_1399=EPG_Mordad_1399.iat[15, 6]
sima_varzesh_duration_Mordad_1399=EPG_Mordad_1399.iat[16, 6]
sima_nasim_duration_Mordad_1399=EPG_Mordad_1399.iat[17, 6]
sima_qoran_duration_Mordad_1399=EPG_Mordad_1399.iat[18, 6]
sima_salamat_duration_Mordad_1399=EPG_Mordad_1399.iat[19, 6]
sima_irankala_duration_Mordad_1399=EPG_Mordad_1399.iat[20, 6]
sima_alalam_duration_Mordad_1399=EPG_Mordad_1399.iat[21, 6]
sima_alkosar_duration_Mordad_1399=EPG_Mordad_1399.iat[22, 6]
sima_presstv_duration_Mordad_1399=EPG_Mordad_1399.iat[23, 6]
sima_sepehr_duration_Mordad_1399=EPG_Mordad_1399.iat[24, 6]

sima_lenz_visit_Mordad_1399=EPG_Mordad_1399.iat[33, 2]
sima_aio_visit_Mordad_1399=EPG_Mordad_1399.iat[34, 2]
sima_anten_visit_Mordad_1399=EPG_Mordad_1399.iat[35, 2]
sima_tva_visit_Mordad_1399=EPG_Mordad_1399.iat[36, 2]
sima_fam_visit_Mordad_1399=EPG_Mordad_1399.iat[37, 2]
sima_televebion_visit_Mordad_1399=EPG_Mordad_1399.iat[38, 2]
sima_sepehr_Mordad_1399=EPG_Mordad_1399.iat[39, 2]
sima_shima_visit_Mordad_1399=EPG_Mordad_1399.iat[40, 2]
sima_site_visit_Mordad_1399=EPG_Mordad_1399.iat[41, 2]

register_user_lenz_Mordad_1399=EPG_Mordad_1399.iat[36, 4]
register_user_aio_Mordad_1399=EPG_Mordad_1399.iat[37, 4]
register_user_anten_Mordad_1399=EPG_Mordad_1399.iat[38, 4]
register_user_tva_Mordad_1399=EPG_Mordad_1399.iat[39, 4]
register_user_fam_Mordad_1399=EPG_Mordad_1399.iat[40, 4]
register_user_televebion_Mordad_1399=EPG_Mordad_1399.iat[41, 4]
register_user_sepehr_Mordad_1399=EPG_Mordad_1399.iat[42, 4]
register_user_shima_Mordad_1399=EPG_Mordad_1399.iat[43, 4]
register_user_site_Mordad_1399=EPG_Mordad_1399.iat[44, 4]

active_user_lenz_Mordad_1399=EPG_Mordad_1399.iat[36, 10]
active_user_aio_Mordad_1399=EPG_Mordad_1399.iat[37, 10]
active_user_anten_Mordad_1399=EPG_Mordad_1399.iat[38, 10]
active_user_tva_Mordad_1399=EPG_Mordad_1399.iat[39, 10]
active_user_fam_Mordad_1399=EPG_Mordad_1399.iat[40, 10]
active_user_televebion_Mordad_1399=EPG_Mordad_1399.iat[41, 10]
active_user_sepehr_Mordad_1399=EPG_Mordad_1399.iat[42, 10]
active_user_shima_Mordad_1399=EPG_Mordad_1399.iat[43, 10]
active_user_site_Mordad_1399=EPG_Mordad_1399.iat[44, 10]

all_visit_Mordad_1399=EPG_Mordad_1399.iat[25, 4]
all_duration_Mordad_1399=EPG_Mordad_1399.iat[25, 6]
all_content_sima_Mordad_1399=EPG_Mordad_1399.iat[25, 2]
all_register_user_Mordad_1399=sum(EPG_Mordad_1399.iloc[36:44, 4])
all_active_user_Mordad_1399=sum(EPG_Mordad_1399.iloc[36:44, 10])

Mordad_1399_sima_visit_channels=pd.DataFrame()
Mordad_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Mordad_1399, sima_2_visit_Mordad_1399, sima_3_visit_Mordad_1399,
                 sima_4_visit_Mordad_1399, sima_5_visit_Mordad_1399, sima_khabar_visit_Mordad_1399,
                 sima_ofogh_visit_Mordad_1399, sima_pooya_visit_Mordad_1399, sima_omid_visit_Mordad_1399,
                 sima_ifilm_visit_Mordad_1399, sima_namayesh_visit_Mordad_1399, sima_tamasha_visit_Mordad_1399,
                 sima_mostanad_visit_Mordad_1399, sima_shoma_visit_Mordad_1399, sima_amozesh_visit_Mordad_1399,
                 sima_varzesh_visit_Mordad_1399, sima_nasim_visit_Mordad_1399, sima_qoran_visit_Mordad_1399,
                 sima_salamat_visit_Mordad_1399, sima_irankala_visit_Mordad_1399, sima_alalam_visit_Mordad_1399,
                 sima_alkosar_visit_Mordad_1399, sima_presstv_visit_Mordad_1399, sima_sepehr_visit_Mordad_1399,],
        'duration': [sima_1_duration_Mordad_1399, sima_2_duration_Mordad_1399, sima_3_duration_Mordad_1399,
                 sima_4_duration_Mordad_1399, sima_5_duration_Mordad_1399, sima_khabar_duration_Mordad_1399,
                 sima_ofogh_duration_Mordad_1399, sima_pooya_duration_Mordad_1399, sima_omid_duration_Mordad_1399,
                 sima_ifilm_duration_Mordad_1399, sima_namayesh_duration_Mordad_1399, sima_tamasha_duration_Mordad_1399,
                 sima_mostanad_duration_Mordad_1399, sima_shoma_duration_Mordad_1399, sima_amozesh_duration_Mordad_1399,
                 sima_varzesh_duration_Mordad_1399, sima_nasim_duration_Mordad_1399, sima_qoran_duration_Mordad_1399,
                 sima_salamat_duration_Mordad_1399, sima_irankala_duration_Mordad_1399, sima_alalam_duration_Mordad_1399,
                 sima_alkosar_duration_Mordad_1399, sima_presstv_duration_Mordad_1399, sima_sepehr_duration_Mordad_1399,],}
Mordad_1399_sima_visit_channels=pd.DataFrame(Mordad_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Mordad_1399_sima_visit_channels=Mordad_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Mordad_1399_operator_data=pd.DataFrame()
Mordad_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Mordad_1399, sima_aio_visit_Mordad_1399, sima_anten_visit_Mordad_1399,
                 sima_tva_visit_Mordad_1399, sima_fam_visit_Mordad_1399, sima_televebion_visit_Mordad_1399,
                 sima_sepehr_visit_Mordad_1399, sima_shima_visit_Mordad_1399, sima_site_visit_Mordad_1399,],
       'register': [register_user_lenz_Mordad_1399, register_user_aio_Mordad_1399, register_user_anten_Mordad_1399,
                 register_user_tva_Mordad_1399, register_user_fam_Mordad_1399, register_user_televebion_Mordad_1399,
                 register_user_sepehr_Mordad_1399, register_user_shima_Mordad_1399, register_user_site_Mordad_1399,],
       'active': [active_user_lenz_Mordad_1399, active_user_aio_Mordad_1399, active_user_anten_Mordad_1399,
                 active_user_tva_Mordad_1399, active_user_fam_Mordad_1399, active_user_televebion_Mordad_1399,
                 active_user_sepehr_Mordad_1399, active_user_shima_Mordad_1399, active_user_site_Mordad_1399,],}

Mordad_1399_operator_data=pd.DataFrame(Mordad_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Mordad_1399_operator_data=Mordad_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Mordad_1399_all_data_summary=pd.DataFrame()
Mordad_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Mordad_1399, all_duration_Mordad_1399,all_content_sima_Mordad_1399,
                      all_register_user_Mordad_1399, all_active_user_Mordad_1399,],}

Mordad_1399_all_data_summary=pd.DataFrame(Mordad_1399_all_data_summary, columns=['parameters', 'statistics'])

Mordad_1399_all_data_summary=Mordad_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه مرداد 1399.xlsx', engine='xlsxwriter')
Mordad_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Mordad_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Mordad_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه مرداد')
writer.save()

        ########################### شهریور #############################
print("EPG Shahrivar 1399")
EPG_Shahrivar_1399=pd.read_excel('EPG/EPG 1399/EPG Shahrivar 1399.xlsx', sheet_name='آمار')
EPG_Shahrivar_1399.fillna(0, inplace=True)
sima_1_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[1, 4]
sima_2_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[2, 4]
sima_3_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[3, 4]
sima_4_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[4, 4]
sima_5_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[5, 4]
sima_khabar_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[6, 4]
sima_ofogh_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[7, 4]
sima_pooya_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[8, 4]
sima_omid_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[9, 4]
sima_ifilm_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[10, 4]
sima_namayesh_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[11, 4]
sima_tamasha_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[12, 4]
sima_mostanad_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[13, 4]
sima_shoma_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[14, 4]
sima_amozesh_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[15, 4]
sima_varzesh_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[16, 4]
sima_nasim_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[17, 4]
sima_qoran_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[18, 4]
sima_salamat_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[19, 4]
sima_irankala_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[20, 4]
sima_alalam_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[21, 4]
sima_alkosar_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[22, 4]
sima_presstv_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[23, 4]
sima_sepehr_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[24, 4]

sima_1_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[1, 6]
sima_2_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[2, 6]
sima_3_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[3, 6]
sima_4_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[4, 6]
sima_5_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[5, 6]
sima_khabar_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[6, 6]
sima_ofogh_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[7, 6]
sima_pooya_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[8, 6]
sima_omid_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[9, 6]
sima_ifilm_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[10, 6]
sima_namayesh_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[11, 6]
sima_tamasha_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[12, 6]
sima_mostanad_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[13, 6]
sima_shoma_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[14, 6]
sima_amozesh_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[15, 6]
sima_varzesh_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[16, 6]
sima_nasim_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[17, 6]
sima_qoran_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[18, 6]
sima_salamat_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[19, 6]
sima_irankala_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[20, 6]
sima_alalam_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[21, 6]
sima_alkosar_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[22, 6]
sima_presstv_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[23, 6]
sima_sepehr_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[24, 6]

sima_lenz_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[33, 2]
sima_aio_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[34, 2]
sima_anten_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[35, 2]
sima_tva_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[36, 2]
sima_fam_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[37, 2]
sima_televebion_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[38, 2]
sima_sepehr_Shahrivar_1399=EPG_Shahrivar_1399.iat[39, 2]
sima_shima_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[40, 2]
sima_site_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[41, 2]

register_user_lenz_Shahrivar_1399=EPG_Shahrivar_1399.iat[36, 4]
register_user_aio_Shahrivar_1399=EPG_Shahrivar_1399.iat[37, 4]
register_user_anten_Shahrivar_1399=EPG_Shahrivar_1399.iat[38, 4]
register_user_tva_Shahrivar_1399=EPG_Shahrivar_1399.iat[39, 4]
register_user_fam_Shahrivar_1399=EPG_Shahrivar_1399.iat[40, 4]
register_user_televebion_Shahrivar_1399=EPG_Shahrivar_1399.iat[41, 4]
register_user_sepehr_Shahrivar_1399=EPG_Shahrivar_1399.iat[42, 4]
register_user_shima_Shahrivar_1399=EPG_Shahrivar_1399.iat[43, 4]
register_user_site_Shahrivar_1399=EPG_Shahrivar_1399.iat[44, 4]

active_user_lenz_Shahrivar_1399=EPG_Shahrivar_1399.iat[36, 10]
active_user_aio_Shahrivar_1399=EPG_Shahrivar_1399.iat[37, 10]
active_user_anten_Shahrivar_1399=EPG_Shahrivar_1399.iat[38, 10]
active_user_tva_Shahrivar_1399=EPG_Shahrivar_1399.iat[39, 10]
active_user_fam_Shahrivar_1399=EPG_Shahrivar_1399.iat[40, 10]
active_user_televebion_Shahrivar_1399=EPG_Shahrivar_1399.iat[41, 10]
active_user_sepehr_Shahrivar_1399=EPG_Shahrivar_1399.iat[42, 10]
active_user_shima_Shahrivar_1399=EPG_Shahrivar_1399.iat[43, 10]
active_user_site_Shahrivar_1399=EPG_Shahrivar_1399.iat[44, 10]

all_visit_Shahrivar_1399=EPG_Shahrivar_1399.iat[25, 4]
all_duration_Shahrivar_1399=EPG_Shahrivar_1399.iat[25, 6]
all_content_sima_Shahrivar_1399=EPG_Shahrivar_1399.iat[25, 2]
all_register_user_Shahrivar_1399=sum(EPG_Shahrivar_1399.iloc[36:44, 4])
all_active_user_Shahrivar_1399=sum(EPG_Shahrivar_1399.iloc[36:44, 10])

Shahrivar_1399_sima_visit_channels=pd.DataFrame()
Shahrivar_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Shahrivar_1399, sima_2_visit_Shahrivar_1399, sima_3_visit_Shahrivar_1399,
                 sima_4_visit_Shahrivar_1399, sima_5_visit_Shahrivar_1399, sima_khabar_visit_Shahrivar_1399,
                 sima_ofogh_visit_Shahrivar_1399, sima_pooya_visit_Shahrivar_1399, sima_omid_visit_Shahrivar_1399,
                 sima_ifilm_visit_Shahrivar_1399, sima_namayesh_visit_Shahrivar_1399, sima_tamasha_visit_Shahrivar_1399,
                 sima_mostanad_visit_Shahrivar_1399, sima_shoma_visit_Shahrivar_1399, sima_amozesh_visit_Shahrivar_1399,
                 sima_varzesh_visit_Shahrivar_1399, sima_nasim_visit_Shahrivar_1399, sima_qoran_visit_Shahrivar_1399,
                 sima_salamat_visit_Shahrivar_1399, sima_irankala_visit_Shahrivar_1399, sima_alalam_visit_Shahrivar_1399,
                 sima_alkosar_visit_Shahrivar_1399, sima_presstv_visit_Shahrivar_1399, sima_sepehr_visit_Shahrivar_1399,],
        'duration': [sima_1_duration_Shahrivar_1399, sima_2_duration_Shahrivar_1399, sima_3_duration_Shahrivar_1399,
                 sima_4_duration_Shahrivar_1399, sima_5_duration_Shahrivar_1399, sima_khabar_duration_Shahrivar_1399,
                 sima_ofogh_duration_Shahrivar_1399, sima_pooya_duration_Shahrivar_1399, sima_omid_duration_Shahrivar_1399,
                 sima_ifilm_duration_Shahrivar_1399, sima_namayesh_duration_Shahrivar_1399, sima_tamasha_duration_Shahrivar_1399,
                 sima_mostanad_duration_Shahrivar_1399, sima_shoma_duration_Shahrivar_1399, sima_amozesh_duration_Shahrivar_1399,
                 sima_varzesh_duration_Shahrivar_1399, sima_nasim_duration_Shahrivar_1399, sima_qoran_duration_Shahrivar_1399,
                 sima_salamat_duration_Shahrivar_1399, sima_irankala_duration_Shahrivar_1399, sima_alalam_duration_Shahrivar_1399,
                 sima_alkosar_duration_Shahrivar_1399, sima_presstv_duration_Shahrivar_1399, sima_sepehr_duration_Shahrivar_1399,],}
Shahrivar_1399_sima_visit_channels=pd.DataFrame(Shahrivar_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Shahrivar_1399_sima_visit_channels=Shahrivar_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Shahrivar_1399_operator_data=pd.DataFrame()
Shahrivar_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Shahrivar_1399, sima_aio_visit_Shahrivar_1399, sima_anten_visit_Shahrivar_1399,
                 sima_tva_visit_Shahrivar_1399, sima_fam_visit_Shahrivar_1399, sima_televebion_visit_Shahrivar_1399,
                 sima_sepehr_visit_Shahrivar_1399, sima_shima_visit_Shahrivar_1399, sima_site_visit_Shahrivar_1399,],
       'register': [register_user_lenz_Shahrivar_1399, register_user_aio_Shahrivar_1399, register_user_anten_Shahrivar_1399,
                 register_user_tva_Shahrivar_1399, register_user_fam_Shahrivar_1399, register_user_televebion_Shahrivar_1399,
                 register_user_sepehr_Shahrivar_1399, register_user_shima_Shahrivar_1399, register_user_site_Shahrivar_1399,],
       'active': [active_user_lenz_Shahrivar_1399, active_user_aio_Shahrivar_1399, active_user_anten_Shahrivar_1399,
                 active_user_tva_Shahrivar_1399, active_user_fam_Shahrivar_1399, active_user_televebion_Shahrivar_1399,
                 active_user_sepehr_Shahrivar_1399, active_user_shima_Shahrivar_1399, active_user_site_Shahrivar_1399,],}

Shahrivar_1399_operator_data=pd.DataFrame(Shahrivar_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Shahrivar_1399_operator_data=Shahrivar_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Shahrivar_1399_all_data_summary=pd.DataFrame()
Shahrivar_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Shahrivar_1399, all_duration_Shahrivar_1399,all_content_sima_Shahrivar_1399,
                      all_register_user_Shahrivar_1399, all_active_user_Shahrivar_1399,],}

Shahrivar_1399_all_data_summary=pd.DataFrame(Shahrivar_1399_all_data_summary, columns=['parameters', 'statistics'])

Shahrivar_1399_all_data_summary=Shahrivar_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه شهریور 1399.xlsx', engine='xlsxwriter')
Shahrivar_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Shahrivar_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Shahrivar_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه شهریور')
writer.save()

        ########################### مهر #############################
print("EPG Mehr 1399")
EPG_Mehr_1399=pd.read_excel('EPG/EPG 1399/EPG Mehr 1399.xlsx', sheet_name='آمار')
EPG_Mehr_1399.fillna(0, inplace=True)
sima_1_visit_Mehr_1399=EPG_Mehr_1399.iat[1, 4]
sima_2_visit_Mehr_1399=EPG_Mehr_1399.iat[2, 4]
sima_3_visit_Mehr_1399=EPG_Mehr_1399.iat[3, 4]
sima_4_visit_Mehr_1399=EPG_Mehr_1399.iat[4, 4]
sima_5_visit_Mehr_1399=EPG_Mehr_1399.iat[5, 4]
sima_khabar_visit_Mehr_1399=EPG_Mehr_1399.iat[6, 4]
sima_ofogh_visit_Mehr_1399=EPG_Mehr_1399.iat[7, 4]
sima_pooya_visit_Mehr_1399=EPG_Mehr_1399.iat[8, 4]
sima_omid_visit_Mehr_1399=EPG_Mehr_1399.iat[9, 4]
sima_ifilm_visit_Mehr_1399=EPG_Mehr_1399.iat[10, 4]
sima_namayesh_visit_Mehr_1399=EPG_Mehr_1399.iat[11, 4]
sima_tamasha_visit_Mehr_1399=EPG_Mehr_1399.iat[12, 4]
sima_mostanad_visit_Mehr_1399=EPG_Mehr_1399.iat[13, 4]
sima_shoma_visit_Mehr_1399=EPG_Mehr_1399.iat[14, 4]
sima_amozesh_visit_Mehr_1399=EPG_Mehr_1399.iat[15, 4]
sima_varzesh_visit_Mehr_1399=EPG_Mehr_1399.iat[16, 4]
sima_nasim_visit_Mehr_1399=EPG_Mehr_1399.iat[17, 4]
sima_qoran_visit_Mehr_1399=EPG_Mehr_1399.iat[18, 4]
sima_salamat_visit_Mehr_1399=EPG_Mehr_1399.iat[19, 4]
sima_irankala_visit_Mehr_1399=EPG_Mehr_1399.iat[20, 4]
sima_alalam_visit_Mehr_1399=EPG_Mehr_1399.iat[21, 4]
sima_alkosar_visit_Mehr_1399=EPG_Mehr_1399.iat[22, 4]
sima_presstv_visit_Mehr_1399=EPG_Mehr_1399.iat[23, 4]
sima_sepehr_visit_Mehr_1399=EPG_Mehr_1399.iat[24, 4]

sima_1_duration_Mehr_1399=EPG_Mehr_1399.iat[1, 6]
sima_2_duration_Mehr_1399=EPG_Mehr_1399.iat[2, 6]
sima_3_duration_Mehr_1399=EPG_Mehr_1399.iat[3, 6]
sima_4_duration_Mehr_1399=EPG_Mehr_1399.iat[4, 6]
sima_5_duration_Mehr_1399=EPG_Mehr_1399.iat[5, 6]
sima_khabar_duration_Mehr_1399=EPG_Mehr_1399.iat[6, 6]
sima_ofogh_duration_Mehr_1399=EPG_Mehr_1399.iat[7, 6]
sima_pooya_duration_Mehr_1399=EPG_Mehr_1399.iat[8, 6]
sima_omid_duration_Mehr_1399=EPG_Mehr_1399.iat[9, 6]
sima_ifilm_duration_Mehr_1399=EPG_Mehr_1399.iat[10, 6]
sima_namayesh_duration_Mehr_1399=EPG_Mehr_1399.iat[11, 6]
sima_tamasha_duration_Mehr_1399=EPG_Mehr_1399.iat[12, 6]
sima_mostanad_duration_Mehr_1399=EPG_Mehr_1399.iat[13, 6]
sima_shoma_duration_Mehr_1399=EPG_Mehr_1399.iat[14, 6]
sima_amozesh_duration_Mehr_1399=EPG_Mehr_1399.iat[15, 6]
sima_varzesh_duration_Mehr_1399=EPG_Mehr_1399.iat[16, 6]
sima_nasim_duration_Mehr_1399=EPG_Mehr_1399.iat[17, 6]
sima_qoran_duration_Mehr_1399=EPG_Mehr_1399.iat[18, 6]
sima_salamat_duration_Mehr_1399=EPG_Mehr_1399.iat[19, 6]
sima_irankala_duration_Mehr_1399=EPG_Mehr_1399.iat[20, 6]
sima_alalam_duration_Mehr_1399=EPG_Mehr_1399.iat[21, 6]
sima_alkosar_duration_Mehr_1399=EPG_Mehr_1399.iat[22, 6]
sima_presstv_duration_Mehr_1399=EPG_Mehr_1399.iat[23, 6]
sima_sepehr_duration_Mehr_1399=EPG_Mehr_1399.iat[24, 6]

sima_lenz_visit_Mehr_1399=EPG_Mehr_1399.iat[33, 2]
sima_aio_visit_Mehr_1399=EPG_Mehr_1399.iat[34, 2]
sima_anten_visit_Mehr_1399=EPG_Mehr_1399.iat[35, 2]
sima_tva_visit_Mehr_1399=EPG_Mehr_1399.iat[36, 2]
sima_fam_visit_Mehr_1399=EPG_Mehr_1399.iat[37, 2]
sima_televebion_visit_Mehr_1399=EPG_Mehr_1399.iat[38, 2]
sima_sepehr_Mehr_1399=EPG_Mehr_1399.iat[39, 2]
sima_shima_visit_Mehr_1399=EPG_Mehr_1399.iat[40, 2]
sima_site_visit_Mehr_1399=EPG_Mehr_1399.iat[41, 2]

register_user_lenz_Mehr_1399=EPG_Mehr_1399.iat[36, 4]
register_user_aio_Mehr_1399=EPG_Mehr_1399.iat[37, 4]
register_user_anten_Mehr_1399=EPG_Mehr_1399.iat[38, 4]
register_user_tva_Mehr_1399=EPG_Mehr_1399.iat[39, 4]
register_user_fam_Mehr_1399=EPG_Mehr_1399.iat[40, 4]
register_user_televebion_Mehr_1399=EPG_Mehr_1399.iat[41, 4]
register_user_sepehr_Mehr_1399=EPG_Mehr_1399.iat[42, 4]
register_user_shima_Mehr_1399=EPG_Mehr_1399.iat[43, 4]
register_user_site_Mehr_1399=EPG_Mehr_1399.iat[44, 4]

active_user_lenz_Mehr_1399=EPG_Mehr_1399.iat[36, 10]
active_user_aio_Mehr_1399=EPG_Mehr_1399.iat[37, 10]
active_user_anten_Mehr_1399=EPG_Mehr_1399.iat[38, 10]
active_user_tva_Mehr_1399=EPG_Mehr_1399.iat[39, 10]
active_user_fam_Mehr_1399=EPG_Mehr_1399.iat[40, 10]
active_user_televebion_Mehr_1399=EPG_Mehr_1399.iat[41, 10]
active_user_sepehr_Mehr_1399=EPG_Mehr_1399.iat[42, 10]
active_user_shima_Mehr_1399=EPG_Mehr_1399.iat[43, 10]
active_user_site_Mehr_1399=EPG_Mehr_1399.iat[44, 10]

all_visit_Mehr_1399=EPG_Mehr_1399.iat[25, 4]
all_duration_Mehr_1399=EPG_Mehr_1399.iat[25, 6]
all_content_sima_Mehr_1399=EPG_Mehr_1399.iat[25, 2]
all_register_user_Mehr_1399=sum(EPG_Mehr_1399.iloc[36:44, 4])
all_active_user_Mehr_1399=sum(EPG_Mehr_1399.iloc[36:44, 10])

Mehr_1399_sima_visit_channels=pd.DataFrame()
Mehr_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Mehr_1399, sima_2_visit_Mehr_1399, sima_3_visit_Mehr_1399,
                 sima_4_visit_Mehr_1399, sima_5_visit_Mehr_1399, sima_khabar_visit_Mehr_1399,
                 sima_ofogh_visit_Mehr_1399, sima_pooya_visit_Mehr_1399, sima_omid_visit_Mehr_1399,
                 sima_ifilm_visit_Mehr_1399, sima_namayesh_visit_Mehr_1399, sima_tamasha_visit_Mehr_1399,
                 sima_mostanad_visit_Mehr_1399, sima_shoma_visit_Mehr_1399, sima_amozesh_visit_Mehr_1399,
                 sima_varzesh_visit_Mehr_1399, sima_nasim_visit_Mehr_1399, sima_qoran_visit_Mehr_1399,
                 sima_salamat_visit_Mehr_1399, sima_irankala_visit_Mehr_1399, sima_alalam_visit_Mehr_1399,
                 sima_alkosar_visit_Mehr_1399, sima_presstv_visit_Mehr_1399, sima_sepehr_visit_Mehr_1399,],
        'duration': [sima_1_duration_Mehr_1399, sima_2_duration_Mehr_1399, sima_3_duration_Mehr_1399,
                 sima_4_duration_Mehr_1399, sima_5_duration_Mehr_1399, sima_khabar_duration_Mehr_1399,
                 sima_ofogh_duration_Mehr_1399, sima_pooya_duration_Mehr_1399, sima_omid_duration_Mehr_1399,
                 sima_ifilm_duration_Mehr_1399, sima_namayesh_duration_Mehr_1399, sima_tamasha_duration_Mehr_1399,
                 sima_mostanad_duration_Mehr_1399, sima_shoma_duration_Mehr_1399, sima_amozesh_duration_Mehr_1399,
                 sima_varzesh_duration_Mehr_1399, sima_nasim_duration_Mehr_1399, sima_qoran_duration_Mehr_1399,
                 sima_salamat_duration_Mehr_1399, sima_irankala_duration_Mehr_1399, sima_alalam_duration_Mehr_1399,
                 sima_alkosar_duration_Mehr_1399, sima_presstv_duration_Mehr_1399, sima_sepehr_duration_Mehr_1399,],}
Mehr_1399_sima_visit_channels=pd.DataFrame(Mehr_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Mehr_1399_sima_visit_channels=Mehr_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Mehr_1399_operator_data=pd.DataFrame()
Mehr_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Mehr_1399, sima_aio_visit_Mehr_1399, sima_anten_visit_Mehr_1399,
                 sima_tva_visit_Mehr_1399, sima_fam_visit_Mehr_1399, sima_televebion_visit_Mehr_1399,
                 sima_sepehr_visit_Mehr_1399, sima_shima_visit_Mehr_1399, sima_site_visit_Mehr_1399,],
       'register': [register_user_lenz_Mehr_1399, register_user_aio_Mehr_1399, register_user_anten_Mehr_1399,
                 register_user_tva_Mehr_1399, register_user_fam_Mehr_1399, register_user_televebion_Mehr_1399,
                 register_user_sepehr_Mehr_1399, register_user_shima_Mehr_1399, register_user_site_Mehr_1399,],
       'active': [active_user_lenz_Mehr_1399, active_user_aio_Mehr_1399, active_user_anten_Mehr_1399,
                 active_user_tva_Mehr_1399, active_user_fam_Mehr_1399, active_user_televebion_Mehr_1399,
                 active_user_sepehr_Mehr_1399, active_user_shima_Mehr_1399, active_user_site_Mehr_1399,],}

Mehr_1399_operator_data=pd.DataFrame(Mehr_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Mehr_1399_operator_data=Mehr_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Mehr_1399_all_data_summary=pd.DataFrame()
Mehr_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Mehr_1399, all_duration_Mehr_1399,all_content_sima_Mehr_1399,
                      all_register_user_Mehr_1399, all_active_user_Mehr_1399,],}

Mehr_1399_all_data_summary=pd.DataFrame(Mehr_1399_all_data_summary, columns=['parameters', 'statistics'])

Mehr_1399_all_data_summary=Mehr_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه مهر 1399.xlsx', engine='xlsxwriter')
Mehr_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Mehr_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Mehr_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه مهر')
writer.save()

        ########################### آبان #############################
print("EPG Aban 1399")
EPG_Aban_1399=pd.read_excel('EPG/EPG 1399/EPG Aban 1399.xlsx', sheet_name='آمار')
EPG_Aban_1399.fillna(0, inplace=True)
sima_1_visit_Aban_1399=EPG_Aban_1399.iat[1, 4]
sima_2_visit_Aban_1399=EPG_Aban_1399.iat[2, 4]
sima_3_visit_Aban_1399=EPG_Aban_1399.iat[3, 4]
sima_4_visit_Aban_1399=EPG_Aban_1399.iat[4, 4]
sima_5_visit_Aban_1399=EPG_Aban_1399.iat[5, 4]
sima_khabar_visit_Aban_1399=EPG_Aban_1399.iat[6, 4]
sima_ofogh_visit_Aban_1399=EPG_Aban_1399.iat[7, 4]
sima_pooya_visit_Aban_1399=EPG_Aban_1399.iat[8, 4]
sima_omid_visit_Aban_1399=EPG_Aban_1399.iat[9, 4]
sima_ifilm_visit_Aban_1399=EPG_Aban_1399.iat[10, 4]
sima_namayesh_visit_Aban_1399=EPG_Aban_1399.iat[11, 4]
sima_tamasha_visit_Aban_1399=EPG_Aban_1399.iat[12, 4]
sima_mostanad_visit_Aban_1399=EPG_Aban_1399.iat[13, 4]
sima_shoma_visit_Aban_1399=EPG_Aban_1399.iat[14, 4]
sima_amozesh_visit_Aban_1399=EPG_Aban_1399.iat[15, 4]
sima_varzesh_visit_Aban_1399=EPG_Aban_1399.iat[16, 4]
sima_nasim_visit_Aban_1399=EPG_Aban_1399.iat[17, 4]
sima_qoran_visit_Aban_1399=EPG_Aban_1399.iat[18, 4]
sima_salamat_visit_Aban_1399=EPG_Aban_1399.iat[19, 4]
sima_irankala_visit_Aban_1399=EPG_Aban_1399.iat[20, 4]
sima_alalam_visit_Aban_1399=EPG_Aban_1399.iat[21, 4]
sima_alkosar_visit_Aban_1399=EPG_Aban_1399.iat[22, 4]
sima_presstv_visit_Aban_1399=EPG_Aban_1399.iat[23, 4]
sima_sepehr_visit_Aban_1399=EPG_Aban_1399.iat[24, 4]

sima_1_duration_Aban_1399=EPG_Aban_1399.iat[1, 6]
sima_2_duration_Aban_1399=EPG_Aban_1399.iat[2, 6]
sima_3_duration_Aban_1399=EPG_Aban_1399.iat[3, 6]
sima_4_duration_Aban_1399=EPG_Aban_1399.iat[4, 6]
sima_5_duration_Aban_1399=EPG_Aban_1399.iat[5, 6]
sima_khabar_duration_Aban_1399=EPG_Aban_1399.iat[6, 6]
sima_ofogh_duration_Aban_1399=EPG_Aban_1399.iat[7, 6]
sima_pooya_duration_Aban_1399=EPG_Aban_1399.iat[8, 6]
sima_omid_duration_Aban_1399=EPG_Aban_1399.iat[9, 6]
sima_ifilm_duration_Aban_1399=EPG_Aban_1399.iat[10, 6]
sima_namayesh_duration_Aban_1399=EPG_Aban_1399.iat[11, 6]
sima_tamasha_duration_Aban_1399=EPG_Aban_1399.iat[12, 6]
sima_mostanad_duration_Aban_1399=EPG_Aban_1399.iat[13, 6]
sima_shoma_duration_Aban_1399=EPG_Aban_1399.iat[14, 6]
sima_amozesh_duration_Aban_1399=EPG_Aban_1399.iat[15, 6]
sima_varzesh_duration_Aban_1399=EPG_Aban_1399.iat[16, 6]
sima_nasim_duration_Aban_1399=EPG_Aban_1399.iat[17, 6]
sima_qoran_duration_Aban_1399=EPG_Aban_1399.iat[18, 6]
sima_salamat_duration_Aban_1399=EPG_Aban_1399.iat[19, 6]
sima_irankala_duration_Aban_1399=EPG_Aban_1399.iat[20, 6]
sima_alalam_duration_Aban_1399=EPG_Aban_1399.iat[21, 6]
sima_alkosar_duration_Aban_1399=EPG_Aban_1399.iat[22, 6]
sima_presstv_duration_Aban_1399=EPG_Aban_1399.iat[23, 6]
sima_sepehr_duration_Aban_1399=EPG_Aban_1399.iat[24, 6]

sima_lenz_visit_Aban_1399=EPG_Aban_1399.iat[33, 2]
sima_aio_visit_Aban_1399=EPG_Aban_1399.iat[34, 2]
sima_anten_visit_Aban_1399=EPG_Aban_1399.iat[35, 2]
sima_tva_visit_Aban_1399=EPG_Aban_1399.iat[36, 2]
sima_fam_visit_Aban_1399=EPG_Aban_1399.iat[37, 2]
sima_televebion_visit_Aban_1399=EPG_Aban_1399.iat[38, 2]
sima_sepehr_Aban_1399=EPG_Aban_1399.iat[39, 2]
sima_shima_visit_Aban_1399=EPG_Aban_1399.iat[40, 2]
sima_site_visit_Aban_1399=EPG_Aban_1399.iat[41, 2]

register_user_lenz_Aban_1399=EPG_Aban_1399.iat[33, 4]
register_user_aio_Aban_1399=EPG_Aban_1399.iat[34, 4]
register_user_anten_Aban_1399=EPG_Aban_1399.iat[35, 4]
register_user_tva_Aban_1399=EPG_Aban_1399.iat[36, 4]
register_user_fam_Aban_1399=EPG_Aban_1399.iat[37, 4]
register_user_televebion_Aban_1399=EPG_Aban_1399.iat[38, 4]
register_user_sepehr_Aban_1399=EPG_Aban_1399.iat[39, 4]
register_user_shima_Aban_1399=EPG_Aban_1399.iat[40, 4]
register_user_site_Aban_1399=EPG_Aban_1399.iat[41, 4]

active_user_lenz_Aban_1399=EPG_Aban_1399.iat[33, 10]
active_user_aio_Aban_1399=EPG_Aban_1399.iat[34, 10]
active_user_anten_Aban_1399=EPG_Aban_1399.iat[35, 10]
active_user_tva_Aban_1399=EPG_Aban_1399.iat[36, 10]
active_user_fam_Aban_1399=EPG_Aban_1399.iat[37, 10]
active_user_televebion_Aban_1399=EPG_Aban_1399.iat[38, 10]
active_user_sepehr_Aban_1399=EPG_Aban_1399.iat[39, 10]
active_user_shima_Aban_1399=EPG_Aban_1399.iat[40, 10]
active_user_site_Aban_1399=EPG_Aban_1399.iat[41, 10]

all_visit_Aban_1399=EPG_Aban_1399.iat[25, 4]
all_duration_Aban_1399=EPG_Aban_1399.iat[25, 6]
all_content_sima_Aban_1399=EPG_Aban_1399.iat[25, 2]
all_register_user_Aban_1399=sum(EPG_Aban_1399.iloc[33:43, 4])
all_active_user_Aban_1399=sum(EPG_Aban_1399.iloc[33:43, 10])

Aban_1399_sima_visit_channels=pd.DataFrame()
Aban_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Aban_1399, sima_2_visit_Aban_1399, sima_3_visit_Aban_1399,
                 sima_4_visit_Aban_1399, sima_5_visit_Aban_1399, sima_khabar_visit_Aban_1399,
                 sima_ofogh_visit_Aban_1399, sima_pooya_visit_Aban_1399, sima_omid_visit_Aban_1399,
                 sima_ifilm_visit_Aban_1399, sima_namayesh_visit_Aban_1399, sima_tamasha_visit_Aban_1399,
                 sima_mostanad_visit_Aban_1399, sima_shoma_visit_Aban_1399, sima_amozesh_visit_Aban_1399,
                 sima_varzesh_visit_Aban_1399, sima_nasim_visit_Aban_1399, sima_qoran_visit_Aban_1399,
                 sima_salamat_visit_Aban_1399, sima_irankala_visit_Aban_1399, sima_alalam_visit_Aban_1399,
                 sima_alkosar_visit_Aban_1399, sima_presstv_visit_Aban_1399, sima_sepehr_visit_Aban_1399,],
        'duration': [sima_1_duration_Aban_1399, sima_2_duration_Aban_1399, sima_3_duration_Aban_1399,
                 sima_4_duration_Aban_1399, sima_5_duration_Aban_1399, sima_khabar_duration_Aban_1399,
                 sima_ofogh_duration_Aban_1399, sima_pooya_duration_Aban_1399, sima_omid_duration_Aban_1399,
                 sima_ifilm_duration_Aban_1399, sima_namayesh_duration_Aban_1399, sima_tamasha_duration_Aban_1399,
                 sima_mostanad_duration_Aban_1399, sima_shoma_duration_Aban_1399, sima_amozesh_duration_Aban_1399,
                 sima_varzesh_duration_Aban_1399, sima_nasim_duration_Aban_1399, sima_qoran_duration_Aban_1399,
                 sima_salamat_duration_Aban_1399, sima_irankala_duration_Aban_1399, sima_alalam_duration_Aban_1399,
                 sima_alkosar_duration_Aban_1399, sima_presstv_duration_Aban_1399, sima_sepehr_duration_Aban_1399,],}
Aban_1399_sima_visit_channels=pd.DataFrame(Aban_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Aban_1399_sima_visit_channels=Aban_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Aban_1399_operator_data=pd.DataFrame()
Aban_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Aban_1399, sima_aio_visit_Aban_1399, sima_anten_visit_Aban_1399,
                 sima_tva_visit_Aban_1399, sima_fam_visit_Aban_1399, sima_televebion_visit_Aban_1399,
                 sima_sepehr_visit_Aban_1399, sima_shima_visit_Aban_1399, sima_site_visit_Aban_1399,],
       'register': [register_user_lenz_Aban_1399, register_user_aio_Aban_1399, register_user_anten_Aban_1399,
                 register_user_tva_Aban_1399, register_user_fam_Aban_1399, register_user_televebion_Aban_1399,
                 register_user_sepehr_Aban_1399, register_user_shima_Aban_1399, register_user_site_Aban_1399,],
       'active': [active_user_lenz_Aban_1399, active_user_aio_Aban_1399, active_user_anten_Aban_1399,
                 active_user_tva_Aban_1399, active_user_fam_Aban_1399, active_user_televebion_Aban_1399,
                 active_user_sepehr_Aban_1399, active_user_shima_Aban_1399, active_user_site_Aban_1399,],}

Aban_1399_operator_data=pd.DataFrame(Aban_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Aban_1399_operator_data=Aban_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Aban_1399_all_data_summary=pd.DataFrame()
Aban_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Aban_1399, all_duration_Aban_1399,all_content_sima_Aban_1399,
                      all_register_user_Aban_1399, all_active_user_Aban_1399,],}

Aban_1399_all_data_summary=pd.DataFrame(Aban_1399_all_data_summary, columns=['parameters', 'statistics'])

Aban_1399_all_data_summary=Aban_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه آبان 1399.xlsx', engine='xlsxwriter')
Aban_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Aban_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Aban_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه آبان')
writer.save()

        ########################### آذر #############################
print("EPG Azar 1399")
EPG_Azar_1399=pd.read_excel('EPG/EPG 1399/EPG Azar 1399.xlsx', sheet_name='آمار')
EPG_Azar_1399.fillna(0, inplace=True)
sima_1_visit_Azar_1399=EPG_Azar_1399.iat[1, 4]
sima_2_visit_Azar_1399=EPG_Azar_1399.iat[2, 4]
sima_3_visit_Azar_1399=EPG_Azar_1399.iat[3, 4]
sima_4_visit_Azar_1399=EPG_Azar_1399.iat[4, 4]
sima_5_visit_Azar_1399=EPG_Azar_1399.iat[5, 4]
sima_khabar_visit_Azar_1399=EPG_Azar_1399.iat[6, 4]
sima_ofogh_visit_Azar_1399=EPG_Azar_1399.iat[7, 4]
sima_pooya_visit_Azar_1399=EPG_Azar_1399.iat[8, 4]
sima_omid_visit_Azar_1399=EPG_Azar_1399.iat[9, 4]
sima_ifilm_visit_Azar_1399=EPG_Azar_1399.iat[10, 4]
sima_namayesh_visit_Azar_1399=EPG_Azar_1399.iat[11, 4]
sima_tamasha_visit_Azar_1399=EPG_Azar_1399.iat[12, 4]
sima_mostanad_visit_Azar_1399=EPG_Azar_1399.iat[13, 4]
sima_shoma_visit_Azar_1399=EPG_Azar_1399.iat[14, 4]
sima_amozesh_visit_Azar_1399=EPG_Azar_1399.iat[15, 4]
sima_varzesh_visit_Azar_1399=EPG_Azar_1399.iat[16, 4]
sima_nasim_visit_Azar_1399=EPG_Azar_1399.iat[17, 4]
sima_qoran_visit_Azar_1399=EPG_Azar_1399.iat[18, 4]
sima_salamat_visit_Azar_1399=EPG_Azar_1399.iat[19, 4]
sima_irankala_visit_Azar_1399=EPG_Azar_1399.iat[20, 4]
sima_alalam_visit_Azar_1399=EPG_Azar_1399.iat[21, 4]
sima_alkosar_visit_Azar_1399=EPG_Azar_1399.iat[22, 4]
sima_presstv_visit_Azar_1399=EPG_Azar_1399.iat[23, 4]
sima_sepehr_visit_Azar_1399=EPG_Azar_1399.iat[24, 4]

sima_1_duration_Azar_1399=EPG_Azar_1399.iat[1, 6]
sima_2_duration_Azar_1399=EPG_Azar_1399.iat[2, 6]
sima_3_duration_Azar_1399=EPG_Azar_1399.iat[3, 6]
sima_4_duration_Azar_1399=EPG_Azar_1399.iat[4, 6]
sima_5_duration_Azar_1399=EPG_Azar_1399.iat[5, 6]
sima_khabar_duration_Azar_1399=EPG_Azar_1399.iat[6, 6]
sima_ofogh_duration_Azar_1399=EPG_Azar_1399.iat[7, 6]
sima_pooya_duration_Azar_1399=EPG_Azar_1399.iat[8, 6]
sima_omid_duration_Azar_1399=EPG_Azar_1399.iat[9, 6]
sima_ifilm_duration_Azar_1399=EPG_Azar_1399.iat[10, 6]
sima_namayesh_duration_Azar_1399=EPG_Azar_1399.iat[11, 6]
sima_tamasha_duration_Azar_1399=EPG_Azar_1399.iat[12, 6]
sima_mostanad_duration_Azar_1399=EPG_Azar_1399.iat[13, 6]
sima_shoma_duration_Azar_1399=EPG_Azar_1399.iat[14, 6]
sima_amozesh_duration_Azar_1399=EPG_Azar_1399.iat[15, 6]
sima_varzesh_duration_Azar_1399=EPG_Azar_1399.iat[16, 6]
sima_nasim_duration_Azar_1399=EPG_Azar_1399.iat[17, 6]
sima_qoran_duration_Azar_1399=EPG_Azar_1399.iat[18, 6]
sima_salamat_duration_Azar_1399=EPG_Azar_1399.iat[19, 6]
sima_irankala_duration_Azar_1399=EPG_Azar_1399.iat[20, 6]
sima_alalam_duration_Azar_1399=EPG_Azar_1399.iat[21, 6]
sima_alkosar_duration_Azar_1399=EPG_Azar_1399.iat[22, 6]
sima_presstv_duration_Azar_1399=EPG_Azar_1399.iat[23, 6]
sima_sepehr_duration_Azar_1399=EPG_Azar_1399.iat[24, 6]

sima_lenz_visit_Azar_1399=EPG_Azar_1399.iat[33, 2]
sima_aio_visit_Azar_1399=EPG_Azar_1399.iat[34, 2]
sima_anten_visit_Azar_1399=EPG_Azar_1399.iat[35, 2]
sima_tva_visit_Azar_1399=EPG_Azar_1399.iat[36, 2]
sima_fam_visit_Azar_1399=EPG_Azar_1399.iat[37, 2]
sima_televebion_visit_Azar_1399=EPG_Azar_1399.iat[38, 2]
sima_sepehr_Azar_1399=EPG_Azar_1399.iat[39, 2]
sima_shima_visit_Azar_1399=EPG_Azar_1399.iat[40, 2]
sima_site_visit_Azar_1399=EPG_Azar_1399.iat[41, 2]

register_user_lenz_azar_1399=EPG_Azar_1399.iat[33, 4]
register_user_aio_azar_1399=EPG_Azar_1399.iat[34, 4]
register_user_anten_azar_1399=EPG_Azar_1399.iat[35, 4]
register_user_tva_azar_1399=EPG_Azar_1399.iat[36, 4]
register_user_fam_azar_1399=EPG_Azar_1399.iat[37, 4]
register_user_televebion_azar_1399=EPG_Azar_1399.iat[38, 4]
register_user_sepehr_azar_1399=EPG_Azar_1399.iat[39, 4]
register_user_shima_azar_1399=EPG_Azar_1399.iat[40, 4]
register_user_site_azar_1399=EPG_Azar_1399.iat[41, 4]

active_user_lenz_azar_1399=EPG_Azar_1399.iat[33, 10]
active_user_aio_azar_1399=EPG_Azar_1399.iat[34, 10]
active_user_anten_azar_1399=EPG_Azar_1399.iat[35, 10]
active_user_tva_azar_1399=EPG_Azar_1399.iat[36, 10]
active_user_fam_azar_1399=EPG_Azar_1399.iat[37, 10]
active_user_televebion_azar_1399=EPG_Azar_1399.iat[38, 10]
active_user_sepehr_azar_1399=EPG_Azar_1399.iat[39, 10]
active_user_shima_azar_1399=EPG_Azar_1399.iat[40, 10]
active_user_site_azar_1399=EPG_Azar_1399.iat[41, 10]

all_visit_azar_1399=EPG_Azar_1399.iat[25, 4]
all_duration_azar_1399=EPG_Azar_1399.iat[25, 6]
all_content_sima_azar_1399=EPG_Azar_1399.iat[25, 2]
all_register_user_azar_1399=sum(EPG_Azar_1399.iloc[33:43, 4])
all_active_user_azar_1399=sum(EPG_Azar_1399.iloc[33:43, 10])

azar_1399_sima_visit_channels=pd.DataFrame()
azar_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Azar_1399, sima_2_visit_Azar_1399, sima_3_visit_Azar_1399,
                 sima_4_visit_Azar_1399, sima_5_visit_Azar_1399, sima_khabar_visit_Azar_1399,
                 sima_ofogh_visit_Azar_1399, sima_pooya_visit_Azar_1399, sima_omid_visit_Azar_1399,
                 sima_ifilm_visit_Azar_1399, sima_namayesh_visit_Azar_1399, sima_tamasha_visit_Azar_1399,
                 sima_mostanad_visit_Azar_1399, sima_shoma_visit_Azar_1399, sima_amozesh_visit_Azar_1399,
                 sima_varzesh_visit_Azar_1399, sima_nasim_visit_Azar_1399, sima_qoran_visit_Azar_1399,
                 sima_salamat_visit_Azar_1399, sima_irankala_visit_Azar_1399, sima_alalam_visit_Azar_1399,
                 sima_alkosar_visit_Azar_1399, sima_presstv_visit_Azar_1399, sima_sepehr_visit_Azar_1399,],
        'duration': [sima_1_duration_Azar_1399, sima_2_duration_Azar_1399, sima_3_duration_Azar_1399,
                 sima_4_duration_Azar_1399, sima_5_duration_Azar_1399, sima_khabar_duration_Azar_1399,
                 sima_ofogh_duration_Azar_1399, sima_pooya_duration_Azar_1399, sima_omid_duration_Azar_1399,
                 sima_ifilm_duration_Azar_1399, sima_namayesh_duration_Azar_1399, sima_tamasha_duration_Azar_1399,
                 sima_mostanad_duration_Azar_1399, sima_shoma_duration_Azar_1399, sima_amozesh_duration_Azar_1399,
                 sima_varzesh_duration_Azar_1399, sima_nasim_duration_Azar_1399, sima_qoran_duration_Azar_1399,
                 sima_salamat_duration_Azar_1399, sima_irankala_duration_Azar_1399, sima_alalam_duration_Azar_1399,
                 sima_alkosar_duration_Azar_1399, sima_presstv_duration_Azar_1399, sima_sepehr_duration_Azar_1399,],}
azar_1399_sima_visit_channels=pd.DataFrame(azar_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

azar_1399_sima_visit_channels=azar_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

azar_1399_operator_data=pd.DataFrame()
azar_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Azar_1399, sima_aio_visit_Azar_1399, sima_anten_visit_Azar_1399,
                 sima_tva_visit_Azar_1399, sima_fam_visit_Azar_1399, sima_televebion_visit_Azar_1399,
                 sima_sepehr_visit_Azar_1399, sima_shima_visit_Azar_1399, sima_site_visit_Azar_1399,],
       'register': [register_user_lenz_azar_1399, register_user_aio_azar_1399, register_user_anten_azar_1399,
                 register_user_tva_azar_1399, register_user_fam_azar_1399, register_user_televebion_azar_1399,
                 register_user_sepehr_azar_1399, register_user_shima_azar_1399, register_user_site_azar_1399,],
       'active': [active_user_lenz_azar_1399, active_user_aio_azar_1399, active_user_anten_azar_1399,
                 active_user_tva_azar_1399, active_user_fam_azar_1399, active_user_televebion_azar_1399,
                 active_user_sepehr_azar_1399, active_user_shima_azar_1399, active_user_site_azar_1399,],}

azar_1399_operator_data=pd.DataFrame(azar_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

azar_1399_operator_data=azar_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

azar_1399_all_data_summary=pd.DataFrame()
azar_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_azar_1399, all_duration_azar_1399,all_content_sima_azar_1399,
                      all_register_user_azar_1399, all_active_user_azar_1399,],}

azar_1399_all_data_summary=pd.DataFrame(azar_1399_all_data_summary, columns=['parameters', 'statistics'])

azar_1399_all_data_summary=azar_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه آذر 1399.xlsx', engine='xlsxwriter')
azar_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
azar_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
azar_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه آذر')
writer.save()

        ########################### دی #############################
print("EPG Dey 1399")
EPG_Dey_1399=pd.read_excel('EPG/EPG 1399/EPG Dey 1399.xlsx', sheet_name='آمار')
EPG_Dey_1399.fillna(0, inplace=True)
sima_1_visit_Dey_1399=EPG_Dey_1399.iat[1, 4]
sima_2_visit_Dey_1399=EPG_Dey_1399.iat[2, 4]
sima_3_visit_Dey_1399=EPG_Dey_1399.iat[3, 4]
sima_4_visit_Dey_1399=EPG_Dey_1399.iat[4, 4]
sima_5_visit_Dey_1399=EPG_Dey_1399.iat[5, 4]
sima_khabar_visit_Dey_1399=EPG_Dey_1399.iat[6, 4]
sima_ofogh_visit_Dey_1399=EPG_Dey_1399.iat[7, 4]
sima_pooya_visit_Dey_1399=EPG_Dey_1399.iat[8, 4]
sima_omid_visit_Dey_1399=EPG_Dey_1399.iat[9, 4]
sima_ifilm_visit_Dey_1399=EPG_Dey_1399.iat[10, 4]
sima_namayesh_visit_Dey_1399=EPG_Dey_1399.iat[11, 4]
sima_tamasha_visit_Dey_1399=EPG_Dey_1399.iat[12, 4]
sima_mostanad_visit_Dey_1399=EPG_Dey_1399.iat[13, 4]
sima_shoma_visit_Dey_1399=EPG_Dey_1399.iat[14, 4]
sima_amozesh_visit_Dey_1399=EPG_Dey_1399.iat[15, 4]
sima_varzesh_visit_Dey_1399=EPG_Dey_1399.iat[16, 4]
sima_nasim_visit_Dey_1399=EPG_Dey_1399.iat[17, 4]
sima_qoran_visit_Dey_1399=EPG_Dey_1399.iat[18, 4]
sima_salamat_visit_Dey_1399=EPG_Dey_1399.iat[19, 4]
sima_irankala_visit_Dey_1399=EPG_Dey_1399.iat[20, 4]
sima_alalam_visit_Dey_1399=EPG_Dey_1399.iat[21, 4]
sima_alkosar_visit_Dey_1399=EPG_Dey_1399.iat[22, 4]
sima_presstv_visit_Dey_1399=EPG_Dey_1399.iat[23, 4]
sima_sepehr_visit_Dey_1399=EPG_Dey_1399.iat[24, 4]

sima_1_duration_Dey_1399=EPG_Dey_1399.iat[1, 6]
sima_2_duration_Dey_1399=EPG_Dey_1399.iat[2, 6]
sima_3_duration_Dey_1399=EPG_Dey_1399.iat[3, 6]
sima_4_duration_Dey_1399=EPG_Dey_1399.iat[4, 6]
sima_5_duration_Dey_1399=EPG_Dey_1399.iat[5, 6]
sima_khabar_duration_Dey_1399=EPG_Dey_1399.iat[6, 6]
sima_ofogh_duration_Dey_1399=EPG_Dey_1399.iat[7, 6]
sima_pooya_duration_Dey_1399=EPG_Dey_1399.iat[8, 6]
sima_omid_duration_Dey_1399=EPG_Dey_1399.iat[9, 6]
sima_ifilm_duration_Dey_1399=EPG_Dey_1399.iat[10, 6]
sima_namayesh_duration_Dey_1399=EPG_Dey_1399.iat[11, 6]
sima_tamasha_duration_Dey_1399=EPG_Dey_1399.iat[12, 6]
sima_mostanad_duration_Dey_1399=EPG_Dey_1399.iat[13, 6]
sima_shoma_duration_Dey_1399=EPG_Dey_1399.iat[14, 6]
sima_amozesh_duration_Dey_1399=EPG_Dey_1399.iat[15, 6]
sima_varzesh_duration_Dey_1399=EPG_Dey_1399.iat[16, 6]
sima_nasim_duration_Dey_1399=EPG_Dey_1399.iat[17, 6]
sima_qoran_duration_Dey_1399=EPG_Dey_1399.iat[18, 6]
sima_salamat_duration_Dey_1399=EPG_Dey_1399.iat[19, 6]
sima_irankala_duration_Dey_1399=EPG_Dey_1399.iat[20, 6]
sima_alalam_duration_Dey_1399=EPG_Dey_1399.iat[21, 6]
sima_alkosar_duration_Dey_1399=EPG_Dey_1399.iat[22, 6]
sima_presstv_duration_Dey_1399=EPG_Dey_1399.iat[23, 6]
sima_sepehr_duration_Dey_1399=EPG_Dey_1399.iat[24, 6]

sima_lenz_visit_Dey_1399=EPG_Dey_1399.iat[33, 2]
sima_aio_visit_Dey_1399=EPG_Dey_1399.iat[34, 2]
sima_anten_visit_Dey_1399=EPG_Dey_1399.iat[35, 2]
sima_tva_visit_Dey_1399=EPG_Dey_1399.iat[36, 2]
sima_fam_visit_Dey_1399=EPG_Dey_1399.iat[37, 2]
sima_televebion_visit_Dey_1399=EPG_Dey_1399.iat[38, 2]
sima_sepehr_Dey_1399=EPG_Dey_1399.iat[39, 2]
sima_shima_visit_Dey_1399=EPG_Dey_1399.iat[40, 2]
sima_site_visit_Dey_1399=EPG_Dey_1399.iat[41, 2]

register_user_lenz_Dey_1399=EPG_Dey_1399.iat[33, 4]
register_user_aio_Dey_1399=EPG_Dey_1399.iat[34, 4]
register_user_anten_Dey_1399=EPG_Dey_1399.iat[35, 4]
register_user_tva_Dey_1399=EPG_Dey_1399.iat[36, 4]
register_user_fam_Dey_1399=EPG_Dey_1399.iat[37, 4]
register_user_televebion_Dey_1399=EPG_Dey_1399.iat[38, 4]
register_user_sepehr_Dey_1399=EPG_Dey_1399.iat[39, 4]
register_user_shima_Dey_1399=EPG_Dey_1399.iat[40, 4]
register_user_site_Dey_1399=EPG_Dey_1399.iat[41, 4]

active_user_lenz_Dey_1399=EPG_Dey_1399.iat[33, 10]
active_user_aio_Dey_1399=EPG_Dey_1399.iat[34, 10]
active_user_anten_Dey_1399=EPG_Dey_1399.iat[35, 10]
active_user_tva_Dey_1399=EPG_Dey_1399.iat[36, 10]
active_user_fam_Dey_1399=EPG_Dey_1399.iat[37, 10]
active_user_televebion_Dey_1399=EPG_Dey_1399.iat[38, 10]
active_user_sepehr_Dey_1399=EPG_Dey_1399.iat[39, 10]
active_user_shima_Dey_1399=EPG_Dey_1399.iat[40, 10]
active_user_site_Dey_1399=EPG_Dey_1399.iat[41, 10]

all_visit_Dey_1399=EPG_Dey_1399.iat[25, 4]
all_duration_Dey_1399=EPG_Dey_1399.iat[25, 6]
all_content_sima_Dey_1399=EPG_Dey_1399.iat[25, 2]
all_register_user_Dey_1399=sum(EPG_Dey_1399.iloc[33:43, 4])
all_active_user_Dey_1399=sum(EPG_Dey_1399.iloc[33:43, 10])

Dey_1399_sima_visit_channels=pd.DataFrame()
Dey_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Dey_1399, sima_2_visit_Dey_1399, sima_3_visit_Dey_1399,
                 sima_4_visit_Dey_1399, sima_5_visit_Dey_1399, sima_khabar_visit_Dey_1399,
                 sima_ofogh_visit_Dey_1399, sima_pooya_visit_Dey_1399, sima_omid_visit_Dey_1399,
                 sima_ifilm_visit_Dey_1399, sima_namayesh_visit_Dey_1399, sima_tamasha_visit_Dey_1399,
                 sima_mostanad_visit_Dey_1399, sima_shoma_visit_Dey_1399, sima_amozesh_visit_Dey_1399,
                 sima_varzesh_visit_Dey_1399, sima_nasim_visit_Dey_1399, sima_qoran_visit_Dey_1399,
                 sima_salamat_visit_Dey_1399, sima_irankala_visit_Dey_1399, sima_alalam_visit_Dey_1399,
                 sima_alkosar_visit_Dey_1399, sima_presstv_visit_Dey_1399, sima_sepehr_visit_Dey_1399,],
        'duration': [sima_1_duration_Dey_1399, sima_2_duration_Dey_1399, sima_3_duration_Dey_1399,
                 sima_4_duration_Dey_1399, sima_5_duration_Dey_1399, sima_khabar_duration_Dey_1399,
                 sima_ofogh_duration_Dey_1399, sima_pooya_duration_Dey_1399, sima_omid_duration_Dey_1399,
                 sima_ifilm_duration_Dey_1399, sima_namayesh_duration_Dey_1399, sima_tamasha_duration_Dey_1399,
                 sima_mostanad_duration_Dey_1399, sima_shoma_duration_Dey_1399, sima_amozesh_duration_Dey_1399,
                 sima_varzesh_duration_Dey_1399, sima_nasim_duration_Dey_1399, sima_qoran_duration_Dey_1399,
                 sima_salamat_duration_Dey_1399, sima_irankala_duration_Dey_1399, sima_alalam_duration_Dey_1399,
                 sima_alkosar_duration_Dey_1399, sima_presstv_duration_Dey_1399, sima_sepehr_duration_Dey_1399,],}
Dey_1399_sima_visit_channels=pd.DataFrame(Dey_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Dey_1399_sima_visit_channels=Dey_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Dey_1399_operator_data=pd.DataFrame()
Dey_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
       'visit': [sima_lenz_visit_Dey_1399, sima_aio_visit_Dey_1399, sima_anten_visit_Dey_1399,
                 sima_tva_visit_Dey_1399, sima_fam_visit_Dey_1399, sima_televebion_visit_Dey_1399,
                 sima_sepehr_visit_Dey_1399, sima_shima_visit_Dey_1399, sima_site_visit_Dey_1399,],
       'register': [register_user_lenz_Dey_1399, register_user_aio_Dey_1399, register_user_anten_Dey_1399,
                 register_user_tva_Dey_1399, register_user_fam_Dey_1399, register_user_televebion_Dey_1399,
                 register_user_sepehr_Dey_1399, register_user_shima_Dey_1399, register_user_site_Dey_1399,],
       'active': [active_user_lenz_Dey_1399, active_user_aio_Dey_1399, active_user_anten_Dey_1399,
                 active_user_tva_Dey_1399, active_user_fam_Dey_1399, active_user_televebion_Dey_1399,
                 active_user_sepehr_Dey_1399, active_user_shima_Dey_1399, active_user_site_Dey_1399,],}

Dey_1399_operator_data=pd.DataFrame(Dey_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])

Dey_1399_operator_data=Dey_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Dey_1399_all_data_summary=pd.DataFrame()
Dey_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
       'statistics': [all_visit_Dey_1399, all_duration_Dey_1399,all_content_sima_Dey_1399,
                      all_register_user_Dey_1399, all_active_user_Dey_1399,],}

Dey_1399_all_data_summary=pd.DataFrame(Dey_1399_all_data_summary, columns=['parameters', 'statistics'])

Dey_1399_all_data_summary=Dey_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه دی 1399.xlsx', engine='xlsxwriter')
Dey_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Dey_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
Dey_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه دی')
writer.save()

 ########################### بهمن #############################
#print("EPG Bahman 1399")
#EPG_Bahman_1399=pd.read_excel('EPG/EPG 1399/EPG Bahman 1399.xlsx', sheet_name='آمار')
#EPG_Bahman_1399.fillna(0, inplace=True)
#sima_1_visit_Bahman_1399=EPG_Bahman_1399.iat[1, 4]
#sima_2_visit_Bahman_1399=EPG_Bahman_1399.iat[2, 4]
#sima_3_visit_Bahman_1399=EPG_Bahman_1399.iat[3, 4]
#sima_4_visit_Bahman_1399=EPG_Bahman_1399.iat[4, 4]
#sima_5_visit_Bahman_1399=EPG_Bahman_1399.iat[5, 4]
#sima_khabar_visit_Bahman_1399=EPG_Bahman_1399.iat[6, 4]
#sima_ofogh_visit_Bahman_1399=EPG_Bahman_1399.iat[7, 4]
#sima_pooya_visit_Bahman_1399=EPG_Bahman_1399.iat[8, 4]
#sima_omid_visit_Bahman_1399=EPG_Bahman_1399.iat[9, 4]
#sima_ifilm_visit_Bahman_1399=EPG_Bahman_1399.iat[10, 4]
#sima_namayesh_visit_Bahman_1399=EPG_Bahman_1399.iat[11, 4]
#sima_tamasha_visit_Bahman_1399=EPG_Bahman_1399.iat[12, 4]
#sima_mostanad_visit_Bahman_1399=EPG_Bahman_1399.iat[13, 4]
#sima_shoma_visit_Bahman_1399=EPG_Bahman_1399.iat[14, 4]
#sima_amozesh_visit_Bahman_1399=EPG_Bahman_1399.iat[15, 4]
#sima_varzesh_visit_Bahman_1399=EPG_Bahman_1399.iat[16, 4]
#sima_nasim_visit_Bahman_1399=EPG_Bahman_1399.iat[17, 4]
#sima_qoran_visit_Bahman_1399=EPG_Bahman_1399.iat[18, 4]
#sima_salamat_visit_Bahman_1399=EPG_Bahman_1399.iat[19, 4]
#sima_irankala_visit_Bahman_1399=EPG_Bahman_1399.iat[20, 4]
#sima_alalam_visit_Bahman_1399=EPG_Bahman_1399.iat[21, 4]
#sima_alkosar_visit_Bahman_1399=EPG_Bahman_1399.iat[22, 4]
#sima_presstv_visit_Bahman_1399=EPG_Bahman_1399.iat[23, 4]
#sima_sepehr_visit_Bahman_1399=EPG_Bahman_1399.iat[24, 4]
#
#sima_1_duration_Bahman_1399=EPG_Bahman_1399.iat[1, 6]
#sima_2_duration_Bahman_1399=EPG_Bahman_1399.iat[2, 6]
#sima_3_duration_Bahman_1399=EPG_Bahman_1399.iat[3, 6]
#sima_4_duration_Bahman_1399=EPG_Bahman_1399.iat[4, 6]
#sima_5_duration_Bahman_1399=EPG_Bahman_1399.iat[5, 6]
#sima_khabar_duration_Bahman_1399=EPG_Bahman_1399.iat[6, 6]
#sima_ofogh_duration_Bahman_1399=EPG_Bahman_1399.iat[7, 6]
#sima_pooya_duration_Bahman_1399=EPG_Bahman_1399.iat[8, 6]
#sima_omid_duration_Bahman_1399=EPG_Bahman_1399.iat[9, 6]
#sima_ifilm_duration_Bahman_1399=EPG_Bahman_1399.iat[10, 6]
#sima_namayesh_duration_Bahman_1399=EPG_Bahman_1399.iat[11, 6]
#sima_tamasha_duration_Bahman_1399=EPG_Bahman_1399.iat[12, 6]
#sima_mostanad_duration_Bahman_1399=EPG_Bahman_1399.iat[13, 6]
#sima_shoma_duration_Bahman_1399=EPG_Bahman_1399.iat[14, 6]
#sima_amozesh_duration_Bahman_1399=EPG_Bahman_1399.iat[15, 6]
#sima_varzesh_duration_Bahman_1399=EPG_Bahman_1399.iat[16, 6]
#sima_nasim_duration_Bahman_1399=EPG_Bahman_1399.iat[17, 6]
#sima_qoran_duration_Bahman_1399=EPG_Bahman_1399.iat[18, 6]
#sima_salamat_duration_Bahman_1399=EPG_Bahman_1399.iat[19, 6]
#sima_irankala_duration_Bahman_1399=EPG_Bahman_1399.iat[20, 6]
#sima_alalam_duration_Bahman_1399=EPG_Bahman_1399.iat[21, 6]
#sima_alkosar_duration_Bahman_1399=EPG_Bahman_1399.iat[22, 6]
#sima_presstv_duration_Bahman_1399=EPG_Bahman_1399.iat[23, 6]
#sima_sepehr_duration_Bahman_1399=EPG_Bahman_1399.iat[24, 6]
#
#sima_lenz_visit_Bahman_1399=EPG_Bahman_1399.iat[33, 2]
#sima_aio_visit_Bahman_1399=EPG_Bahman_1399.iat[34, 2]
#sima_anten_visit_Bahman_1399=EPG_Bahman_1399.iat[35, 2]
#sima_tva_visit_Bahman_1399=EPG_Bahman_1399.iat[36, 2]
#sima_fam_visit_Bahman_1399=EPG_Bahman_1399.iat[37, 2]
#sima_televebion_visit_Bahman_1399=EPG_Bahman_1399.iat[38, 2]
#sima_sepehr_Bahman_1399=EPG_Bahman_1399.iat[39, 2]
#sima_shima_visit_Bahman_1399=EPG_Bahman_1399.iat[40, 2]
#sima_site_visit_Bahman_1399=EPG_Bahman_1399.iat[41, 2]
#
#register_user_lenz_Bahman_1399=EPG_Bahman_1399.iat[33, 4]
#register_user_aio_Bahman_1399=EPG_Bahman_1399.iat[34, 4]
#register_user_anten_Bahman_1399=EPG_Bahman_1399.iat[35, 4]
#register_user_tva_Bahman_1399=EPG_Bahman_1399.iat[36, 4]
#register_user_fam_Bahman_1399=EPG_Bahman_1399.iat[37, 4]
#register_user_televebion_Bahman_1399=EPG_Bahman_1399.iat[38, 4]
#register_user_sepehr_Bahman_1399=EPG_Bahman_1399.iat[39, 4]
#register_user_shima_Bahman_1399=EPG_Bahman_1399.iat[40, 4]
#register_user_site_Bahman_1399=EPG_Bahman_1399.iat[41, 4]
#
#active_user_lenz_Bahman_1399=EPG_Bahman_1399.iat[33, 10]
#active_user_aio_Bahman_1399=EPG_Bahman_1399.iat[34, 10]
#active_user_anten_Bahman_1399=EPG_Bahman_1399.iat[35, 10]
#active_user_tva_Bahman_1399=EPG_Bahman_1399.iat[36, 10]
#active_user_fam_Bahman_1399=EPG_Bahman_1399.iat[37, 10]
#active_user_televebion_Bahman_1399=EPG_Bahman_1399.iat[38, 10]
#active_user_sepehr_Bahman_1399=EPG_Bahman_1399.iat[39, 10]
#active_user_shima_Bahman_1399=EPG_Bahman_1399.iat[40, 10]
#active_user_site_Bahman_1399=EPG_Bahman_1399.iat[41, 10]
#
#all_visit_Bahman_1399=EPG_Bahman_1399.iat[25, 4]
#all_duration_Bahman_1399=EPG_Bahman_1399.iat[25, 6]
#all_content_sima_Bahman_1399=EPG_Bahman_1399.iat[25, 2]
#all_register_user_Bahman_1399=sum(EPG_Bahman_1399.iloc[33:43, 4])
#all_active_user_Bahman_1399=sum(EPG_Bahman_1399.iloc[33:43, 10])
#
#Bahman_1399_sima_visit_channels=pd.DataFrame()
#Bahman_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
#                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
#                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
#                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
#                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
#       'visit': [sima_1_visit_Bahman_1399, sima_2_visit_Bahman_1399, sima_3_visit_Bahman_1399,
#                 sima_4_visit_Bahman_1399, sima_5_visit_Bahman_1399, sima_khabar_visit_Bahman_1399,
#                 sima_ofogh_visit_Bahman_1399, sima_pooya_visit_Bahman_1399, sima_omid_visit_Bahman_1399,
#                 sima_ifilm_visit_Bahman_1399, sima_namayesh_visit_Bahman_1399, sima_tamasha_visit_Bahman_1399,
#                 sima_mostanad_visit_Bahman_1399, sima_shoma_visit_Bahman_1399, sima_amozesh_visit_Bahman_1399,
#                 sima_varzesh_visit_Bahman_1399, sima_nasim_visit_Bahman_1399, sima_qoran_visit_Bahman_1399,
#                 sima_salamat_visit_Bahman_1399, sima_irankala_visit_Bahman_1399, sima_alalam_visit_Bahman_1399,
#                 sima_alkosar_visit_Bahman_1399, sima_presstv_visit_Bahman_1399, sima_sepehr_visit_Bahman_1399,],
#        'duration': [sima_1_duration_Bahman_1399, sima_2_duration_Bahman_1399, sima_3_duration_Bahman_1399,
#                 sima_4_duration_Bahman_1399, sima_5_duration_Bahman_1399, sima_khabar_duration_Bahman_1399,
#                 sima_ofogh_duration_Bahman_1399, sima_pooya_duration_Bahman_1399, sima_omid_duration_Bahman_1399,
#                 sima_ifilm_duration_Bahman_1399, sima_namayesh_duration_Bahman_1399, sima_tamasha_duration_Bahman_1399,
#                 sima_mostanad_duration_Bahman_1399, sima_shoma_duration_Bahman_1399, sima_amozesh_duration_Bahman_1399,
#                 sima_varzesh_duration_Bahman_1399, sima_nasim_duration_Bahman_1399, sima_qoran_duration_Bahman_1399,
#                 sima_salamat_duration_Bahman_1399, sima_irankala_duration_Bahman_1399, sima_alalam_duration_Bahman_1399,
#                 sima_alkosar_duration_Bahman_1399, sima_presstv_duration_Bahman_1399, sima_sepehr_duration_Bahman_1399,],}
#Bahman_1399_sima_visit_channels=pd.DataFrame(Bahman_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])
#
#Bahman_1399_sima_visit_channels=Bahman_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})
#
#Bahman_1399_operator_data=pd.DataFrame()
#Bahman_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
#       'visit': [sima_lenz_visit_Bahman_1399, sima_aio_visit_Bahman_1399, sima_anten_visit_Bahman_1399,
#                 sima_tva_visit_Bahman_1399, sima_fam_visit_Bahman_1399, sima_televebion_visit_Bahman_1399,
#                 sima_sepehr_visit_Bahman_1399, sima_shima_visit_Bahman_1399, sima_site_visit_Bahman_1399,],
#       'register': [register_user_lenz_Bahman_1399, register_user_aio_Bahman_1399, register_user_anten_Bahman_1399,
#                 register_user_tva_Bahman_1399, register_user_fam_Bahman_1399, register_user_televebion_Bahman_1399,
#                 register_user_sepehr_Bahman_1399, register_user_shima_Bahman_1399, register_user_site_Bahman_1399,],
#       'active': [active_user_lenz_Bahman_1399, active_user_aio_Bahman_1399, active_user_anten_Bahman_1399,
#                 active_user_tva_Bahman_1399, active_user_fam_Bahman_1399, active_user_televebion_Bahman_1399,
#                 active_user_sepehr_Bahman_1399, active_user_shima_Bahman_1399, active_user_site_Bahman_1399,],}
#
#Bahman_1399_operator_data=pd.DataFrame(Bahman_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])
#
#Bahman_1399_operator_data=Bahman_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})
#
#Bahman_1399_all_data_summary=pd.DataFrame()
#Bahman_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
#       'statistics': [all_visit_Bahman_1399, all_duration_Bahman_1399,all_content_sima_Bahman_1399,
#                      all_register_user_Bahman_1399, all_active_user_Bahman_1399,],}
#
#Bahman_1399_all_data_summary=pd.DataFrame(Bahman_1399_all_data_summary, columns=['parameters', 'statistics'])
#
#Bahman_1399_all_data_summary=Bahman_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
#
#writer = pd.ExcelWriter('output/ماه بهمن 1399.xlsx', engine='xlsxwriter')
#Bahman_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
#Bahman_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
#Bahman_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه بهمن')
#writer.save()
# 
#  ########################### اسفند #############################
#print("EPG Esfand 1399")
#EPG_Esfand_1399=pd.read_excel('EPG/EPG 1399/EPG Esfand 1399.xlsx', sheet_name='آمار')
#EPG_Esfand_1399.fillna(0, inplace=True)
#sima_1_visit_Esfand_1399=EPG_Esfand_1399.iat[1, 4]
#sima_2_visit_Esfand_1399=EPG_Esfand_1399.iat[2, 4]
#sima_3_visit_Esfand_1399=EPG_Esfand_1399.iat[3, 4]
#sima_4_visit_Esfand_1399=EPG_Esfand_1399.iat[4, 4]
#sima_5_visit_Esfand_1399=EPG_Esfand_1399.iat[5, 4]
#sima_khabar_visit_Esfand_1399=EPG_Esfand_1399.iat[6, 4]
#sima_ofogh_visit_Esfand_1399=EPG_Esfand_1399.iat[7, 4]
#sima_pooya_visit_Esfand_1399=EPG_Esfand_1399.iat[8, 4]
#sima_omid_visit_Esfand_1399=EPG_Esfand_1399.iat[9, 4]
#sima_ifilm_visit_Esfand_1399=EPG_Esfand_1399.iat[10, 4]
#sima_namayesh_visit_Esfand_1399=EPG_Esfand_1399.iat[11, 4]
#sima_tamasha_visit_Esfand_1399=EPG_Esfand_1399.iat[12, 4]
#sima_mostanad_visit_Esfand_1399=EPG_Esfand_1399.iat[13, 4]
#sima_shoma_visit_Esfand_1399=EPG_Esfand_1399.iat[14, 4]
#sima_amozesh_visit_Esfand_1399=EPG_Esfand_1399.iat[15, 4]
#sima_varzesh_visit_Esfand_1399=EPG_Esfand_1399.iat[16, 4]
#sima_nasim_visit_Esfand_1399=EPG_Esfand_1399.iat[17, 4]
#sima_qoran_visit_Esfand_1399=EPG_Esfand_1399.iat[18, 4]
#sima_salamat_visit_Esfand_1399=EPG_Esfand_1399.iat[19, 4]
#sima_irankala_visit_Esfand_1399=EPG_Esfand_1399.iat[20, 4]
#sima_alalam_visit_Esfand_1399=EPG_Esfand_1399.iat[21, 4]
#sima_alkosar_visit_Esfand_1399=EPG_Esfand_1399.iat[22, 4]
#sima_presstv_visit_Esfand_1399=EPG_Esfand_1399.iat[23, 4]
#sima_sepehr_visit_Esfand_1399=EPG_Esfand_1399.iat[24, 4]
#
#sima_1_duration_Esfand_1399=EPG_Esfand_1399.iat[1, 6]
#sima_2_duration_Esfand_1399=EPG_Esfand_1399.iat[2, 6]
#sima_3_duration_Esfand_1399=EPG_Esfand_1399.iat[3, 6]
#sima_4_duration_Esfand_1399=EPG_Esfand_1399.iat[4, 6]
#sima_5_duration_Esfand_1399=EPG_Esfand_1399.iat[5, 6]
#sima_khabar_duration_Esfand_1399=EPG_Esfand_1399.iat[6, 6]
#sima_ofogh_duration_Esfand_1399=EPG_Esfand_1399.iat[7, 6]
#sima_pooya_duration_Esfand_1399=EPG_Esfand_1399.iat[8, 6]
#sima_omid_duration_Esfand_1399=EPG_Esfand_1399.iat[9, 6]
#sima_ifilm_duration_Esfand_1399=EPG_Esfand_1399.iat[10, 6]
#sima_namayesh_duration_Esfand_1399=EPG_Esfand_1399.iat[11, 6]
#sima_tamasha_duration_Esfand_1399=EPG_Esfand_1399.iat[12, 6]
#sima_mostanad_duration_Esfand_1399=EPG_Esfand_1399.iat[13, 6]
#sima_shoma_duration_Esfand_1399=EPG_Esfand_1399.iat[14, 6]
#sima_amozesh_duration_Esfand_1399=EPG_Esfand_1399.iat[15, 6]
#sima_varzesh_duration_Esfand_1399=EPG_Esfand_1399.iat[16, 6]
#sima_nasim_duration_Esfand_1399=EPG_Esfand_1399.iat[17, 6]
#sima_qoran_duration_Esfand_1399=EPG_Esfand_1399.iat[18, 6]
#sima_salamat_duration_Esfand_1399=EPG_Esfand_1399.iat[19, 6]
#sima_irankala_duration_Esfand_1399=EPG_Esfand_1399.iat[20, 6]
#sima_alalam_duration_Esfand_1399=EPG_Esfand_1399.iat[21, 6]
#sima_alkosar_duration_Esfand_1399=EPG_Esfand_1399.iat[22, 6]
#sima_presstv_duration_Esfand_1399=EPG_Esfand_1399.iat[23, 6]
#sima_sepehr_duration_Esfand_1399=EPG_Esfand_1399.iat[24, 6]
#
#sima_lenz_visit_Esfand_1399=EPG_Esfand_1399.iat[33, 2]
#sima_aio_visit_Esfand_1399=EPG_Esfand_1399.iat[34, 2]
#sima_anten_visit_Esfand_1399=EPG_Esfand_1399.iat[35, 2]
#sima_tva_visit_Esfand_1399=EPG_Esfand_1399.iat[36, 2]
#sima_fam_visit_Esfand_1399=EPG_Esfand_1399.iat[37, 2]
#sima_televebion_visit_Esfand_1399=EPG_Esfand_1399.iat[38, 2]
#sima_sepehr_Esfand_1399=EPG_Esfand_1399.iat[39, 2]
#sima_shima_visit_Esfand_1399=EPG_Esfand_1399.iat[40, 2]
#sima_site_visit_Esfand_1399=EPG_Esfand_1399.iat[41, 2]
#
#register_user_lenz_Esfand_1399=EPG_Esfand_1399.iat[33, 4]
#register_user_aio_Esfand_1399=EPG_Esfand_1399.iat[34, 4]
#register_user_anten_Esfand_1399=EPG_Esfand_1399.iat[35, 4]
#register_user_tva_Esfand_1399=EPG_Esfand_1399.iat[36, 4]
#register_user_fam_Esfand_1399=EPG_Esfand_1399.iat[37, 4]
#register_user_televebion_Esfand_1399=EPG_Esfand_1399.iat[38, 4]
#register_user_sepehr_Esfand_1399=EPG_Esfand_1399.iat[39, 4]
#register_user_shima_Esfand_1399=EPG_Esfand_1399.iat[40, 4]
#register_user_site_Esfand_1399=EPG_Esfand_1399.iat[41, 4]
#
#active_user_lenz_Esfand_1399=EPG_Esfand_1399.iat[33, 10]
#active_user_aio_Esfand_1399=EPG_Esfand_1399.iat[34, 10]
#active_user_anten_Esfand_1399=EPG_Esfand_1399.iat[35, 10]
#active_user_tva_Esfand_1399=EPG_Esfand_1399.iat[36, 10]
#active_user_fam_Esfand_1399=EPG_Esfand_1399.iat[37, 10]
#active_user_televebion_Esfand_1399=EPG_Esfand_1399.iat[38, 10]
#active_user_sepehr_Esfand_1399=EPG_Esfand_1399.iat[39, 10]
#active_user_shima_Esfand_1399=EPG_Esfand_1399.iat[40, 10]
#active_user_site_Esfand_1399=EPG_Esfand_1399.iat[41, 10]
#
#all_visit_Esfand_1399=EPG_Esfand_1399.iat[25, 4]
#all_duration_Esfand_1399=EPG_Esfand_1399.iat[25, 6]
#all_content_sima_Esfand_1399=EPG_Esfand_1399.iat[25, 2]
#all_register_user_Esfand_1399=sum(EPG_Esfand_1399.iloc[33:43, 4])
#all_active_user_Esfand_1399=sum(EPG_Esfand_1399.iloc[33:43, 10])
#
#Esfand_1399_sima_visit_channels=pd.DataFrame()
#Esfand_1399_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
#                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
#                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
#                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
#                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
#       'visit': [sima_1_visit_Esfand_1399, sima_2_visit_Esfand_1399, sima_3_visit_Esfand_1399,
#                 sima_4_visit_Esfand_1399, sima_5_visit_Esfand_1399, sima_khabar_visit_Esfand_1399,
#                 sima_ofogh_visit_Esfand_1399, sima_pooya_visit_Esfand_1399, sima_omid_visit_Esfand_1399,
#                 sima_ifilm_visit_Esfand_1399, sima_namayesh_visit_Esfand_1399, sima_tamasha_visit_Esfand_1399,
#                 sima_mostanad_visit_Esfand_1399, sima_shoma_visit_Esfand_1399, sima_amozesh_visit_Esfand_1399,
#                 sima_varzesh_visit_Esfand_1399, sima_nasim_visit_Esfand_1399, sima_qoran_visit_Esfand_1399,
#                 sima_salamat_visit_Esfand_1399, sima_irankala_visit_Esfand_1399, sima_alalam_visit_Esfand_1399,
#                 sima_alkosar_visit_Esfand_1399, sima_presstv_visit_Esfand_1399, sima_sepehr_visit_Esfand_1399,],
#        'duration': [sima_1_duration_Esfand_1399, sima_2_duration_Esfand_1399, sima_3_duration_Esfand_1399,
#                 sima_4_duration_Esfand_1399, sima_5_duration_Esfand_1399, sima_khabar_duration_Esfand_1399,
#                 sima_ofogh_duration_Esfand_1399, sima_pooya_duration_Esfand_1399, sima_omid_duration_Esfand_1399,
#                 sima_ifilm_duration_Esfand_1399, sima_namayesh_duration_Esfand_1399, sima_tamasha_duration_Esfand_1399,
#                 sima_mostanad_duration_Esfand_1399, sima_shoma_duration_Esfand_1399, sima_amozesh_duration_Esfand_1399,
#                 sima_varzesh_duration_Esfand_1399, sima_nasim_duration_Esfand_1399, sima_qoran_duration_Esfand_1399,
#                 sima_salamat_duration_Esfand_1399, sima_irankala_duration_Esfand_1399, sima_alalam_duration_Esfand_1399,
#                 sima_alkosar_duration_Esfand_1399, sima_presstv_duration_Esfand_1399, sima_sepehr_duration_Esfand_1399,],}
#Esfand_1399_sima_visit_channels=pd.DataFrame(Esfand_1399_sima_visit_channels, columns=['channels', 'visit', 'duration'])
#
#Esfand_1399_sima_visit_channels=Esfand_1399_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})
#
#Esfand_1399_operator_data=pd.DataFrame()
#Esfand_1399_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
#       'visit': [sima_lenz_visit_Esfand_1399, sima_aio_visit_Esfand_1399, sima_anten_visit_Esfand_1399,
#                 sima_tva_visit_Esfand_1399, sima_fam_visit_Esfand_1399, sima_televebion_visit_Esfand_1399,
#                 sima_sepehr_visit_Esfand_1399, sima_shima_visit_Esfand_1399, sima_site_visit_Esfand_1399,],
#       'register': [register_user_lenz_Esfand_1399, register_user_aio_Esfand_1399, register_user_anten_Esfand_1399,
#                 register_user_tva_Esfand_1399, register_user_fam_Esfand_1399, register_user_televebion_Esfand_1399,
#                 register_user_sepehr_Esfand_1399, register_user_shima_Esfand_1399, register_user_site_Esfand_1399,],
#       'active': [active_user_lenz_Esfand_1399, active_user_aio_Esfand_1399, active_user_anten_Esfand_1399,
#                 active_user_tva_Esfand_1399, active_user_fam_Esfand_1399, active_user_televebion_Esfand_1399,
#                 active_user_sepehr_Esfand_1399, active_user_shima_Esfand_1399, active_user_site_Esfand_1399,],}
#
#Esfand_1399_operator_data=pd.DataFrame(Esfand_1399_operator_data, columns=['operators', 'visit', 'register', 'active'])
#
#Esfand_1399_operator_data=Esfand_1399_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})
#
#Esfand_1399_all_data_summary=pd.DataFrame()
#Esfand_1399_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
#       'statistics': [all_visit_Esfand_1399, all_duration_Esfand_1399,all_content_sima_Esfand_1399,
#                      all_register_user_Esfand_1399, all_active_user_Esfand_1399,],}
#
#Esfand_1399_all_data_summary=pd.DataFrame(Esfand_1399_all_data_summary, columns=['parameters', 'statistics'])
#
#Esfand_1399_all_data_summary=Esfand_1399_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
#
#writer = pd.ExcelWriter('output/ماه اسفند 1399.xlsx', engine='xlsxwriter')
#Esfand_1399_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
#Esfand_1399_operator_data.to_excel(writer, 'آمار اپراتورها')
#Esfand_1399_all_data_summary.to_excel(writer, 'خلاصه آمار ماه اسفند')
#writer.save()
#  
######################### آمار سال 1398 #########################

        ########################### فروردین #############################
print("EPG Farvardin 1398")
EPG_Farvardin_1398=pd.read_excel('EPG/EPG 1398/EPG Farvardin 1398.xlsx', sheet_name='آمار')
EPG_Farvardin_1398.fillna(0, inplace=True)
EPG_Farvardin_1398['مدت زمان بازدید']=EPG_Farvardin_1398.iloc[0:25, 9]*60
sima_1_visit_Farvardin_1398=EPG_Farvardin_1398.iat[0, 1]
sima_2_visit_Farvardin_1398=EPG_Farvardin_1398.iat[1, 1]
sima_3_visit_Farvardin_1398=EPG_Farvardin_1398.iat[2, 1]
sima_4_visit_Farvardin_1398=EPG_Farvardin_1398.iat[3, 1]
sima_5_visit_Farvardin_1398=EPG_Farvardin_1398.iat[4, 1]
sima_khabar_visit_Farvardin_1398=EPG_Farvardin_1398.iat[5, 1]
sima_ofogh_visit_Farvardin_1398=EPG_Farvardin_1398.iat[6, 1]
sima_pooya_visit_Farvardin_1398=EPG_Farvardin_1398.iat[7, 1]
sima_omid_visit_Farvardin_1398=EPG_Farvardin_1398.iat[8, 1]
sima_ifilm_visit_Farvardin_1398=EPG_Farvardin_1398.iat[9, 1]
sima_namayesh_visit_Farvardin_1398=EPG_Farvardin_1398.iat[10, 1]
sima_tamasha_visit_Farvardin_1398=EPG_Farvardin_1398.iat[11, 1]
sima_mostanad_visit_Farvardin_1398=EPG_Farvardin_1398.iat[12, 1]
sima_shoma_visit_Farvardin_1398=EPG_Farvardin_1398.iat[13, 1]
sima_amozesh_visit_Farvardin_1398=EPG_Farvardin_1398.iat[14, 1]
sima_varzesh_visit_Farvardin_1398=EPG_Farvardin_1398.iat[15, 1]
sima_nasim_visit_Farvardin_1398=EPG_Farvardin_1398.iat[16, 1]
sima_qoran_visit_Farvardin_1398=EPG_Farvardin_1398.iat[17, 1]
sima_salamat_visit_Farvardin_1398=EPG_Farvardin_1398.iat[18, 1]
sima_irankala_visit_Farvardin_1398=EPG_Farvardin_1398.iat[19, 1]
sima_alalam_visit_Farvardin_1398=EPG_Farvardin_1398.iat[20, 1]
sima_alkosar_visit_Farvardin_1398=EPG_Farvardin_1398.iat[21, 1]
sima_presstv_visit_Farvardin_1398=EPG_Farvardin_1398.iat[22, 1]
sima_sepehr_visit_Farvardin_1398=EPG_Farvardin_1398.iat[23, 1]

sima_1_duration_Farvardin_1398=EPG_Farvardin_1398.iat[0, 9]
sima_2_duration_Farvardin_1398=EPG_Farvardin_1398.iat[1, 9]
sima_3_duration_Farvardin_1398=EPG_Farvardin_1398.iat[2, 9]
sima_4_duration_Farvardin_1398=EPG_Farvardin_1398.iat[3, 9]
sima_5_duration_Farvardin_1398=EPG_Farvardin_1398.iat[4, 9]
sima_khabar_duration_Farvardin_1398=EPG_Farvardin_1398.iat[5, 9]
sima_ofogh_duration_Farvardin_1398=EPG_Farvardin_1398.iat[6, 9]
sima_pooya_duration_Farvardin_1398=EPG_Farvardin_1398.iat[7, 9]
sima_omid_duration_Farvardin_1398=EPG_Farvardin_1398.iat[8, 9]
sima_ifilm_duration_Farvardin_1398=EPG_Farvardin_1398.iat[9, 9]
sima_namayesh_duration_Farvardin_1398=EPG_Farvardin_1398.iat[10, 9]
sima_tamasha_duration_Farvardin_1398=EPG_Farvardin_1398.iat[11, 9]
sima_mostanad_duration_Farvardin_1398=EPG_Farvardin_1398.iat[12, 9]
sima_shoma_duration_Farvardin_1398=EPG_Farvardin_1398.iat[13, 9]
sima_amozesh_duration_Farvardin_1398=EPG_Farvardin_1398.iat[14, 9]
sima_varzesh_duration_Farvardin_1398=EPG_Farvardin_1398.iat[15, 9]
sima_nasim_duration_Farvardin_1398=EPG_Farvardin_1398.iat[16, 9]
sima_qoran_duration_Farvardin_1398=EPG_Farvardin_1398.iat[17, 9]
sima_salamat_duration_Farvardin_1398=EPG_Farvardin_1398.iat[18, 9]
sima_irankala_duration_Farvardin_1398=EPG_Farvardin_1398.iat[19, 9]
sima_alalam_duration_Farvardin_1398=EPG_Farvardin_1398.iat[20, 9]
sima_alkosar_duration_Farvardin_1398=EPG_Farvardin_1398.iat[21, 9]
sima_presstv_duration_Farvardin_1398=EPG_Farvardin_1398.iat[22, 9]
sima_sepehr_duration_Farvardin_1398=EPG_Farvardin_1398.iat[23, 9]

sima_lenz_visit_Farvardin_1398=EPG_Farvardin_1398.iat[35, 2]
sima_aio_visit_Farvardin_1398=EPG_Farvardin_1398.iat[36, 2]
sima_anten_visit_Farvardin_1398=EPG_Farvardin_1398.iat[37, 2]
sima_tva_visit_Farvardin_1398=EPG_Farvardin_1398.iat[38, 2]
sima_fam_visit_Farvardin_1398=EPG_Farvardin_1398.iat[39, 2]
#sima_televebion_visit_Farvardin_1398=EPG_Farvardin_1398.iat[38, 2]
#sima_sepehr_Farvardin_1398=EPG_Farvardin_1398.iat[39, 2]
#sima_shima_visit_Farvardin_1398=EPG_Farvardin_1398.iat[40, 2]
#sima_site_visit_Farvardin_1398=EPG_Farvardin_1398.iat[41, 2]

register_user_lenz_Farvardin_1398=EPG_Farvardin_1398.iat[35, 7]
register_user_aio_Farvardin_1398=EPG_Farvardin_1398.iat[36, 7]
register_user_anten_Farvardin_1398=EPG_Farvardin_1398.iat[37, 7]
register_user_tva_Farvardin_1398=EPG_Farvardin_1398.iat[38, 7]
register_user_fam_Farvardin_1398=EPG_Farvardin_1398.iat[39, 7]
#register_user_televebion_Farvardin_1398=EPG_Farvardin_1398.iat[41, 7]
#register_user_sepehr_Farvardin_1398=EPG_Farvardin_1398.iat[42, 7]
#register_user_shima_Farvardin_1398=EPG_Farvardin_1398.iat[43, 7]
#register_user_site_Farvardin_1398=EPG_Farvardin_1398.iat[44, 7]

#active_user_lenz_Farvardin_1398=EPG_Farvardin_1398.iat[36, 10]
#active_user_aio_Farvardin_1398=EPG_Farvardin_1398.iat[37, 10]
#active_user_anten_Farvardin_1398=EPG_Farvardin_1398.iat[38, 10]
#active_user_tva_Farvardin_1398=EPG_Farvardin_1398.iat[39, 10]
#active_user_fam_Farvardin_1398=EPG_Farvardin_1398.iat[40, 10]
#active_user_televebion_Farvardin_1398=EPG_Farvardin_1398.iat[41, 10]
#active_user_sepehr_Farvardin_1398=EPG_Farvardin_1398.iat[42, 10]
#active_user_shima_Farvardin_1398=EPG_Farvardin_1398.iat[43, 10]
#active_user_site_Farvardin_1398=EPG_Farvardin_1398.iat[44, 10]

all_visit_Farvardin_1398=EPG_Farvardin_1398.iat[24, 1]
all_duration_Farvardin_1398=sum(EPG_Farvardin_1398.iloc[0:24, 9])
all_content_sima_Farvardin_1398=EPG_Farvardin_1398.iat[24, 4]
all_register_user_Farvardin_1398=sum(EPG_Farvardin_1398.iloc[35:40, 7])
#all_active_user_Farvardin_1398=sum(EPG_Farvardin_1398.iloc[36:44, 10])

Farvardin_1398_sima_visit_channels=pd.DataFrame()
Farvardin_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Farvardin_1398, sima_2_visit_Farvardin_1398, sima_3_visit_Farvardin_1398,
                 sima_4_visit_Farvardin_1398, sima_5_visit_Farvardin_1398, sima_khabar_visit_Farvardin_1398,
                 sima_ofogh_visit_Farvardin_1398, sima_pooya_visit_Farvardin_1398, sima_omid_visit_Farvardin_1398,
                 sima_ifilm_visit_Farvardin_1398, sima_namayesh_visit_Farvardin_1398, sima_tamasha_visit_Farvardin_1398,
                 sima_mostanad_visit_Farvardin_1398, sima_shoma_visit_Farvardin_1398, sima_amozesh_visit_Farvardin_1398,
                 sima_varzesh_visit_Farvardin_1398, sima_nasim_visit_Farvardin_1398, sima_qoran_visit_Farvardin_1398,
                 sima_salamat_visit_Farvardin_1398, sima_irankala_visit_Farvardin_1398, sima_alalam_visit_Farvardin_1398,
                 sima_alkosar_visit_Farvardin_1398, sima_presstv_visit_Farvardin_1398, sima_sepehr_visit_Farvardin_1398,],
        'duration': [sima_1_duration_Farvardin_1398, sima_2_duration_Farvardin_1398, sima_3_duration_Farvardin_1398,
                 sima_4_duration_Farvardin_1398, sima_5_duration_Farvardin_1398, sima_khabar_duration_Farvardin_1398,
                 sima_ofogh_duration_Farvardin_1398, sima_pooya_duration_Farvardin_1398, sima_omid_duration_Farvardin_1398,
                 sima_ifilm_duration_Farvardin_1398, sima_namayesh_duration_Farvardin_1398, sima_tamasha_duration_Farvardin_1398,
                 sima_mostanad_duration_Farvardin_1398, sima_shoma_duration_Farvardin_1398, sima_amozesh_duration_Farvardin_1398,
                 sima_varzesh_duration_Farvardin_1398, sima_nasim_duration_Farvardin_1398, sima_qoran_duration_Farvardin_1398,
                 sima_salamat_duration_Farvardin_1398, sima_irankala_duration_Farvardin_1398, sima_alalam_duration_Farvardin_1398,
                 sima_alkosar_duration_Farvardin_1398, sima_presstv_duration_Farvardin_1398, sima_sepehr_duration_Farvardin_1398,],}
Farvardin_1398_sima_visit_channels=pd.DataFrame(Farvardin_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Farvardin_1398_sima_visit_channels=Farvardin_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Farvardin_1398_operator_data=pd.DataFrame()
Farvardin_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام'],
       'visit': [sima_lenz_visit_Farvardin_1398, sima_aio_visit_Farvardin_1398, sima_anten_visit_Farvardin_1398,
                 sima_tva_visit_Farvardin_1398, sima_fam_visit_Farvardin_1398,],
       'register': [register_user_lenz_Farvardin_1398, register_user_aio_Farvardin_1398, register_user_anten_Farvardin_1398,
                 register_user_tva_Farvardin_1398, register_user_fam_Farvardin_1398,],}

Farvardin_1398_operator_data=pd.DataFrame(Farvardin_1398_operator_data, columns=['operators', 'visit', 'register'])

Farvardin_1398_operator_data=Farvardin_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی'})

Farvardin_1398_all_data_summary=pd.DataFrame()
Farvardin_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Farvardin_1398, all_duration_Farvardin_1398,all_content_sima_Farvardin_1398, all_register_user_Farvardin_1398,],}

Farvardin_1398_all_data_summary=pd.DataFrame(Farvardin_1398_all_data_summary, columns=['parameters', 'statistics'])

Farvardin_1398_all_data_summary=Farvardin_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه فروردین 1398.xlsx', engine='xlsxwriter')
Farvardin_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Farvardin_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Farvardin_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه فروردین')
writer.save()

        ########################### اردیبهشت #############################
print("EPG Ordibehesht 1398")
EPG_Ordibehesht_1398=pd.read_excel('EPG/EPG 1398/EPG Ordibehesht 1398.xlsx', sheet_name='آمار')
EPG_Ordibehesht_1398.fillna(0, inplace=True)
EPG_Ordibehesht_1398['مدت زمان بازدید']=EPG_Ordibehesht_1398.iloc[0:25, 9]*60
sima_1_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[0, 1]
sima_2_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[1, 1]
sima_3_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[2, 1]
sima_4_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[3, 1]
sima_5_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[4, 1]
sima_khabar_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[5, 1]
sima_ofogh_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[6, 1]
sima_pooya_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[7, 1]
sima_omid_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[8, 1]
sima_ifilm_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[9, 1]
sima_namayesh_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[10, 1]
sima_tamasha_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[11, 1]
sima_mostanad_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[12, 1]
sima_shoma_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[13, 1]
sima_amozesh_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[14, 1]
sima_varzesh_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[15, 1]
sima_nasim_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[16, 1]
sima_qoran_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[17, 1]
sima_salamat_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[18, 1]
sima_irankala_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[19, 1]
sima_alalam_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[20, 1]
sima_alkosar_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[21, 1]
sima_presstv_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[22, 1]
sima_sepehr_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[23, 1]

sima_1_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[0, 9]
sima_2_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[1, 9]
sima_3_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[2, 9]
sima_4_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[3, 9]
sima_5_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[4, 9]
sima_khabar_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[5, 9]
sima_ofogh_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[6, 9]
sima_pooya_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[7, 9]
sima_omid_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[8, 9]
sima_ifilm_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[9, 9]
sima_namayesh_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[10, 9]
sima_tamasha_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[11, 9]
sima_mostanad_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[12, 9]
sima_shoma_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[13, 9]
sima_amozesh_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[14, 9]
sima_varzesh_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[15, 9]
sima_nasim_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[16, 9]
sima_qoran_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[17, 9]
sima_salamat_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[18, 9]
sima_irankala_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[19, 9]
sima_alalam_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[20, 9]
sima_alkosar_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[21, 9]
sima_presstv_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[22, 9]
sima_sepehr_duration_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[23, 9]

sima_lenz_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[35, 2]
sima_aio_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[36, 2]
sima_anten_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[37, 2]
sima_tva_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[38, 2]
sima_fam_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[39, 2]
#sima_televebion_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[38, 2]
#sima_sepehr_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[39, 2]
#sima_shima_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[40, 2]
#sima_site_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[41, 2]

register_user_lenz_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[35, 7]
register_user_aio_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[36, 7]
register_user_anten_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[37, 7]
register_user_tva_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[38, 7]
register_user_fam_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[39, 7]
#register_user_televebion_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[41, 7]
#register_user_sepehr_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[42, 7]
#register_user_shima_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[43, 7]
#register_user_site_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[44, 7]

#active_user_lenz_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[36, 10]
#active_user_aio_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[37, 10]
#active_user_anten_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[38, 10]
#active_user_tva_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[39, 10]
#active_user_fam_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[40, 10]
#active_user_televebion_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[41, 10]
#active_user_sepehr_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[42, 10]
#active_user_shima_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[43, 10]
#active_user_site_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[44, 10]

all_visit_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[24, 1]
all_duration_Ordibehesht_1398=sum(EPG_Ordibehesht_1398.iloc[0:24, 9])
all_content_sima_Ordibehesht_1398=EPG_Ordibehesht_1398.iat[24, 4]
all_register_user_Ordibehesht_1398=sum(EPG_Ordibehesht_1398.iloc[35:40, 7])
#all_active_user_Ordibehesht_1398=sum(EPG_Ordibehesht_1398.iloc[36:44, 10])

Ordibehesht_1398_sima_visit_channels=pd.DataFrame()
Ordibehesht_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Ordibehesht_1398, sima_2_visit_Ordibehesht_1398, sima_3_visit_Ordibehesht_1398,
                 sima_4_visit_Ordibehesht_1398, sima_5_visit_Ordibehesht_1398, sima_khabar_visit_Ordibehesht_1398,
                 sima_ofogh_visit_Ordibehesht_1398, sima_pooya_visit_Ordibehesht_1398, sima_omid_visit_Ordibehesht_1398,
                 sima_ifilm_visit_Ordibehesht_1398, sima_namayesh_visit_Ordibehesht_1398, sima_tamasha_visit_Ordibehesht_1398,
                 sima_mostanad_visit_Ordibehesht_1398, sima_shoma_visit_Ordibehesht_1398, sima_amozesh_visit_Ordibehesht_1398,
                 sima_varzesh_visit_Ordibehesht_1398, sima_nasim_visit_Ordibehesht_1398, sima_qoran_visit_Ordibehesht_1398,
                 sima_salamat_visit_Ordibehesht_1398, sima_irankala_visit_Ordibehesht_1398, sima_alalam_visit_Ordibehesht_1398,
                 sima_alkosar_visit_Ordibehesht_1398, sima_presstv_visit_Ordibehesht_1398, sima_sepehr_visit_Ordibehesht_1398,],
        'duration': [sima_1_duration_Ordibehesht_1398, sima_2_duration_Ordibehesht_1398, sima_3_duration_Ordibehesht_1398,
                 sima_4_duration_Ordibehesht_1398, sima_5_duration_Ordibehesht_1398, sima_khabar_duration_Ordibehesht_1398,
                 sima_ofogh_duration_Ordibehesht_1398, sima_pooya_duration_Ordibehesht_1398, sima_omid_duration_Ordibehesht_1398,
                 sima_ifilm_duration_Ordibehesht_1398, sima_namayesh_duration_Ordibehesht_1398, sima_tamasha_duration_Ordibehesht_1398,
                 sima_mostanad_duration_Ordibehesht_1398, sima_shoma_duration_Ordibehesht_1398, sima_amozesh_duration_Ordibehesht_1398,
                 sima_varzesh_duration_Ordibehesht_1398, sima_nasim_duration_Ordibehesht_1398, sima_qoran_duration_Ordibehesht_1398,
                 sima_salamat_duration_Ordibehesht_1398, sima_irankala_duration_Ordibehesht_1398, sima_alalam_duration_Ordibehesht_1398,
                 sima_alkosar_duration_Ordibehesht_1398, sima_presstv_duration_Ordibehesht_1398, sima_sepehr_duration_Ordibehesht_1398,],}
Ordibehesht_1398_sima_visit_channels=pd.DataFrame(Ordibehesht_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Ordibehesht_1398_sima_visit_channels=Ordibehesht_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Ordibehesht_1398_operator_data=pd.DataFrame()
Ordibehesht_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام'],
       'visit': [sima_lenz_visit_Ordibehesht_1398, sima_aio_visit_Ordibehesht_1398, sima_anten_visit_Ordibehesht_1398,
                 sima_tva_visit_Ordibehesht_1398, sima_fam_visit_Ordibehesht_1398,],
       'register': [register_user_lenz_Ordibehesht_1398, register_user_aio_Ordibehesht_1398, register_user_anten_Ordibehesht_1398,
                 register_user_tva_Ordibehesht_1398, register_user_fam_Ordibehesht_1398,],}

Ordibehesht_1398_operator_data=pd.DataFrame(Ordibehesht_1398_operator_data, columns=['operators', 'visit', 'register'])

Ordibehesht_1398_operator_data=Ordibehesht_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی'})

Ordibehesht_1398_all_data_summary=pd.DataFrame()
Ordibehesht_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Ordibehesht_1398, all_duration_Ordibehesht_1398,all_content_sima_Ordibehesht_1398, all_register_user_Ordibehesht_1398,],}

Ordibehesht_1398_all_data_summary=pd.DataFrame(Ordibehesht_1398_all_data_summary, columns=['parameters', 'statistics'])

Ordibehesht_1398_all_data_summary=Ordibehesht_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه اردیبهشت 1398.xlsx', engine='xlsxwriter')
Ordibehesht_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Ordibehesht_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Ordibehesht_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه اردیبهشت')
writer.save()

        ########################### خرداد #############################
print("EPG Khordad 1398")
EPG_Khordad_1398=pd.read_excel('EPG/EPG 1398/EPG Khordad 1398.xlsx', sheet_name='آمار')
EPG_Khordad_1398.fillna(0, inplace=True)
EPG_Khordad_1398['مدت زمان بازدید']=EPG_Khordad_1398.iloc[0:25, 9]*60
sima_1_visit_Khordad_1398=EPG_Khordad_1398.iat[0, 3]
sima_2_visit_Khordad_1398=EPG_Khordad_1398.iat[1, 3]
sima_3_visit_Khordad_1398=EPG_Khordad_1398.iat[2, 3]
sima_4_visit_Khordad_1398=EPG_Khordad_1398.iat[3, 3]
sima_5_visit_Khordad_1398=EPG_Khordad_1398.iat[4, 3]
sima_khabar_visit_Khordad_1398=EPG_Khordad_1398.iat[5, 3]
sima_ofogh_visit_Khordad_1398=EPG_Khordad_1398.iat[6, 3]
sima_pooya_visit_Khordad_1398=EPG_Khordad_1398.iat[7, 3]
sima_omid_visit_Khordad_1398=EPG_Khordad_1398.iat[8, 3]
sima_ifilm_visit_Khordad_1398=EPG_Khordad_1398.iat[9, 3]
sima_namayesh_visit_Khordad_1398=EPG_Khordad_1398.iat[10, 3]
sima_tamasha_visit_Khordad_1398=EPG_Khordad_1398.iat[11, 3]
sima_mostanad_visit_Khordad_1398=EPG_Khordad_1398.iat[12, 3]
sima_shoma_visit_Khordad_1398=EPG_Khordad_1398.iat[13, 3]
sima_amozesh_visit_Khordad_1398=EPG_Khordad_1398.iat[14, 3]
sima_varzesh_visit_Khordad_1398=EPG_Khordad_1398.iat[15, 3]
sima_nasim_visit_Khordad_1398=EPG_Khordad_1398.iat[16, 3]
sima_qoran_visit_Khordad_1398=EPG_Khordad_1398.iat[17, 3]
sima_salamat_visit_Khordad_1398=EPG_Khordad_1398.iat[18, 3]
sima_irankala_visit_Khordad_1398=EPG_Khordad_1398.iat[19, 3]
sima_alalam_visit_Khordad_1398=EPG_Khordad_1398.iat[20, 3]
sima_alkosar_visit_Khordad_1398=EPG_Khordad_1398.iat[21, 3]
sima_presstv_visit_Khordad_1398=EPG_Khordad_1398.iat[22, 3]
sima_sepehr_visit_Khordad_1398=EPG_Khordad_1398.iat[23, 3]

sima_1_duration_Khordad_1398=EPG_Khordad_1398.iat[0, 9]
sima_2_duration_Khordad_1398=EPG_Khordad_1398.iat[1, 9]
sima_3_duration_Khordad_1398=EPG_Khordad_1398.iat[2, 9]
sima_4_duration_Khordad_1398=EPG_Khordad_1398.iat[3, 9]
sima_5_duration_Khordad_1398=EPG_Khordad_1398.iat[4, 9]
sima_khabar_duration_Khordad_1398=EPG_Khordad_1398.iat[5, 9]
sima_ofogh_duration_Khordad_1398=EPG_Khordad_1398.iat[6, 9]
sima_pooya_duration_Khordad_1398=EPG_Khordad_1398.iat[7, 9]
sima_omid_duration_Khordad_1398=EPG_Khordad_1398.iat[8, 9]
sima_ifilm_duration_Khordad_1398=EPG_Khordad_1398.iat[9, 9]
sima_namayesh_duration_Khordad_1398=EPG_Khordad_1398.iat[10, 9]
sima_tamasha_duration_Khordad_1398=EPG_Khordad_1398.iat[11, 9]
sima_mostanad_duration_Khordad_1398=EPG_Khordad_1398.iat[12, 9]
sima_shoma_duration_Khordad_1398=EPG_Khordad_1398.iat[13, 9]
sima_amozesh_duration_Khordad_1398=EPG_Khordad_1398.iat[14, 9]
sima_varzesh_duration_Khordad_1398=EPG_Khordad_1398.iat[15, 9]
sima_nasim_duration_Khordad_1398=EPG_Khordad_1398.iat[16, 9]
sima_qoran_duration_Khordad_1398=EPG_Khordad_1398.iat[17, 9]
sima_salamat_duration_Khordad_1398=EPG_Khordad_1398.iat[18, 9]
sima_irankala_duration_Khordad_1398=EPG_Khordad_1398.iat[19, 9]
sima_alalam_duration_Khordad_1398=EPG_Khordad_1398.iat[20, 9]
sima_alkosar_duration_Khordad_1398=EPG_Khordad_1398.iat[21, 9]
sima_presstv_duration_Khordad_1398=EPG_Khordad_1398.iat[22, 9]
sima_sepehr_duration_Khordad_1398=EPG_Khordad_1398.iat[23, 9]

sima_lenz_visit_Khordad_1398=EPG_Khordad_1398.iat[35, 1]
sima_aio_visit_Khordad_1398=EPG_Khordad_1398.iat[36, 1]
sima_anten_visit_Khordad_1398=EPG_Khordad_1398.iat[37, 1]
sima_tva_visit_Khordad_1398=EPG_Khordad_1398.iat[38, 1]
sima_fam_visit_Khordad_1398=EPG_Khordad_1398.iat[39, 1]
#sima_televebion_visit_Khordad_1398=EPG_Khordad_1398.iat[38, 2]
#sima_sepehr_Khordad_1398=EPG_Khordad_1398.iat[39, 2]
#sima_shima_visit_Khordad_1398=EPG_Khordad_1398.iat[40, 2]
#sima_site_visit_Khordad_1398=EPG_Khordad_1398.iat[41, 2]

register_user_lenz_Khordad_1398=EPG_Khordad_1398.iat[35, 3]
register_user_aio_Khordad_1398=EPG_Khordad_1398.iat[36, 3]
register_user_anten_Khordad_1398=EPG_Khordad_1398.iat[37, 3]
register_user_tva_Khordad_1398=EPG_Khordad_1398.iat[38, 3]
register_user_fam_Khordad_1398=EPG_Khordad_1398.iat[39, 3]
#register_user_televebion_Khordad_1398=EPG_Khordad_1398.iat[41, 7]
#register_user_sepehr_Khordad_1398=EPG_Khordad_1398.iat[42, 7]
#register_user_shima_Khordad_1398=EPG_Khordad_1398.iat[43, 7]
#register_user_site_Khordad_1398=EPG_Khordad_1398.iat[44, 7]

#active_user_lenz_Khordad_1398=EPG_Khordad_1398.iat[36, 10]
#active_user_aio_Khordad_1398=EPG_Khordad_1398.iat[37, 10]
#active_user_anten_Khordad_1398=EPG_Khordad_1398.iat[38, 10]
#active_user_tva_Khordad_1398=EPG_Khordad_1398.iat[39, 10]
#active_user_fam_Khordad_1398=EPG_Khordad_1398.iat[40, 10]
#active_user_televebion_Khordad_1398=EPG_Khordad_1398.iat[41, 10]
#active_user_sepehr_Khordad_1398=EPG_Khordad_1398.iat[42, 10]
#active_user_shima_Khordad_1398=EPG_Khordad_1398.iat[43, 10]
#active_user_site_Khordad_1398=EPG_Khordad_1398.iat[44, 10]

all_visit_Khordad_1398=EPG_Khordad_1398.iat[24, 3]
all_duration_Khordad_1398=sum(EPG_Khordad_1398.iloc[0:24, 9])
all_content_sima_Khordad_1398=EPG_Khordad_1398.iat[24, 1]
all_register_user_Khordad_1398=sum(EPG_Khordad_1398.iloc[35:40, 3])
#all_active_user_Khordad_1398=sum(EPG_Khordad_1398.iloc[36:44, 10])

Khordad_1398_sima_visit_channels=pd.DataFrame()
Khordad_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Khordad_1398, sima_2_visit_Khordad_1398, sima_3_visit_Khordad_1398,
                 sima_4_visit_Khordad_1398, sima_5_visit_Khordad_1398, sima_khabar_visit_Khordad_1398,
                 sima_ofogh_visit_Khordad_1398, sima_pooya_visit_Khordad_1398, sima_omid_visit_Khordad_1398,
                 sima_ifilm_visit_Khordad_1398, sima_namayesh_visit_Khordad_1398, sima_tamasha_visit_Khordad_1398,
                 sima_mostanad_visit_Khordad_1398, sima_shoma_visit_Khordad_1398, sima_amozesh_visit_Khordad_1398,
                 sima_varzesh_visit_Khordad_1398, sima_nasim_visit_Khordad_1398, sima_qoran_visit_Khordad_1398,
                 sima_salamat_visit_Khordad_1398, sima_irankala_visit_Khordad_1398, sima_alalam_visit_Khordad_1398,
                 sima_alkosar_visit_Khordad_1398, sima_presstv_visit_Khordad_1398, sima_sepehr_visit_Khordad_1398,],
        'duration': [sima_1_duration_Khordad_1398, sima_2_duration_Khordad_1398, sima_3_duration_Khordad_1398,
                 sima_4_duration_Khordad_1398, sima_5_duration_Khordad_1398, sima_khabar_duration_Khordad_1398,
                 sima_ofogh_duration_Khordad_1398, sima_pooya_duration_Khordad_1398, sima_omid_duration_Khordad_1398,
                 sima_ifilm_duration_Khordad_1398, sima_namayesh_duration_Khordad_1398, sima_tamasha_duration_Khordad_1398,
                 sima_mostanad_duration_Khordad_1398, sima_shoma_duration_Khordad_1398, sima_amozesh_duration_Khordad_1398,
                 sima_varzesh_duration_Khordad_1398, sima_nasim_duration_Khordad_1398, sima_qoran_duration_Khordad_1398,
                 sima_salamat_duration_Khordad_1398, sima_irankala_duration_Khordad_1398, sima_alalam_duration_Khordad_1398,
                 sima_alkosar_duration_Khordad_1398, sima_presstv_duration_Khordad_1398, sima_sepehr_duration_Khordad_1398,],}
Khordad_1398_sima_visit_channels=pd.DataFrame(Khordad_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Khordad_1398_sima_visit_channels=Khordad_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Khordad_1398_operator_data=pd.DataFrame()
Khordad_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام'],
       'visit': [sima_lenz_visit_Khordad_1398, sima_aio_visit_Khordad_1398, sima_anten_visit_Khordad_1398,
                 sima_tva_visit_Khordad_1398, sima_fam_visit_Khordad_1398,],
       'register': [register_user_lenz_Khordad_1398, register_user_aio_Khordad_1398, register_user_anten_Khordad_1398,
                 register_user_tva_Khordad_1398, register_user_fam_Khordad_1398,],}

Khordad_1398_operator_data=pd.DataFrame(Khordad_1398_operator_data, columns=['operators', 'visit', 'register'])

Khordad_1398_operator_data=Khordad_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی'})

Khordad_1398_all_data_summary=pd.DataFrame()
Khordad_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Khordad_1398, all_duration_Khordad_1398,all_content_sima_Khordad_1398, all_register_user_Khordad_1398,],}

Khordad_1398_all_data_summary=pd.DataFrame(Khordad_1398_all_data_summary, columns=['parameters', 'statistics'])

Khordad_1398_all_data_summary=Khordad_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه خرداد 1398.xlsx', engine='xlsxwriter')
Khordad_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Khordad_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Khordad_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه خرداد')
writer.save()

        ########################### تیر #############################
print("EPG Tir 1398")
EPG_Tir_1398=pd.read_excel('EPG/EPG 1398/EPG Tir 1398.xlsx', sheet_name='آمار')
EPG_Tir_1398.fillna(0, inplace=True)
EPG_Tir_1398['مدت زمان بازدید']=EPG_Tir_1398.iloc[1:26, 9]*60
sima_1_visit_Tir_1398=EPG_Tir_1398.iat[1, 4]
sima_2_visit_Tir_1398=EPG_Tir_1398.iat[2, 4]
sima_3_visit_Tir_1398=EPG_Tir_1398.iat[3, 4]
sima_4_visit_Tir_1398=EPG_Tir_1398.iat[4, 4]
sima_5_visit_Tir_1398=EPG_Tir_1398.iat[5, 4]
sima_khabar_visit_Tir_1398=EPG_Tir_1398.iat[6, 4]
sima_ofogh_visit_Tir_1398=EPG_Tir_1398.iat[7, 4]
sima_pooya_visit_Tir_1398=EPG_Tir_1398.iat[8, 4]
sima_omid_visit_Tir_1398=EPG_Tir_1398.iat[9, 4]
sima_ifilm_visit_Tir_1398=EPG_Tir_1398.iat[10, 4]
sima_namayesh_visit_Tir_1398=EPG_Tir_1398.iat[11, 4]
sima_tamasha_visit_Tir_1398=EPG_Tir_1398.iat[12, 4]
sima_mostanad_visit_Tir_1398=EPG_Tir_1398.iat[13, 4]
sima_shoma_visit_Tir_1398=EPG_Tir_1398.iat[14, 4]
sima_amozesh_visit_Tir_1398=EPG_Tir_1398.iat[15, 4]
sima_varzesh_visit_Tir_1398=EPG_Tir_1398.iat[16, 4]
sima_nasim_visit_Tir_1398=EPG_Tir_1398.iat[17, 4]
sima_qoran_visit_Tir_1398=EPG_Tir_1398.iat[18, 4]
sima_salamat_visit_Tir_1398=EPG_Tir_1398.iat[19, 4]
sima_irankala_visit_Tir_1398=EPG_Tir_1398.iat[20, 4]
sima_alalam_visit_Tir_1398=EPG_Tir_1398.iat[21, 4]
sima_alkosar_visit_Tir_1398=EPG_Tir_1398.iat[22, 4]
sima_presstv_visit_Tir_1398=EPG_Tir_1398.iat[23, 4]
sima_sepehr_visit_Tir_1398=EPG_Tir_1398.iat[24, 4]

sima_1_duration_Tir_1398=EPG_Tir_1398.iat[1, 6]
sima_2_duration_Tir_1398=EPG_Tir_1398.iat[2, 6]
sima_3_duration_Tir_1398=EPG_Tir_1398.iat[3, 6]
sima_4_duration_Tir_1398=EPG_Tir_1398.iat[4, 6]
sima_5_duration_Tir_1398=EPG_Tir_1398.iat[5, 6]
sima_khabar_duration_Tir_1398=EPG_Tir_1398.iat[6, 6]
sima_ofogh_duration_Tir_1398=EPG_Tir_1398.iat[7, 6]
sima_pooya_duration_Tir_1398=EPG_Tir_1398.iat[8, 6]
sima_omid_duration_Tir_1398=EPG_Tir_1398.iat[9, 6]
sima_ifilm_duration_Tir_1398=EPG_Tir_1398.iat[10, 6]
sima_namayesh_duration_Tir_1398=EPG_Tir_1398.iat[11, 6]
sima_tamasha_duration_Tir_1398=EPG_Tir_1398.iat[12, 6]
sima_mostanad_duration_Tir_1398=EPG_Tir_1398.iat[13, 6]
sima_shoma_duration_Tir_1398=EPG_Tir_1398.iat[14, 6]
sima_amozesh_duration_Tir_1398=EPG_Tir_1398.iat[15, 6]
sima_varzesh_duration_Tir_1398=EPG_Tir_1398.iat[16, 6]
sima_nasim_duration_Tir_1398=EPG_Tir_1398.iat[17, 6]
sima_qoran_duration_Tir_1398=EPG_Tir_1398.iat[18, 6]
sima_salamat_duration_Tir_1398=EPG_Tir_1398.iat[19, 6]
sima_irankala_duration_Tir_1398=EPG_Tir_1398.iat[20, 6]
sima_alalam_duration_Tir_1398=EPG_Tir_1398.iat[21, 6]
sima_alkosar_duration_Tir_1398=EPG_Tir_1398.iat[22, 6]
sima_presstv_duration_Tir_1398=EPG_Tir_1398.iat[23, 6]
sima_sepehr_duration_Tir_1398=EPG_Tir_1398.iat[24, 6]

sima_lenz_visit_Tir_1398=EPG_Tir_1398.iat[36, 2]
sima_aio_visit_Tir_1398=EPG_Tir_1398.iat[37, 2]
sima_anten_visit_Tir_1398=EPG_Tir_1398.iat[38, 2]
sima_tva_visit_Tir_1398=EPG_Tir_1398.iat[39, 2]
sima_fam_visit_Tir_1398=EPG_Tir_1398.iat[40, 2]
sima_televebion_visit_Tir_1398=EPG_Tir_1398.iat[41, 2]
#sima_sepehr_Tir_1398=EPG_Tir_1398.iat[39, 2]
#sima_shima_visit_Tir_1398=EPG_Tir_1398.iat[40, 2]
#sima_site_visit_Tir_1398=EPG_Tir_1398.iat[41, 2]

register_user_lenz_Tir_1398=EPG_Tir_1398.iat[36, 4]
register_user_aio_Tir_1398=EPG_Tir_1398.iat[37, 4]
register_user_anten_Tir_1398=EPG_Tir_1398.iat[38, 4]
register_user_tva_Tir_1398=EPG_Tir_1398.iat[39, 4]
register_user_fam_Tir_1398=EPG_Tir_1398.iat[40, 4]
register_user_televebion_Tir_1398=EPG_Tir_1398.iat[41, 4]
#register_user_sepehr_Tir_1398=EPG_Tir_1398.iat[42, 7]
#register_user_shima_Tir_1398=EPG_Tir_1398.iat[43, 7]
#register_user_site_Tir_1398=EPG_Tir_1398.iat[44, 7]

#active_user_lenz_Tir_1398=EPG_Tir_1398.iat[36, 10]
#active_user_aio_Tir_1398=EPG_Tir_1398.iat[37, 10]
#active_user_anten_Tir_1398=EPG_Tir_1398.iat[38, 10]
#active_user_tva_Tir_1398=EPG_Tir_1398.iat[39, 10]
#active_user_fam_Tir_1398=EPG_Tir_1398.iat[40, 10]
#active_user_televebion_Tir_1398=EPG_Tir_1398.iat[41, 10]
#active_user_sepehr_Tir_1398=EPG_Tir_1398.iat[42, 10]
#active_user_shima_Tir_1398=EPG_Tir_1398.iat[43, 10]
#active_user_site_Tir_1398=EPG_Tir_1398.iat[44, 10]

all_visit_Tir_1398=EPG_Tir_1398.iat[25, 4]
all_duration_Tir_1398=sum(EPG_Tir_1398.iloc[1:24, 6])
all_content_sima_Tir_1398=EPG_Tir_1398.iat[25, 2]
all_register_user_Tir_1398=sum(EPG_Tir_1398.iloc[36:42, 4])
#all_active_user_Tir_1398=sum(EPG_Tir_1398.iloc[36:44, 10])

Tir_1398_sima_visit_channels=pd.DataFrame()
Tir_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Tir_1398, sima_2_visit_Tir_1398, sima_3_visit_Tir_1398,
                 sima_4_visit_Tir_1398, sima_5_visit_Tir_1398, sima_khabar_visit_Tir_1398,
                 sima_ofogh_visit_Tir_1398, sima_pooya_visit_Tir_1398, sima_omid_visit_Tir_1398,
                 sima_ifilm_visit_Tir_1398, sima_namayesh_visit_Tir_1398, sima_tamasha_visit_Tir_1398,
                 sima_mostanad_visit_Tir_1398, sima_shoma_visit_Tir_1398, sima_amozesh_visit_Tir_1398,
                 sima_varzesh_visit_Tir_1398, sima_nasim_visit_Tir_1398, sima_qoran_visit_Tir_1398,
                 sima_salamat_visit_Tir_1398, sima_irankala_visit_Tir_1398, sima_alalam_visit_Tir_1398,
                 sima_alkosar_visit_Tir_1398, sima_presstv_visit_Tir_1398, sima_sepehr_visit_Tir_1398,],
        'duration': [sima_1_duration_Tir_1398, sima_2_duration_Tir_1398, sima_3_duration_Tir_1398,
                 sima_4_duration_Tir_1398, sima_5_duration_Tir_1398, sima_khabar_duration_Tir_1398,
                 sima_ofogh_duration_Tir_1398, sima_pooya_duration_Tir_1398, sima_omid_duration_Tir_1398,
                 sima_ifilm_duration_Tir_1398, sima_namayesh_duration_Tir_1398, sima_tamasha_duration_Tir_1398,
                 sima_mostanad_duration_Tir_1398, sima_shoma_duration_Tir_1398, sima_amozesh_duration_Tir_1398,
                 sima_varzesh_duration_Tir_1398, sima_nasim_duration_Tir_1398, sima_qoran_duration_Tir_1398,
                 sima_salamat_duration_Tir_1398, sima_irankala_duration_Tir_1398, sima_alalam_duration_Tir_1398,
                 sima_alkosar_duration_Tir_1398, sima_presstv_duration_Tir_1398, sima_sepehr_duration_Tir_1398,],}
Tir_1398_sima_visit_channels=pd.DataFrame(Tir_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Tir_1398_sima_visit_channels=Tir_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Tir_1398_operator_data=pd.DataFrame()
Tir_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون'],
       'visit': [sima_lenz_visit_Tir_1398, sima_aio_visit_Tir_1398, sima_anten_visit_Tir_1398,
                 sima_tva_visit_Tir_1398, sima_fam_visit_Tir_1398,sima_televebion_visit_Tir_1398,],
       'register': [register_user_lenz_Tir_1398, register_user_aio_Tir_1398, register_user_anten_Tir_1398,
                 register_user_tva_Tir_1398, register_user_fam_Tir_1398, register_user_televebion_Tir_1398,],}

Tir_1398_operator_data=pd.DataFrame(Tir_1398_operator_data, columns=['operators', 'visit', 'register'])

Tir_1398_operator_data=Tir_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی'})

Tir_1398_all_data_summary=pd.DataFrame()
Tir_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Tir_1398, all_duration_Tir_1398,all_content_sima_Tir_1398, all_register_user_Tir_1398,],}

Tir_1398_all_data_summary=pd.DataFrame(Tir_1398_all_data_summary, columns=['parameters', 'statistics'])

Tir_1398_all_data_summary=Tir_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه تیر 1398.xlsx', engine='xlsxwriter')
Tir_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Tir_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Tir_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه تیر')
writer.save()

        ########################### مرداد #############################
print("EPG Mordad 1398")
EPG_Mordad_1398=pd.read_excel('EPG/EPG 1398/EPG Mordad 1398.xlsx', sheet_name='آمار')
EPG_Mordad_1398.fillna(0, inplace=True)
EPG_Mordad_1398['مدت زمان بازدید']=EPG_Mordad_1398.iloc[1:26, 9]*60
sima_1_visit_Mordad_1398=EPG_Mordad_1398.iat[1, 4]
sima_2_visit_Mordad_1398=EPG_Mordad_1398.iat[2, 4]
sima_3_visit_Mordad_1398=EPG_Mordad_1398.iat[3, 4]
sima_4_visit_Mordad_1398=EPG_Mordad_1398.iat[4, 4]
sima_5_visit_Mordad_1398=EPG_Mordad_1398.iat[5, 4]
sima_khabar_visit_Mordad_1398=EPG_Mordad_1398.iat[6, 4]
sima_ofogh_visit_Mordad_1398=EPG_Mordad_1398.iat[7, 4]
sima_pooya_visit_Mordad_1398=EPG_Mordad_1398.iat[8, 4]
sima_omid_visit_Mordad_1398=EPG_Mordad_1398.iat[9, 4]
sima_ifilm_visit_Mordad_1398=EPG_Mordad_1398.iat[10, 4]
sima_namayesh_visit_Mordad_1398=EPG_Mordad_1398.iat[11, 4]
sima_tamasha_visit_Mordad_1398=EPG_Mordad_1398.iat[12, 4]
sima_mostanad_visit_Mordad_1398=EPG_Mordad_1398.iat[13, 4]
sima_shoma_visit_Mordad_1398=EPG_Mordad_1398.iat[14, 4]
sima_amozesh_visit_Mordad_1398=EPG_Mordad_1398.iat[15, 4]
sima_varzesh_visit_Mordad_1398=EPG_Mordad_1398.iat[16, 4]
sima_nasim_visit_Mordad_1398=EPG_Mordad_1398.iat[17, 4]
sima_qoran_visit_Mordad_1398=EPG_Mordad_1398.iat[18, 4]
sima_salamat_visit_Mordad_1398=EPG_Mordad_1398.iat[19, 4]
sima_irankala_visit_Mordad_1398=EPG_Mordad_1398.iat[20, 4]
sima_alalam_visit_Mordad_1398=EPG_Mordad_1398.iat[21, 4]
sima_alkosar_visit_Mordad_1398=EPG_Mordad_1398.iat[22, 4]
sima_presstv_visit_Mordad_1398=EPG_Mordad_1398.iat[23, 4]
sima_sepehr_visit_Mordad_1398=EPG_Mordad_1398.iat[24, 4]

sima_1_duration_Mordad_1398=EPG_Mordad_1398.iat[1, 6]
sima_2_duration_Mordad_1398=EPG_Mordad_1398.iat[2, 6]
sima_3_duration_Mordad_1398=EPG_Mordad_1398.iat[3, 6]
sima_4_duration_Mordad_1398=EPG_Mordad_1398.iat[4, 6]
sima_5_duration_Mordad_1398=EPG_Mordad_1398.iat[5, 6]
sima_khabar_duration_Mordad_1398=EPG_Mordad_1398.iat[6, 6]
sima_ofogh_duration_Mordad_1398=EPG_Mordad_1398.iat[7, 6]
sima_pooya_duration_Mordad_1398=EPG_Mordad_1398.iat[8, 6]
sima_omid_duration_Mordad_1398=EPG_Mordad_1398.iat[9, 6]
sima_ifilm_duration_Mordad_1398=EPG_Mordad_1398.iat[10, 6]
sima_namayesh_duration_Mordad_1398=EPG_Mordad_1398.iat[11, 6]
sima_tamasha_duration_Mordad_1398=EPG_Mordad_1398.iat[12, 6]
sima_mostanad_duration_Mordad_1398=EPG_Mordad_1398.iat[13, 6]
sima_shoma_duration_Mordad_1398=EPG_Mordad_1398.iat[14, 6]
sima_amozesh_duration_Mordad_1398=EPG_Mordad_1398.iat[15, 6]
sima_varzesh_duration_Mordad_1398=EPG_Mordad_1398.iat[16, 6]
sima_nasim_duration_Mordad_1398=EPG_Mordad_1398.iat[17, 6]
sima_qoran_duration_Mordad_1398=EPG_Mordad_1398.iat[18, 6]
sima_salamat_duration_Mordad_1398=EPG_Mordad_1398.iat[19, 6]
sima_irankala_duration_Mordad_1398=EPG_Mordad_1398.iat[20, 6]
sima_alalam_duration_Mordad_1398=EPG_Mordad_1398.iat[21, 6]
sima_alkosar_duration_Mordad_1398=EPG_Mordad_1398.iat[22, 6]
sima_presstv_duration_Mordad_1398=EPG_Mordad_1398.iat[23, 6]
sima_sepehr_duration_Mordad_1398=EPG_Mordad_1398.iat[24, 6]

sima_lenz_visit_Mordad_1398=EPG_Mordad_1398.iat[36, 2]
sima_aio_visit_Mordad_1398=EPG_Mordad_1398.iat[37, 2]
sima_anten_visit_Mordad_1398=EPG_Mordad_1398.iat[38, 2]
sima_tva_visit_Mordad_1398=EPG_Mordad_1398.iat[39, 2]
sima_fam_visit_Mordad_1398=EPG_Mordad_1398.iat[40, 2]
sima_televebion_visit_Mordad_1398=EPG_Mordad_1398.iat[41, 2]
sima_sepehr_visit_Mordad_1398=EPG_Mordad_1398.iat[42, 2]
#sima_shima_visit_Mordad_1398=EPG_Mordad_1398.iat[40, 2]
#sima_site_visit_Mordad_1398=EPG_Mordad_1398.iat[41, 2]

register_user_lenz_Mordad_1398=EPG_Mordad_1398.iat[36, 4]
register_user_aio_Mordad_1398=EPG_Mordad_1398.iat[37, 4]
register_user_anten_Mordad_1398=EPG_Mordad_1398.iat[38, 4]
register_user_tva_Mordad_1398=EPG_Mordad_1398.iat[39, 4]
register_user_fam_Mordad_1398=EPG_Mordad_1398.iat[40, 4]
register_user_televebion_Mordad_1398=EPG_Mordad_1398.iat[41, 4]
register_user_sepehr_Mordad_1398=EPG_Mordad_1398.iat[42, 4]
#register_user_shima_Mordad_1398=EPG_Mordad_1398.iat[43, 7]
#register_user_site_Mordad_1398=EPG_Mordad_1398.iat[44, 7]

active_user_lenz_Mordad_1398=EPG_Mordad_1398.iat[36, 10]
active_user_aio_Mordad_1398=EPG_Mordad_1398.iat[37, 10]
active_user_anten_Mordad_1398=EPG_Mordad_1398.iat[38, 10]
active_user_tva_Mordad_1398=EPG_Mordad_1398.iat[39, 10]
active_user_fam_Mordad_1398=EPG_Mordad_1398.iat[40, 10]
active_user_televebion_Mordad_1398=EPG_Mordad_1398.iat[41, 10]
active_user_sepehr_Mordad_1398=EPG_Mordad_1398.iat[42, 10]
#active_user_shima_Mordad_1398=EPG_Mordad_1398.iat[43, 10]
#active_user_site_Mordad_1398=EPG_Mordad_1398.iat[44, 10]

all_visit_Mordad_1398=EPG_Mordad_1398.iat[25, 4]
all_duration_Mordad_1398=sum(EPG_Mordad_1398.iloc[1:24, 6])
all_content_sima_Mordad_1398=EPG_Mordad_1398.iat[25, 2]
all_register_user_Mordad_1398=sum(EPG_Mordad_1398.iloc[36:43, 4])
all_active_user_Mordad_1398=sum(EPG_Mordad_1398.iloc[36:43, 10])

Mordad_1398_sima_visit_channels=pd.DataFrame()
Mordad_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Mordad_1398, sima_2_visit_Mordad_1398, sima_3_visit_Mordad_1398,
                 sima_4_visit_Mordad_1398, sima_5_visit_Mordad_1398, sima_khabar_visit_Mordad_1398,
                 sima_ofogh_visit_Mordad_1398, sima_pooya_visit_Mordad_1398, sima_omid_visit_Mordad_1398,
                 sima_ifilm_visit_Mordad_1398, sima_namayesh_visit_Mordad_1398, sima_tamasha_visit_Mordad_1398,
                 sima_mostanad_visit_Mordad_1398, sima_shoma_visit_Mordad_1398, sima_amozesh_visit_Mordad_1398,
                 sima_varzesh_visit_Mordad_1398, sima_nasim_visit_Mordad_1398, sima_qoran_visit_Mordad_1398,
                 sima_salamat_visit_Mordad_1398, sima_irankala_visit_Mordad_1398, sima_alalam_visit_Mordad_1398,
                 sima_alkosar_visit_Mordad_1398, sima_presstv_visit_Mordad_1398, sima_sepehr_visit_Mordad_1398,],
        'duration': [sima_1_duration_Mordad_1398, sima_2_duration_Mordad_1398, sima_3_duration_Mordad_1398,
                 sima_4_duration_Mordad_1398, sima_5_duration_Mordad_1398, sima_khabar_duration_Mordad_1398,
                 sima_ofogh_duration_Mordad_1398, sima_pooya_duration_Mordad_1398, sima_omid_duration_Mordad_1398,
                 sima_ifilm_duration_Mordad_1398, sima_namayesh_duration_Mordad_1398, sima_tamasha_duration_Mordad_1398,
                 sima_mostanad_duration_Mordad_1398, sima_shoma_duration_Mordad_1398, sima_amozesh_duration_Mordad_1398,
                 sima_varzesh_duration_Mordad_1398, sima_nasim_duration_Mordad_1398, sima_qoran_duration_Mordad_1398,
                 sima_salamat_duration_Mordad_1398, sima_irankala_duration_Mordad_1398, sima_alalam_duration_Mordad_1398,
                 sima_alkosar_duration_Mordad_1398, sima_presstv_duration_Mordad_1398, sima_sepehr_duration_Mordad_1398,],}
Mordad_1398_sima_visit_channels=pd.DataFrame(Mordad_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Mordad_1398_sima_visit_channels=Mordad_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Mordad_1398_operator_data=pd.DataFrame()
Mordad_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر'],
       'visit': [sima_lenz_visit_Mordad_1398, sima_aio_visit_Mordad_1398, sima_anten_visit_Mordad_1398,
                 sima_tva_visit_Mordad_1398, sima_fam_visit_Mordad_1398,sima_televebion_visit_Mordad_1398,
                 sima_sepehr_visit_Mordad_1398,],
       'register': [register_user_lenz_Mordad_1398, register_user_aio_Mordad_1398, register_user_anten_Mordad_1398,
                 register_user_tva_Mordad_1398, register_user_fam_Mordad_1398, register_user_televebion_Mordad_1398,
                 register_user_sepehr_Mordad_1398,],
       'active': [active_user_lenz_Mordad_1398, active_user_aio_Mordad_1398, active_user_anten_Mordad_1398,
                 active_user_tva_Mordad_1398, active_user_fam_Mordad_1398, active_user_televebion_Mordad_1398,
                 active_user_sepehr_Mordad_1398,],}

Mordad_1398_operator_data=pd.DataFrame(Mordad_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])

Mordad_1398_operator_data=Mordad_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Mordad_1398_all_data_summary=pd.DataFrame()
Mordad_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Mordad_1398, all_duration_Mordad_1398,all_content_sima_Mordad_1398, all_register_user_Mordad_1398,],}

Mordad_1398_all_data_summary=pd.DataFrame(Mordad_1398_all_data_summary, columns=['parameters', 'statistics'])

Mordad_1398_all_data_summary=Mordad_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه مرداد 1398.xlsx', engine='xlsxwriter')
Mordad_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Mordad_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Mordad_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه مرداد')
writer.save()

        ########################### شهریور #############################
print("EPG Shahrivar 1398")
EPG_Shahrivar_1398=pd.read_excel('EPG/EPG 1398/EPG Shahrivar 1398.xlsx', sheet_name='آمار')
EPG_Shahrivar_1398.fillna(0, inplace=True)
EPG_Shahrivar_1398['مدت زمان بازدید']=EPG_Shahrivar_1398.iloc[1:26, 9]*60
sima_1_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[1, 4]
sima_2_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[2, 4]
sima_3_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[3, 4]
sima_4_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[4, 4]
sima_5_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[5, 4]
sima_khabar_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[6, 4]
sima_ofogh_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[7, 4]
sima_pooya_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[8, 4]
sima_omid_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[9, 4]
sima_ifilm_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[10, 4]
sima_namayesh_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[11, 4]
sima_tamasha_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[12, 4]
sima_mostanad_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[13, 4]
sima_shoma_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[14, 4]
sima_amozesh_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[15, 4]
sima_varzesh_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[16, 4]
sima_nasim_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[17, 4]
sima_qoran_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[18, 4]
sima_salamat_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[19, 4]
sima_irankala_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[20, 4]
sima_alalam_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[21, 4]
sima_alkosar_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[22, 4]
sima_presstv_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[23, 4]
sima_sepehr_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[24, 4]

sima_1_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[1, 6]
sima_2_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[2, 6]
sima_3_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[3, 6]
sima_4_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[4, 6]
sima_5_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[5, 6]
sima_khabar_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[6, 6]
sima_ofogh_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[7, 6]
sima_pooya_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[8, 6]
sima_omid_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[9, 6]
sima_ifilm_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[10, 6]
sima_namayesh_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[11, 6]
sima_tamasha_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[12, 6]
sima_mostanad_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[13, 6]
sima_shoma_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[14, 6]
sima_amozesh_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[15, 6]
sima_varzesh_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[16, 6]
sima_nasim_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[17, 6]
sima_qoran_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[18, 6]
sima_salamat_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[19, 6]
sima_irankala_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[20, 6]
sima_alalam_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[21, 6]
sima_alkosar_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[22, 6]
sima_presstv_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[23, 6]
sima_sepehr_duration_Shahrivar_1398=EPG_Shahrivar_1398.iat[24, 6]

sima_lenz_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[36, 2]
sima_aio_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[37, 2]
sima_anten_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[38, 2]
sima_tva_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[39, 2]
sima_fam_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[40, 2]
sima_televebion_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[41, 2]
sima_sepehr_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[42, 2]
sima_shima_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[43, 2]
#sima_site_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[41, 2]

register_user_lenz_Shahrivar_1398=EPG_Shahrivar_1398.iat[36, 4]
register_user_aio_Shahrivar_1398=EPG_Shahrivar_1398.iat[37, 4]
register_user_anten_Shahrivar_1398=EPG_Shahrivar_1398.iat[38, 4]
register_user_tva_Shahrivar_1398=EPG_Shahrivar_1398.iat[39, 4]
register_user_fam_Shahrivar_1398=EPG_Shahrivar_1398.iat[40, 4]
register_user_televebion_Shahrivar_1398=EPG_Shahrivar_1398.iat[41, 4]
register_user_sepehr_Shahrivar_1398=EPG_Shahrivar_1398.iat[42, 4]
register_user_shima_Shahrivar_1398=EPG_Shahrivar_1398.iat[43, 4]
#register_user_site_Shahrivar_1398=EPG_Shahrivar_1398.iat[44, 7]

active_user_lenz_Shahrivar_1398=EPG_Shahrivar_1398.iat[36, 10]
active_user_aio_Shahrivar_1398=EPG_Shahrivar_1398.iat[37, 10]
active_user_anten_Shahrivar_1398=EPG_Shahrivar_1398.iat[38, 10]
active_user_tva_Shahrivar_1398=EPG_Shahrivar_1398.iat[39, 10]
active_user_fam_Shahrivar_1398=EPG_Shahrivar_1398.iat[40, 10]
active_user_televebion_Shahrivar_1398=EPG_Shahrivar_1398.iat[41, 10]
active_user_sepehr_Shahrivar_1398=EPG_Shahrivar_1398.iat[42, 10]
active_user_shima_Shahrivar_1398=EPG_Shahrivar_1398.iat[43, 10]
#active_user_site_Shahrivar_1398=EPG_Shahrivar_1398.iat[44, 10]

all_visit_Shahrivar_1398=EPG_Shahrivar_1398.iat[25, 4]
all_duration_Shahrivar_1398=sum(EPG_Shahrivar_1398.iloc[1:24, 6])
all_content_sima_Shahrivar_1398=EPG_Shahrivar_1398.iat[25, 2]
all_register_user_Shahrivar_1398=sum(EPG_Shahrivar_1398.iloc[36:44, 4])
all_active_user_Shahrivar_1398=sum(EPG_Shahrivar_1398.iloc[36:44, 10])

Shahrivar_1398_sima_visit_channels=pd.DataFrame()
Shahrivar_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Shahrivar_1398, sima_2_visit_Shahrivar_1398, sima_3_visit_Shahrivar_1398,
                 sima_4_visit_Shahrivar_1398, sima_5_visit_Shahrivar_1398, sima_khabar_visit_Shahrivar_1398,
                 sima_ofogh_visit_Shahrivar_1398, sima_pooya_visit_Shahrivar_1398, sima_omid_visit_Shahrivar_1398,
                 sima_ifilm_visit_Shahrivar_1398, sima_namayesh_visit_Shahrivar_1398, sima_tamasha_visit_Shahrivar_1398,
                 sima_mostanad_visit_Shahrivar_1398, sima_shoma_visit_Shahrivar_1398, sima_amozesh_visit_Shahrivar_1398,
                 sima_varzesh_visit_Shahrivar_1398, sima_nasim_visit_Shahrivar_1398, sima_qoran_visit_Shahrivar_1398,
                 sima_salamat_visit_Shahrivar_1398, sima_irankala_visit_Shahrivar_1398, sima_alalam_visit_Shahrivar_1398,
                 sima_alkosar_visit_Shahrivar_1398, sima_presstv_visit_Shahrivar_1398, sima_sepehr_visit_Shahrivar_1398,],
        'duration': [sima_1_duration_Shahrivar_1398, sima_2_duration_Shahrivar_1398, sima_3_duration_Shahrivar_1398,
                 sima_4_duration_Shahrivar_1398, sima_5_duration_Shahrivar_1398, sima_khabar_duration_Shahrivar_1398,
                 sima_ofogh_duration_Shahrivar_1398, sima_pooya_duration_Shahrivar_1398, sima_omid_duration_Shahrivar_1398,
                 sima_ifilm_duration_Shahrivar_1398, sima_namayesh_duration_Shahrivar_1398, sima_tamasha_duration_Shahrivar_1398,
                 sima_mostanad_duration_Shahrivar_1398, sima_shoma_duration_Shahrivar_1398, sima_amozesh_duration_Shahrivar_1398,
                 sima_varzesh_duration_Shahrivar_1398, sima_nasim_duration_Shahrivar_1398, sima_qoran_duration_Shahrivar_1398,
                 sima_salamat_duration_Shahrivar_1398, sima_irankala_duration_Shahrivar_1398, sima_alalam_duration_Shahrivar_1398,
                 sima_alkosar_duration_Shahrivar_1398, sima_presstv_duration_Shahrivar_1398, sima_sepehr_duration_Shahrivar_1398,],}
Shahrivar_1398_sima_visit_channels=pd.DataFrame(Shahrivar_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Shahrivar_1398_sima_visit_channels=Shahrivar_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Shahrivar_1398_operator_data=pd.DataFrame()
Shahrivar_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما'],
       'visit': [sima_lenz_visit_Shahrivar_1398, sima_aio_visit_Shahrivar_1398, sima_anten_visit_Shahrivar_1398,
                 sima_tva_visit_Shahrivar_1398, sima_fam_visit_Shahrivar_1398,sima_televebion_visit_Shahrivar_1398,
                 sima_sepehr_visit_Shahrivar_1398, sima_shima_visit_Shahrivar_1398,],
       'register': [register_user_lenz_Shahrivar_1398, register_user_aio_Shahrivar_1398, register_user_anten_Shahrivar_1398,
                 register_user_tva_Shahrivar_1398, register_user_fam_Shahrivar_1398, register_user_televebion_Shahrivar_1398,
                 register_user_sepehr_Shahrivar_1398, register_user_shima_Shahrivar_1398,],
       'active': [active_user_lenz_Shahrivar_1398, active_user_aio_Shahrivar_1398, active_user_anten_Shahrivar_1398,
                 active_user_tva_Shahrivar_1398, active_user_fam_Shahrivar_1398, active_user_televebion_Shahrivar_1398,
                 active_user_sepehr_Shahrivar_1398, active_user_shima_Shahrivar_1398,],}

Shahrivar_1398_operator_data=pd.DataFrame(Shahrivar_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])

Shahrivar_1398_operator_data=Shahrivar_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Shahrivar_1398_all_data_summary=pd.DataFrame()
Shahrivar_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Shahrivar_1398, all_duration_Shahrivar_1398,all_content_sima_Shahrivar_1398, all_register_user_Shahrivar_1398,],}

Shahrivar_1398_all_data_summary=pd.DataFrame(Shahrivar_1398_all_data_summary, columns=['parameters', 'statistics'])

Shahrivar_1398_all_data_summary=Shahrivar_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه شهریور 1398.xlsx', engine='xlsxwriter')
Shahrivar_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Shahrivar_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Shahrivar_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه شهریور')
writer.save()

        ########################### مهر #############################
print("EPG Mehr 1398")
EPG_Mehr_1398=pd.read_excel('EPG/EPG 1398/EPG Mehr 1398.xlsx', sheet_name='آمار')
EPG_Mehr_1398.fillna(0, inplace=True)
EPG_Mehr_1398['مدت زمان بازدید']=EPG_Mehr_1398.iloc[1:26, 9]*60
sima_1_visit_Mehr_1398=EPG_Mehr_1398.iat[1, 4]
sima_2_visit_Mehr_1398=EPG_Mehr_1398.iat[2, 4]
sima_3_visit_Mehr_1398=EPG_Mehr_1398.iat[3, 4]
sima_4_visit_Mehr_1398=EPG_Mehr_1398.iat[4, 4]
sima_5_visit_Mehr_1398=EPG_Mehr_1398.iat[5, 4]
sima_khabar_visit_Mehr_1398=EPG_Mehr_1398.iat[6, 4]
sima_ofogh_visit_Mehr_1398=EPG_Mehr_1398.iat[7, 4]
sima_pooya_visit_Mehr_1398=EPG_Mehr_1398.iat[8, 4]
sima_omid_visit_Mehr_1398=EPG_Mehr_1398.iat[9, 4]
sima_ifilm_visit_Mehr_1398=EPG_Mehr_1398.iat[10, 4]
sima_namayesh_visit_Mehr_1398=EPG_Mehr_1398.iat[11, 4]
sima_tamasha_visit_Mehr_1398=EPG_Mehr_1398.iat[12, 4]
sima_mostanad_visit_Mehr_1398=EPG_Mehr_1398.iat[13, 4]
sima_shoma_visit_Mehr_1398=EPG_Mehr_1398.iat[14, 4]
sima_amozesh_visit_Mehr_1398=EPG_Mehr_1398.iat[15, 4]
sima_varzesh_visit_Mehr_1398=EPG_Mehr_1398.iat[16, 4]
sima_nasim_visit_Mehr_1398=EPG_Mehr_1398.iat[17, 4]
sima_qoran_visit_Mehr_1398=EPG_Mehr_1398.iat[18, 4]
sima_salamat_visit_Mehr_1398=EPG_Mehr_1398.iat[19, 4]
sima_irankala_visit_Mehr_1398=EPG_Mehr_1398.iat[20, 4]
sima_alalam_visit_Mehr_1398=EPG_Mehr_1398.iat[21, 4]
sima_alkosar_visit_Mehr_1398=EPG_Mehr_1398.iat[22, 4]
sima_presstv_visit_Mehr_1398=EPG_Mehr_1398.iat[23, 4]
sima_sepehr_visit_Mehr_1398=EPG_Mehr_1398.iat[24, 4]

sima_1_duration_Mehr_1398=EPG_Mehr_1398.iat[1, 6]
sima_2_duration_Mehr_1398=EPG_Mehr_1398.iat[2, 6]
sima_3_duration_Mehr_1398=EPG_Mehr_1398.iat[3, 6]
sima_4_duration_Mehr_1398=EPG_Mehr_1398.iat[4, 6]
sima_5_duration_Mehr_1398=EPG_Mehr_1398.iat[5, 6]
sima_khabar_duration_Mehr_1398=EPG_Mehr_1398.iat[6, 6]
sima_ofogh_duration_Mehr_1398=EPG_Mehr_1398.iat[7, 6]
sima_pooya_duration_Mehr_1398=EPG_Mehr_1398.iat[8, 6]
sima_omid_duration_Mehr_1398=EPG_Mehr_1398.iat[9, 6]
sima_ifilm_duration_Mehr_1398=EPG_Mehr_1398.iat[10, 6]
sima_namayesh_duration_Mehr_1398=EPG_Mehr_1398.iat[11, 6]
sima_tamasha_duration_Mehr_1398=EPG_Mehr_1398.iat[12, 6]
sima_mostanad_duration_Mehr_1398=EPG_Mehr_1398.iat[13, 6]
sima_shoma_duration_Mehr_1398=EPG_Mehr_1398.iat[14, 6]
sima_amozesh_duration_Mehr_1398=EPG_Mehr_1398.iat[15, 6]
sima_varzesh_duration_Mehr_1398=EPG_Mehr_1398.iat[16, 6]
sima_nasim_duration_Mehr_1398=EPG_Mehr_1398.iat[17, 6]
sima_qoran_duration_Mehr_1398=EPG_Mehr_1398.iat[18, 6]
sima_salamat_duration_Mehr_1398=EPG_Mehr_1398.iat[19, 6]
sima_irankala_duration_Mehr_1398=EPG_Mehr_1398.iat[20, 6]
sima_alalam_duration_Mehr_1398=EPG_Mehr_1398.iat[21, 6]
sima_alkosar_duration_Mehr_1398=EPG_Mehr_1398.iat[22, 6]
sima_presstv_duration_Mehr_1398=EPG_Mehr_1398.iat[23, 6]
sima_sepehr_duration_Mehr_1398=EPG_Mehr_1398.iat[24, 6]

sima_lenz_visit_Mehr_1398=EPG_Mehr_1398.iat[36, 2]
sima_aio_visit_Mehr_1398=EPG_Mehr_1398.iat[37, 2]
sima_anten_visit_Mehr_1398=EPG_Mehr_1398.iat[38, 2]
sima_tva_visit_Mehr_1398=EPG_Mehr_1398.iat[39, 2]
sima_fam_visit_Mehr_1398=EPG_Mehr_1398.iat[40, 2]
sima_televebion_visit_Mehr_1398=EPG_Mehr_1398.iat[41, 2]
sima_sepehr_visit_Mehr_1398=EPG_Mehr_1398.iat[42, 2]
sima_shima_visit_Mehr_1398=EPG_Mehr_1398.iat[43, 2]
sima_site_visit_Mehr_1398=EPG_Mehr_1398.iat[44, 2]

register_user_lenz_Mehr_1398=EPG_Mehr_1398.iat[36, 4]
register_user_aio_Mehr_1398=EPG_Mehr_1398.iat[37, 4]
register_user_anten_Mehr_1398=EPG_Mehr_1398.iat[38, 4]
register_user_tva_Mehr_1398=EPG_Mehr_1398.iat[39, 4]
register_user_fam_Mehr_1398=EPG_Mehr_1398.iat[40, 4]
register_user_televebion_Mehr_1398=EPG_Mehr_1398.iat[41, 4]
register_user_sepehr_Mehr_1398=EPG_Mehr_1398.iat[42, 4]
register_user_shima_Mehr_1398=EPG_Mehr_1398.iat[43, 4]
register_user_site_Mehr_1398=EPG_Mehr_1398.iat[44, 4]

active_user_lenz_Mehr_1398=EPG_Mehr_1398.iat[36, 10]
active_user_aio_Mehr_1398=EPG_Mehr_1398.iat[37, 10]
active_user_anten_Mehr_1398=EPG_Mehr_1398.iat[38, 10]
active_user_tva_Mehr_1398=EPG_Mehr_1398.iat[39, 10]
active_user_fam_Mehr_1398=EPG_Mehr_1398.iat[40, 10]
active_user_televebion_Mehr_1398=EPG_Mehr_1398.iat[41, 10]
active_user_sepehr_Mehr_1398=EPG_Mehr_1398.iat[42, 10]
active_user_shima_Mehr_1398=EPG_Mehr_1398.iat[43, 10]
active_user_site_Mehr_1398=EPG_Mehr_1398.iat[44, 10]

all_visit_Mehr_1398=EPG_Mehr_1398.iat[25, 4]
all_duration_Mehr_1398=sum(EPG_Mehr_1398.iloc[1:24, 6])
all_content_sima_Mehr_1398=EPG_Mehr_1398.iat[25, 2]
all_register_user_Mehr_1398=sum(EPG_Mehr_1398.iloc[36:45, 4])
all_active_user_Mehr_1398=sum(EPG_Mehr_1398.iloc[36:45, 10])

Mehr_1398_sima_visit_channels=pd.DataFrame()
Mehr_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Mehr_1398, sima_2_visit_Mehr_1398, sima_3_visit_Mehr_1398,
                 sima_4_visit_Mehr_1398, sima_5_visit_Mehr_1398, sima_khabar_visit_Mehr_1398,
                 sima_ofogh_visit_Mehr_1398, sima_pooya_visit_Mehr_1398, sima_omid_visit_Mehr_1398,
                 sima_ifilm_visit_Mehr_1398, sima_namayesh_visit_Mehr_1398, sima_tamasha_visit_Mehr_1398,
                 sima_mostanad_visit_Mehr_1398, sima_shoma_visit_Mehr_1398, sima_amozesh_visit_Mehr_1398,
                 sima_varzesh_visit_Mehr_1398, sima_nasim_visit_Mehr_1398, sima_qoran_visit_Mehr_1398,
                 sima_salamat_visit_Mehr_1398, sima_irankala_visit_Mehr_1398, sima_alalam_visit_Mehr_1398,
                 sima_alkosar_visit_Mehr_1398, sima_presstv_visit_Mehr_1398, sima_sepehr_visit_Mehr_1398,],
        'duration': [sima_1_duration_Mehr_1398, sima_2_duration_Mehr_1398, sima_3_duration_Mehr_1398,
                 sima_4_duration_Mehr_1398, sima_5_duration_Mehr_1398, sima_khabar_duration_Mehr_1398,
                 sima_ofogh_duration_Mehr_1398, sima_pooya_duration_Mehr_1398, sima_omid_duration_Mehr_1398,
                 sima_ifilm_duration_Mehr_1398, sima_namayesh_duration_Mehr_1398, sima_tamasha_duration_Mehr_1398,
                 sima_mostanad_duration_Mehr_1398, sima_shoma_duration_Mehr_1398, sima_amozesh_duration_Mehr_1398,
                 sima_varzesh_duration_Mehr_1398, sima_nasim_duration_Mehr_1398, sima_qoran_duration_Mehr_1398,
                 sima_salamat_duration_Mehr_1398, sima_irankala_duration_Mehr_1398, sima_alalam_duration_Mehr_1398,
                 sima_alkosar_duration_Mehr_1398, sima_presstv_duration_Mehr_1398, sima_sepehr_duration_Mehr_1398,],}
Mehr_1398_sima_visit_channels=pd.DataFrame(Mehr_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Mehr_1398_sima_visit_channels=Mehr_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Mehr_1398_operator_data=pd.DataFrame()
Mehr_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها'],
       'visit': [sima_lenz_visit_Mehr_1398, sima_aio_visit_Mehr_1398, sima_anten_visit_Mehr_1398,
                 sima_tva_visit_Mehr_1398, sima_fam_visit_Mehr_1398,sima_televebion_visit_Mehr_1398,
                 sima_sepehr_visit_Mehr_1398, sima_shima_visit_Mehr_1398, sima_site_visit_Mehr_1398,],
       'register': [register_user_lenz_Mehr_1398, register_user_aio_Mehr_1398, register_user_anten_Mehr_1398,
                 register_user_tva_Mehr_1398, register_user_fam_Mehr_1398, register_user_televebion_Mehr_1398,
                 register_user_sepehr_Mehr_1398, register_user_shima_Mehr_1398, register_user_site_Mehr_1398,],
       'active': [active_user_lenz_Mehr_1398, active_user_aio_Mehr_1398, active_user_anten_Mehr_1398,
                 active_user_tva_Mehr_1398, active_user_fam_Mehr_1398, active_user_televebion_Mehr_1398,
                 active_user_sepehr_Mehr_1398, active_user_shima_Mehr_1398, register_user_site_Mehr_1398,],}

Mehr_1398_operator_data=pd.DataFrame(Mehr_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])

Mehr_1398_operator_data=Mehr_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Mehr_1398_all_data_summary=pd.DataFrame()
Mehr_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Mehr_1398, all_duration_Mehr_1398,all_content_sima_Mehr_1398, all_register_user_Mehr_1398,],}

Mehr_1398_all_data_summary=pd.DataFrame(Mehr_1398_all_data_summary, columns=['parameters', 'statistics'])

Mehr_1398_all_data_summary=Mehr_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه مهر 1398.xlsx', engine='xlsxwriter')
Mehr_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Mehr_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Mehr_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه مهر')
writer.save()

        ########################### آبان #############################
print("EPG Aban 1398")
EPG_Aban_1398=pd.read_excel('EPG/EPG 1398/EPG Aban 1398.xlsx', sheet_name='آمار')
EPG_Aban_1398.fillna(0, inplace=True)
EPG_Aban_1398['مدت زمان بازدید']=EPG_Aban_1398.iloc[1:26, 9]*60
sima_1_visit_Aban_1398=EPG_Aban_1398.iat[1, 4]
sima_2_visit_Aban_1398=EPG_Aban_1398.iat[2, 4]
sima_3_visit_Aban_1398=EPG_Aban_1398.iat[3, 4]
sima_4_visit_Aban_1398=EPG_Aban_1398.iat[4, 4]
sima_5_visit_Aban_1398=EPG_Aban_1398.iat[5, 4]
sima_khabar_visit_Aban_1398=EPG_Aban_1398.iat[6, 4]
sima_ofogh_visit_Aban_1398=EPG_Aban_1398.iat[7, 4]
sima_pooya_visit_Aban_1398=EPG_Aban_1398.iat[8, 4]
sima_omid_visit_Aban_1398=EPG_Aban_1398.iat[9, 4]
sima_ifilm_visit_Aban_1398=EPG_Aban_1398.iat[10, 4]
sima_namayesh_visit_Aban_1398=EPG_Aban_1398.iat[11, 4]
sima_tamasha_visit_Aban_1398=EPG_Aban_1398.iat[12, 4]
sima_mostanad_visit_Aban_1398=EPG_Aban_1398.iat[13, 4]
sima_shoma_visit_Aban_1398=EPG_Aban_1398.iat[14, 4]
sima_amozesh_visit_Aban_1398=EPG_Aban_1398.iat[15, 4]
sima_varzesh_visit_Aban_1398=EPG_Aban_1398.iat[16, 4]
sima_nasim_visit_Aban_1398=EPG_Aban_1398.iat[17, 4]
sima_qoran_visit_Aban_1398=EPG_Aban_1398.iat[18, 4]
sima_salamat_visit_Aban_1398=EPG_Aban_1398.iat[19, 4]
sima_irankala_visit_Aban_1398=EPG_Aban_1398.iat[20, 4]
sima_alalam_visit_Aban_1398=EPG_Aban_1398.iat[21, 4]
sima_alkosar_visit_Aban_1398=EPG_Aban_1398.iat[22, 4]
sima_presstv_visit_Aban_1398=EPG_Aban_1398.iat[23, 4]
sima_sepehr_visit_Aban_1398=EPG_Aban_1398.iat[24, 4]

sima_1_duration_Aban_1398=EPG_Aban_1398.iat[1, 6]
sima_2_duration_Aban_1398=EPG_Aban_1398.iat[2, 6]
sima_3_duration_Aban_1398=EPG_Aban_1398.iat[3, 6]
sima_4_duration_Aban_1398=EPG_Aban_1398.iat[4, 6]
sima_5_duration_Aban_1398=EPG_Aban_1398.iat[5, 6]
sima_khabar_duration_Aban_1398=EPG_Aban_1398.iat[6, 6]
sima_ofogh_duration_Aban_1398=EPG_Aban_1398.iat[7, 6]
sima_pooya_duration_Aban_1398=EPG_Aban_1398.iat[8, 6]
sima_omid_duration_Aban_1398=EPG_Aban_1398.iat[9, 6]
sima_ifilm_duration_Aban_1398=EPG_Aban_1398.iat[10, 6]
sima_namayesh_duration_Aban_1398=EPG_Aban_1398.iat[11, 6]
sima_tamasha_duration_Aban_1398=EPG_Aban_1398.iat[12, 6]
sima_mostanad_duration_Aban_1398=EPG_Aban_1398.iat[13, 6]
sima_shoma_duration_Aban_1398=EPG_Aban_1398.iat[14, 6]
sima_amozesh_duration_Aban_1398=EPG_Aban_1398.iat[15, 6]
sima_varzesh_duration_Aban_1398=EPG_Aban_1398.iat[16, 6]
sima_nasim_duration_Aban_1398=EPG_Aban_1398.iat[17, 6]
sima_qoran_duration_Aban_1398=EPG_Aban_1398.iat[18, 6]
sima_salamat_duration_Aban_1398=EPG_Aban_1398.iat[19, 6]
sima_irankala_duration_Aban_1398=EPG_Aban_1398.iat[20, 6]
sima_alalam_duration_Aban_1398=EPG_Aban_1398.iat[21, 6]
sima_alkosar_duration_Aban_1398=EPG_Aban_1398.iat[22, 6]
sima_presstv_duration_Aban_1398=EPG_Aban_1398.iat[23, 6]
sima_sepehr_duration_Aban_1398=EPG_Aban_1398.iat[24, 6]

sima_lenz_visit_Aban_1398=EPG_Aban_1398.iat[36, 2]
sima_aio_visit_Aban_1398=EPG_Aban_1398.iat[37, 2]
sima_anten_visit_Aban_1398=EPG_Aban_1398.iat[38, 2]
sima_tva_visit_Aban_1398=EPG_Aban_1398.iat[39, 2]
sima_fam_visit_Aban_1398=EPG_Aban_1398.iat[40, 2]
sima_televebion_visit_Aban_1398=EPG_Aban_1398.iat[41, 2]
sima_sepehr_visit_Aban_1398=EPG_Aban_1398.iat[42, 2]
sima_shima_visit_Aban_1398=EPG_Aban_1398.iat[43, 2]
sima_site_visit_Aban_1398=EPG_Aban_1398.iat[44, 2]

register_user_lenz_Aban_1398=EPG_Aban_1398.iat[36, 4]
register_user_aio_Aban_1398=EPG_Aban_1398.iat[37, 4]
register_user_anten_Aban_1398=EPG_Aban_1398.iat[38, 4]
register_user_tva_Aban_1398=EPG_Aban_1398.iat[39, 4]
register_user_fam_Aban_1398=EPG_Aban_1398.iat[40, 4]
register_user_televebion_Aban_1398=EPG_Aban_1398.iat[41, 4]
register_user_sepehr_Aban_1398=EPG_Aban_1398.iat[42, 4]
register_user_shima_Aban_1398=EPG_Aban_1398.iat[43, 4]
register_user_site_Aban_1398=EPG_Aban_1398.iat[44, 4]

active_user_lenz_Aban_1398=EPG_Aban_1398.iat[36, 10]
active_user_aio_Aban_1398=EPG_Aban_1398.iat[37, 10]
active_user_anten_Aban_1398=EPG_Aban_1398.iat[38, 10]
active_user_tva_Aban_1398=EPG_Aban_1398.iat[39, 10]
active_user_fam_Aban_1398=EPG_Aban_1398.iat[40, 10]
active_user_televebion_Aban_1398=EPG_Aban_1398.iat[41, 10]
active_user_sepehr_Aban_1398=EPG_Aban_1398.iat[42, 10]
active_user_shima_Aban_1398=EPG_Aban_1398.iat[43, 10]
active_user_site_Aban_1398=EPG_Aban_1398.iat[44, 10]

all_visit_Aban_1398=EPG_Aban_1398.iat[25, 4]
all_duration_Aban_1398=sum(EPG_Aban_1398.iloc[1:24, 6])
all_content_sima_Aban_1398=EPG_Aban_1398.iat[25, 2]
all_register_user_Aban_1398=sum(EPG_Aban_1398.iloc[36:45, 4])
all_active_user_Aban_1398=sum(EPG_Aban_1398.iloc[36:45, 10])

Aban_1398_sima_visit_channels=pd.DataFrame()
Aban_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Aban_1398, sima_2_visit_Aban_1398, sima_3_visit_Aban_1398,
                 sima_4_visit_Aban_1398, sima_5_visit_Aban_1398, sima_khabar_visit_Aban_1398,
                 sima_ofogh_visit_Aban_1398, sima_pooya_visit_Aban_1398, sima_omid_visit_Aban_1398,
                 sima_ifilm_visit_Aban_1398, sima_namayesh_visit_Aban_1398, sima_tamasha_visit_Aban_1398,
                 sima_mostanad_visit_Aban_1398, sima_shoma_visit_Aban_1398, sima_amozesh_visit_Aban_1398,
                 sima_varzesh_visit_Aban_1398, sima_nasim_visit_Aban_1398, sima_qoran_visit_Aban_1398,
                 sima_salamat_visit_Aban_1398, sima_irankala_visit_Aban_1398, sima_alalam_visit_Aban_1398,
                 sima_alkosar_visit_Aban_1398, sima_presstv_visit_Aban_1398, sima_sepehr_visit_Aban_1398,],
        'duration': [sima_1_duration_Aban_1398, sima_2_duration_Aban_1398, sima_3_duration_Aban_1398,
                 sima_4_duration_Aban_1398, sima_5_duration_Aban_1398, sima_khabar_duration_Aban_1398,
                 sima_ofogh_duration_Aban_1398, sima_pooya_duration_Aban_1398, sima_omid_duration_Aban_1398,
                 sima_ifilm_duration_Aban_1398, sima_namayesh_duration_Aban_1398, sima_tamasha_duration_Aban_1398,
                 sima_mostanad_duration_Aban_1398, sima_shoma_duration_Aban_1398, sima_amozesh_duration_Aban_1398,
                 sima_varzesh_duration_Aban_1398, sima_nasim_duration_Aban_1398, sima_qoran_duration_Aban_1398,
                 sima_salamat_duration_Aban_1398, sima_irankala_duration_Aban_1398, sima_alalam_duration_Aban_1398,
                 sima_alkosar_duration_Aban_1398, sima_presstv_duration_Aban_1398, sima_sepehr_duration_Aban_1398,],}
Aban_1398_sima_visit_channels=pd.DataFrame(Aban_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Aban_1398_sima_visit_channels=Aban_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Aban_1398_operator_data=pd.DataFrame()
Aban_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها'],
       'visit': [sima_lenz_visit_Aban_1398, sima_aio_visit_Aban_1398, sima_anten_visit_Aban_1398,
                 sima_tva_visit_Aban_1398, sima_fam_visit_Aban_1398,sima_televebion_visit_Aban_1398,
                 sima_sepehr_visit_Aban_1398, sima_shima_visit_Aban_1398, sima_site_visit_Aban_1398,],
       'register': [register_user_lenz_Aban_1398, register_user_aio_Aban_1398, register_user_anten_Aban_1398,
                 register_user_tva_Aban_1398, register_user_fam_Aban_1398, register_user_televebion_Aban_1398,
                 register_user_sepehr_Aban_1398, register_user_shima_Aban_1398, register_user_site_Aban_1398,],
       'active': [active_user_lenz_Aban_1398, active_user_aio_Aban_1398, active_user_anten_Aban_1398,
                 active_user_tva_Aban_1398, active_user_fam_Aban_1398, active_user_televebion_Aban_1398,
                 active_user_sepehr_Aban_1398, active_user_shima_Aban_1398, register_user_site_Aban_1398,],}

Aban_1398_operator_data=pd.DataFrame(Aban_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])

Aban_1398_operator_data=Aban_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Aban_1398_all_data_summary=pd.DataFrame()
Aban_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Aban_1398, all_duration_Aban_1398,all_content_sima_Aban_1398, all_register_user_Aban_1398,],}

Aban_1398_all_data_summary=pd.DataFrame(Aban_1398_all_data_summary, columns=['parameters', 'statistics'])

Aban_1398_all_data_summary=Aban_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه آبان 1398.xlsx', engine='xlsxwriter')
Aban_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Aban_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Aban_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه آبان')
writer.save()

        ########################### آذر #############################
print("EPG Azar 1398")
EPG_Azar_1398=pd.read_excel('EPG/EPG 1398/EPG Azar 1398.xlsx', sheet_name='آمار')
EPG_Azar_1398.fillna(0, inplace=True)
EPG_Azar_1398['مدت زمان بازدید']=EPG_Azar_1398.iloc[1:26, 9]*60
sima_1_visit_Azar_1398=EPG_Azar_1398.iat[1, 4]
sima_2_visit_Azar_1398=EPG_Azar_1398.iat[2, 4]
sima_3_visit_Azar_1398=EPG_Azar_1398.iat[3, 4]
sima_4_visit_Azar_1398=EPG_Azar_1398.iat[4, 4]
sima_5_visit_Azar_1398=EPG_Azar_1398.iat[5, 4]
sima_khabar_visit_Azar_1398=EPG_Azar_1398.iat[6, 4]
sima_ofogh_visit_Azar_1398=EPG_Azar_1398.iat[7, 4]
sima_pooya_visit_Azar_1398=EPG_Azar_1398.iat[8, 4]
sima_omid_visit_Azar_1398=EPG_Azar_1398.iat[9, 4]
sima_ifilm_visit_Azar_1398=EPG_Azar_1398.iat[10, 4]
sima_namayesh_visit_Azar_1398=EPG_Azar_1398.iat[11, 4]
sima_tamasha_visit_Azar_1398=EPG_Azar_1398.iat[12, 4]
sima_mostanad_visit_Azar_1398=EPG_Azar_1398.iat[13, 4]
sima_shoma_visit_Azar_1398=EPG_Azar_1398.iat[14, 4]
sima_amozesh_visit_Azar_1398=EPG_Azar_1398.iat[15, 4]
sima_varzesh_visit_Azar_1398=EPG_Azar_1398.iat[16, 4]
sima_nasim_visit_Azar_1398=EPG_Azar_1398.iat[17, 4]
sima_qoran_visit_Azar_1398=EPG_Azar_1398.iat[18, 4]
sima_salamat_visit_Azar_1398=EPG_Azar_1398.iat[19, 4]
sima_irankala_visit_Azar_1398=EPG_Azar_1398.iat[20, 4]
sima_alalam_visit_Azar_1398=EPG_Azar_1398.iat[21, 4]
sima_alkosar_visit_Azar_1398=EPG_Azar_1398.iat[22, 4]
sima_presstv_visit_Azar_1398=EPG_Azar_1398.iat[23, 4]
sima_sepehr_visit_Azar_1398=EPG_Azar_1398.iat[24, 4]

sima_1_duration_Azar_1398=EPG_Azar_1398.iat[1, 6]
sima_2_duration_Azar_1398=EPG_Azar_1398.iat[2, 6]
sima_3_duration_Azar_1398=EPG_Azar_1398.iat[3, 6]
sima_4_duration_Azar_1398=EPG_Azar_1398.iat[4, 6]
sima_5_duration_Azar_1398=EPG_Azar_1398.iat[5, 6]
sima_khabar_duration_Azar_1398=EPG_Azar_1398.iat[6, 6]
sima_ofogh_duration_Azar_1398=EPG_Azar_1398.iat[7, 6]
sima_pooya_duration_Azar_1398=EPG_Azar_1398.iat[8, 6]
sima_omid_duration_Azar_1398=EPG_Azar_1398.iat[9, 6]
sima_ifilm_duration_Azar_1398=EPG_Azar_1398.iat[10, 6]
sima_namayesh_duration_Azar_1398=EPG_Azar_1398.iat[11, 6]
sima_tamasha_duration_Azar_1398=EPG_Azar_1398.iat[12, 6]
sima_mostanad_duration_Azar_1398=EPG_Azar_1398.iat[13, 6]
sima_shoma_duration_Azar_1398=EPG_Azar_1398.iat[14, 6]
sima_amozesh_duration_Azar_1398=EPG_Azar_1398.iat[15, 6]
sima_varzesh_duration_Azar_1398=EPG_Azar_1398.iat[16, 6]
sima_nasim_duration_Azar_1398=EPG_Azar_1398.iat[17, 6]
sima_qoran_duration_Azar_1398=EPG_Azar_1398.iat[18, 6]
sima_salamat_duration_Azar_1398=EPG_Azar_1398.iat[19, 6]
sima_irankala_duration_Azar_1398=EPG_Azar_1398.iat[20, 6]
sima_alalam_duration_Azar_1398=EPG_Azar_1398.iat[21, 6]
sima_alkosar_duration_Azar_1398=EPG_Azar_1398.iat[22, 6]
sima_presstv_duration_Azar_1398=EPG_Azar_1398.iat[23, 6]
sima_sepehr_duration_Azar_1398=EPG_Azar_1398.iat[24, 6]

sima_lenz_visit_Azar_1398=EPG_Azar_1398.iat[36, 2]
sima_aio_visit_Azar_1398=EPG_Azar_1398.iat[37, 2]
sima_anten_visit_Azar_1398=EPG_Azar_1398.iat[38, 2]
sima_tva_visit_Azar_1398=EPG_Azar_1398.iat[39, 2]
sima_fam_visit_Azar_1398=EPG_Azar_1398.iat[40, 2]
sima_televebion_visit_Azar_1398=EPG_Azar_1398.iat[41, 2]
sima_sepehr_visit_Azar_1398=EPG_Azar_1398.iat[42, 2]
sima_shima_visit_Azar_1398=EPG_Azar_1398.iat[43, 2]
sima_site_visit_Azar_1398=EPG_Azar_1398.iat[44, 2]

register_user_lenz_Azar_1398=EPG_Azar_1398.iat[36, 4]
register_user_aio_Azar_1398=EPG_Azar_1398.iat[37, 4]
register_user_anten_Azar_1398=EPG_Azar_1398.iat[38, 4]
register_user_tva_Azar_1398=EPG_Azar_1398.iat[39, 4]
register_user_fam_Azar_1398=EPG_Azar_1398.iat[40, 4]
register_user_televebion_Azar_1398=EPG_Azar_1398.iat[41, 4]
register_user_sepehr_Azar_1398=EPG_Azar_1398.iat[42, 4]
register_user_shima_Azar_1398=EPG_Azar_1398.iat[43, 4]
register_user_site_Azar_1398=EPG_Azar_1398.iat[44, 4]

active_user_lenz_Azar_1398=EPG_Azar_1398.iat[36, 10]
active_user_aio_Azar_1398=EPG_Azar_1398.iat[37, 10]
active_user_anten_Azar_1398=EPG_Azar_1398.iat[38, 10]
active_user_tva_Azar_1398=EPG_Azar_1398.iat[39, 10]
active_user_fam_Azar_1398=EPG_Azar_1398.iat[40, 10]
active_user_televebion_Azar_1398=EPG_Azar_1398.iat[41, 10]
active_user_sepehr_Azar_1398=EPG_Azar_1398.iat[42, 10]
active_user_shima_Azar_1398=EPG_Azar_1398.iat[43, 10]
active_user_site_Azar_1398=EPG_Azar_1398.iat[44, 10]

all_visit_Azar_1398=EPG_Azar_1398.iat[25, 4]
all_duration_Azar_1398=sum(EPG_Azar_1398.iloc[1:24, 6])
all_content_sima_Azar_1398=EPG_Azar_1398.iat[25, 2]
all_register_user_Azar_1398=sum(EPG_Azar_1398.iloc[36:45, 4])
all_active_user_Azar_1398=sum(EPG_Azar_1398.iloc[36:45, 10])

Azar_1398_sima_visit_channels=pd.DataFrame()
Azar_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Azar_1398, sima_2_visit_Azar_1398, sima_3_visit_Azar_1398,
                 sima_4_visit_Azar_1398, sima_5_visit_Azar_1398, sima_khabar_visit_Azar_1398,
                 sima_ofogh_visit_Azar_1398, sima_pooya_visit_Azar_1398, sima_omid_visit_Azar_1398,
                 sima_ifilm_visit_Azar_1398, sima_namayesh_visit_Azar_1398, sima_tamasha_visit_Azar_1398,
                 sima_mostanad_visit_Azar_1398, sima_shoma_visit_Azar_1398, sima_amozesh_visit_Azar_1398,
                 sima_varzesh_visit_Azar_1398, sima_nasim_visit_Azar_1398, sima_qoran_visit_Azar_1398,
                 sima_salamat_visit_Azar_1398, sima_irankala_visit_Azar_1398, sima_alalam_visit_Azar_1398,
                 sima_alkosar_visit_Azar_1398, sima_presstv_visit_Azar_1398, sima_sepehr_visit_Azar_1398,],
        'duration': [sima_1_duration_Azar_1398, sima_2_duration_Azar_1398, sima_3_duration_Azar_1398,
                 sima_4_duration_Azar_1398, sima_5_duration_Azar_1398, sima_khabar_duration_Azar_1398,
                 sima_ofogh_duration_Azar_1398, sima_pooya_duration_Azar_1398, sima_omid_duration_Azar_1398,
                 sima_ifilm_duration_Azar_1398, sima_namayesh_duration_Azar_1398, sima_tamasha_duration_Azar_1398,
                 sima_mostanad_duration_Azar_1398, sima_shoma_duration_Azar_1398, sima_amozesh_duration_Azar_1398,
                 sima_varzesh_duration_Azar_1398, sima_nasim_duration_Azar_1398, sima_qoran_duration_Azar_1398,
                 sima_salamat_duration_Azar_1398, sima_irankala_duration_Azar_1398, sima_alalam_duration_Azar_1398,
                 sima_alkosar_duration_Azar_1398, sima_presstv_duration_Azar_1398, sima_sepehr_duration_Azar_1398,],}
Azar_1398_sima_visit_channels=pd.DataFrame(Azar_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Azar_1398_sima_visit_channels=Azar_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Azar_1398_operator_data=pd.DataFrame()
Azar_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها'],
       'visit': [sima_lenz_visit_Azar_1398, sima_aio_visit_Azar_1398, sima_anten_visit_Azar_1398,
                 sima_tva_visit_Azar_1398, sima_fam_visit_Azar_1398,sima_televebion_visit_Azar_1398,
                 sima_sepehr_visit_Azar_1398, sima_shima_visit_Azar_1398, sima_site_visit_Azar_1398,],
       'register': [register_user_lenz_Azar_1398, register_user_aio_Azar_1398, register_user_anten_Azar_1398,
                 register_user_tva_Azar_1398, register_user_fam_Azar_1398, register_user_televebion_Azar_1398,
                 register_user_sepehr_Azar_1398, register_user_shima_Azar_1398, register_user_site_Azar_1398,],
       'active': [active_user_lenz_Azar_1398, active_user_aio_Azar_1398, active_user_anten_Azar_1398,
                 active_user_tva_Azar_1398, active_user_fam_Azar_1398, active_user_televebion_Azar_1398,
                 active_user_sepehr_Azar_1398, active_user_shima_Azar_1398, register_user_site_Azar_1398,],}

Azar_1398_operator_data=pd.DataFrame(Azar_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])

Azar_1398_operator_data=Azar_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Azar_1398_all_data_summary=pd.DataFrame()
Azar_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Azar_1398, all_duration_Azar_1398,all_content_sima_Azar_1398, all_register_user_Azar_1398,],}

Azar_1398_all_data_summary=pd.DataFrame(Azar_1398_all_data_summary, columns=['parameters', 'statistics'])

Azar_1398_all_data_summary=Azar_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه آذر 1398.xlsx', engine='xlsxwriter')
Azar_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Azar_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Azar_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه آذر')
writer.save()

        ########################### دی #############################
print("EPG Dey 1398")
EPG_Dey_1398=pd.read_excel('EPG/EPG 1398/EPG Dey 1398.xlsx', sheet_name='آمار')
EPG_Dey_1398.fillna(0, inplace=True)
EPG_Dey_1398['مدت زمان بازدید']=EPG_Dey_1398.iloc[1:26, 9]*60
sima_1_visit_Dey_1398=EPG_Dey_1398.iat[1, 4]
sima_2_visit_Dey_1398=EPG_Dey_1398.iat[2, 4]
sima_3_visit_Dey_1398=EPG_Dey_1398.iat[3, 4]
sima_4_visit_Dey_1398=EPG_Dey_1398.iat[4, 4]
sima_5_visit_Dey_1398=EPG_Dey_1398.iat[5, 4]
sima_khabar_visit_Dey_1398=EPG_Dey_1398.iat[6, 4]
sima_ofogh_visit_Dey_1398=EPG_Dey_1398.iat[7, 4]
sima_pooya_visit_Dey_1398=EPG_Dey_1398.iat[8, 4]
sima_omid_visit_Dey_1398=EPG_Dey_1398.iat[9, 4]
sima_ifilm_visit_Dey_1398=EPG_Dey_1398.iat[10, 4]
sima_namayesh_visit_Dey_1398=EPG_Dey_1398.iat[11, 4]
sima_tamasha_visit_Dey_1398=EPG_Dey_1398.iat[12, 4]
sima_mostanad_visit_Dey_1398=EPG_Dey_1398.iat[13, 4]
sima_shoma_visit_Dey_1398=EPG_Dey_1398.iat[14, 4]
sima_amozesh_visit_Dey_1398=EPG_Dey_1398.iat[15, 4]
sima_varzesh_visit_Dey_1398=EPG_Dey_1398.iat[16, 4]
sima_nasim_visit_Dey_1398=EPG_Dey_1398.iat[17, 4]
sima_qoran_visit_Dey_1398=EPG_Dey_1398.iat[18, 4]
sima_salamat_visit_Dey_1398=EPG_Dey_1398.iat[19, 4]
sima_irankala_visit_Dey_1398=EPG_Dey_1398.iat[20, 4]
sima_alalam_visit_Dey_1398=EPG_Dey_1398.iat[21, 4]
sima_alkosar_visit_Dey_1398=EPG_Dey_1398.iat[22, 4]
sima_presstv_visit_Dey_1398=EPG_Dey_1398.iat[23, 4]
sima_sepehr_visit_Dey_1398=EPG_Dey_1398.iat[24, 4]

sima_1_duration_Dey_1398=EPG_Dey_1398.iat[1, 6]
sima_2_duration_Dey_1398=EPG_Dey_1398.iat[2, 6]
sima_3_duration_Dey_1398=EPG_Dey_1398.iat[3, 6]
sima_4_duration_Dey_1398=EPG_Dey_1398.iat[4, 6]
sima_5_duration_Dey_1398=EPG_Dey_1398.iat[5, 6]
sima_khabar_duration_Dey_1398=EPG_Dey_1398.iat[6, 6]
sima_ofogh_duration_Dey_1398=EPG_Dey_1398.iat[7, 6]
sima_pooya_duration_Dey_1398=EPG_Dey_1398.iat[8, 6]
sima_omid_duration_Dey_1398=EPG_Dey_1398.iat[9, 6]
sima_ifilm_duration_Dey_1398=EPG_Dey_1398.iat[10, 6]
sima_namayesh_duration_Dey_1398=EPG_Dey_1398.iat[11, 6]
sima_tamasha_duration_Dey_1398=EPG_Dey_1398.iat[12, 6]
sima_mostanad_duration_Dey_1398=EPG_Dey_1398.iat[13, 6]
sima_shoma_duration_Dey_1398=EPG_Dey_1398.iat[14, 6]
sima_amozesh_duration_Dey_1398=EPG_Dey_1398.iat[15, 6]
sima_varzesh_duration_Dey_1398=EPG_Dey_1398.iat[16, 6]
sima_nasim_duration_Dey_1398=EPG_Dey_1398.iat[17, 6]
sima_qoran_duration_Dey_1398=EPG_Dey_1398.iat[18, 6]
sima_salamat_duration_Dey_1398=EPG_Dey_1398.iat[19, 6]
sima_irankala_duration_Dey_1398=EPG_Dey_1398.iat[20, 6]
sima_alalam_duration_Dey_1398=EPG_Dey_1398.iat[21, 6]
sima_alkosar_duration_Dey_1398=EPG_Dey_1398.iat[22, 6]
sima_presstv_duration_Dey_1398=EPG_Dey_1398.iat[23, 6]
sima_sepehr_duration_Dey_1398=EPG_Dey_1398.iat[24, 6]

sima_lenz_visit_Dey_1398=EPG_Dey_1398.iat[36, 2]
sima_aio_visit_Dey_1398=EPG_Dey_1398.iat[37, 2]
sima_anten_visit_Dey_1398=EPG_Dey_1398.iat[38, 2]
sima_tva_visit_Dey_1398=EPG_Dey_1398.iat[39, 2]
sima_fam_visit_Dey_1398=EPG_Dey_1398.iat[40, 2]
sima_televebion_visit_Dey_1398=EPG_Dey_1398.iat[41, 2]
sima_sepehr_visit_Dey_1398=EPG_Dey_1398.iat[42, 2]
sima_shima_visit_Dey_1398=EPG_Dey_1398.iat[43, 2]
sima_site_visit_Dey_1398=EPG_Dey_1398.iat[44, 2]

register_user_lenz_Dey_1398=EPG_Dey_1398.iat[36, 4]
register_user_aio_Dey_1398=EPG_Dey_1398.iat[37, 4]
register_user_anten_Dey_1398=EPG_Dey_1398.iat[38, 4]
register_user_tva_Dey_1398=EPG_Dey_1398.iat[39, 4]
register_user_fam_Dey_1398=EPG_Dey_1398.iat[40, 4]
register_user_televebion_Dey_1398=EPG_Dey_1398.iat[41, 4]
register_user_sepehr_Dey_1398=EPG_Dey_1398.iat[42, 4]
register_user_shima_Dey_1398=EPG_Dey_1398.iat[43, 4]
register_user_site_Dey_1398=EPG_Dey_1398.iat[44, 4]

active_user_lenz_Dey_1398=EPG_Dey_1398.iat[36, 10]
active_user_aio_Dey_1398=EPG_Dey_1398.iat[37, 10]
active_user_anten_Dey_1398=EPG_Dey_1398.iat[38, 10]
active_user_tva_Dey_1398=EPG_Dey_1398.iat[39, 10]
active_user_fam_Dey_1398=EPG_Dey_1398.iat[40, 10]
active_user_televebion_Dey_1398=EPG_Dey_1398.iat[41, 10]
active_user_sepehr_Dey_1398=EPG_Dey_1398.iat[42, 10]
active_user_shima_Dey_1398=EPG_Dey_1398.iat[43, 10]
active_user_site_Dey_1398=EPG_Dey_1398.iat[44, 10]

all_visit_Dey_1398=EPG_Dey_1398.iat[25, 4]
all_duration_Dey_1398=sum(EPG_Dey_1398.iloc[1:24, 6])
all_content_sima_Dey_1398=EPG_Dey_1398.iat[25, 2]
all_register_user_Dey_1398=sum(EPG_Dey_1398.iloc[36:45, 4])
all_active_user_Dey_1398=sum(EPG_Dey_1398.iloc[36:45, 10])

Dey_1398_sima_visit_channels=pd.DataFrame()
Dey_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Dey_1398, sima_2_visit_Dey_1398, sima_3_visit_Dey_1398,
                 sima_4_visit_Dey_1398, sima_5_visit_Dey_1398, sima_khabar_visit_Dey_1398,
                 sima_ofogh_visit_Dey_1398, sima_pooya_visit_Dey_1398, sima_omid_visit_Dey_1398,
                 sima_ifilm_visit_Dey_1398, sima_namayesh_visit_Dey_1398, sima_tamasha_visit_Dey_1398,
                 sima_mostanad_visit_Dey_1398, sima_shoma_visit_Dey_1398, sima_amozesh_visit_Dey_1398,
                 sima_varzesh_visit_Dey_1398, sima_nasim_visit_Dey_1398, sima_qoran_visit_Dey_1398,
                 sima_salamat_visit_Dey_1398, sima_irankala_visit_Dey_1398, sima_alalam_visit_Dey_1398,
                 sima_alkosar_visit_Dey_1398, sima_presstv_visit_Dey_1398, sima_sepehr_visit_Dey_1398,],
        'duration': [sima_1_duration_Dey_1398, sima_2_duration_Dey_1398, sima_3_duration_Dey_1398,
                 sima_4_duration_Dey_1398, sima_5_duration_Dey_1398, sima_khabar_duration_Dey_1398,
                 sima_ofogh_duration_Dey_1398, sima_pooya_duration_Dey_1398, sima_omid_duration_Dey_1398,
                 sima_ifilm_duration_Dey_1398, sima_namayesh_duration_Dey_1398, sima_tamasha_duration_Dey_1398,
                 sima_mostanad_duration_Dey_1398, sima_shoma_duration_Dey_1398, sima_amozesh_duration_Dey_1398,
                 sima_varzesh_duration_Dey_1398, sima_nasim_duration_Dey_1398, sima_qoran_duration_Dey_1398,
                 sima_salamat_duration_Dey_1398, sima_irankala_duration_Dey_1398, sima_alalam_duration_Dey_1398,
                 sima_alkosar_duration_Dey_1398, sima_presstv_duration_Dey_1398, sima_sepehr_duration_Dey_1398,],}
Dey_1398_sima_visit_channels=pd.DataFrame(Dey_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Dey_1398_sima_visit_channels=Dey_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Dey_1398_operator_data=pd.DataFrame()
Dey_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها'],
       'visit': [sima_lenz_visit_Dey_1398, sima_aio_visit_Dey_1398, sima_anten_visit_Dey_1398,
                 sima_tva_visit_Dey_1398, sima_fam_visit_Dey_1398,sima_televebion_visit_Dey_1398,
                 sima_sepehr_visit_Dey_1398, sima_shima_visit_Dey_1398, sima_site_visit_Dey_1398,],
       'register': [register_user_lenz_Dey_1398, register_user_aio_Dey_1398, register_user_anten_Dey_1398,
                 register_user_tva_Dey_1398, register_user_fam_Dey_1398, register_user_televebion_Dey_1398,
                 register_user_sepehr_Dey_1398, register_user_shima_Dey_1398, register_user_site_Dey_1398,],
       'active': [active_user_lenz_Dey_1398, active_user_aio_Dey_1398, active_user_anten_Dey_1398,
                 active_user_tva_Dey_1398, active_user_fam_Dey_1398, active_user_televebion_Dey_1398,
                 active_user_sepehr_Dey_1398, active_user_shima_Dey_1398, register_user_site_Dey_1398,],}

Dey_1398_operator_data=pd.DataFrame(Dey_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])

Dey_1398_operator_data=Dey_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Dey_1398_all_data_summary=pd.DataFrame()
Dey_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Dey_1398, all_duration_Dey_1398,all_content_sima_Dey_1398, all_register_user_Dey_1398,],}

Dey_1398_all_data_summary=pd.DataFrame(Dey_1398_all_data_summary, columns=['parameters', 'statistics'])

Dey_1398_all_data_summary=Dey_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه دی 1398.xlsx', engine='xlsxwriter')
Dey_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Dey_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Dey_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه دی')
writer.save()

        ########################### بهمن #############################
print("EPG Bahman 1398")
EPG_Bahman_1398=pd.read_excel('EPG/EPG 1398/EPG Bahman 1398.xlsx', sheet_name='آمار')
EPG_Bahman_1398.fillna(0, inplace=True)
EPG_Bahman_1398['مدت زمان بازدید']=EPG_Bahman_1398.iloc[1:26, 9]*60
sima_1_visit_Bahman_1398=EPG_Bahman_1398.iat[1, 4]
sima_2_visit_Bahman_1398=EPG_Bahman_1398.iat[2, 4]
sima_3_visit_Bahman_1398=EPG_Bahman_1398.iat[3, 4]
sima_4_visit_Bahman_1398=EPG_Bahman_1398.iat[4, 4]
sima_5_visit_Bahman_1398=EPG_Bahman_1398.iat[5, 4]
sima_khabar_visit_Bahman_1398=EPG_Bahman_1398.iat[6, 4]
sima_ofogh_visit_Bahman_1398=EPG_Bahman_1398.iat[7, 4]
sima_pooya_visit_Bahman_1398=EPG_Bahman_1398.iat[8, 4]
sima_omid_visit_Bahman_1398=EPG_Bahman_1398.iat[9, 4]
sima_ifilm_visit_Bahman_1398=EPG_Bahman_1398.iat[10, 4]
sima_namayesh_visit_Bahman_1398=EPG_Bahman_1398.iat[11, 4]
sima_tamasha_visit_Bahman_1398=EPG_Bahman_1398.iat[12, 4]
sima_mostanad_visit_Bahman_1398=EPG_Bahman_1398.iat[13, 4]
sima_shoma_visit_Bahman_1398=EPG_Bahman_1398.iat[14, 4]
sima_amozesh_visit_Bahman_1398=EPG_Bahman_1398.iat[15, 4]
sima_varzesh_visit_Bahman_1398=EPG_Bahman_1398.iat[16, 4]
sima_nasim_visit_Bahman_1398=EPG_Bahman_1398.iat[17, 4]
sima_qoran_visit_Bahman_1398=EPG_Bahman_1398.iat[18, 4]
sima_salamat_visit_Bahman_1398=EPG_Bahman_1398.iat[19, 4]
sima_irankala_visit_Bahman_1398=EPG_Bahman_1398.iat[20, 4]
sima_alalam_visit_Bahman_1398=EPG_Bahman_1398.iat[21, 4]
sima_alkosar_visit_Bahman_1398=EPG_Bahman_1398.iat[22, 4]
sima_presstv_visit_Bahman_1398=EPG_Bahman_1398.iat[23, 4]
sima_sepehr_visit_Bahman_1398=EPG_Bahman_1398.iat[24, 4]

sima_1_duration_Bahman_1398=EPG_Bahman_1398.iat[1, 6]
sima_2_duration_Bahman_1398=EPG_Bahman_1398.iat[2, 6]
sima_3_duration_Bahman_1398=EPG_Bahman_1398.iat[3, 6]
sima_4_duration_Bahman_1398=EPG_Bahman_1398.iat[4, 6]
sima_5_duration_Bahman_1398=EPG_Bahman_1398.iat[5, 6]
sima_khabar_duration_Bahman_1398=EPG_Bahman_1398.iat[6, 6]
sima_ofogh_duration_Bahman_1398=EPG_Bahman_1398.iat[7, 6]
sima_pooya_duration_Bahman_1398=EPG_Bahman_1398.iat[8, 6]
sima_omid_duration_Bahman_1398=EPG_Bahman_1398.iat[9, 6]
sima_ifilm_duration_Bahman_1398=EPG_Bahman_1398.iat[10, 6]
sima_namayesh_duration_Bahman_1398=EPG_Bahman_1398.iat[11, 6]
sima_tamasha_duration_Bahman_1398=EPG_Bahman_1398.iat[12, 6]
sima_mostanad_duration_Bahman_1398=EPG_Bahman_1398.iat[13, 6]
sima_shoma_duration_Bahman_1398=EPG_Bahman_1398.iat[14, 6]
sima_amozesh_duration_Bahman_1398=EPG_Bahman_1398.iat[15, 6]
sima_varzesh_duration_Bahman_1398=EPG_Bahman_1398.iat[16, 6]
sima_nasim_duration_Bahman_1398=EPG_Bahman_1398.iat[17, 6]
sima_qoran_duration_Bahman_1398=EPG_Bahman_1398.iat[18, 6]
sima_salamat_duration_Bahman_1398=EPG_Bahman_1398.iat[19, 6]
sima_irankala_duration_Bahman_1398=EPG_Bahman_1398.iat[20, 6]
sima_alalam_duration_Bahman_1398=EPG_Bahman_1398.iat[21, 6]
sima_alkosar_duration_Bahman_1398=EPG_Bahman_1398.iat[22, 6]
sima_presstv_duration_Bahman_1398=EPG_Bahman_1398.iat[23, 6]
sima_sepehr_duration_Bahman_1398=EPG_Bahman_1398.iat[24, 6]

sima_lenz_visit_Bahman_1398=EPG_Bahman_1398.iat[36, 2]
sima_aio_visit_Bahman_1398=EPG_Bahman_1398.iat[37, 2]
sima_anten_visit_Bahman_1398=EPG_Bahman_1398.iat[38, 2]
sima_tva_visit_Bahman_1398=EPG_Bahman_1398.iat[39, 2]
sima_fam_visit_Bahman_1398=EPG_Bahman_1398.iat[40, 2]
sima_televebion_visit_Bahman_1398=EPG_Bahman_1398.iat[41, 2]
sima_sepehr_visit_Bahman_1398=EPG_Bahman_1398.iat[42, 2]
sima_shima_visit_Bahman_1398=EPG_Bahman_1398.iat[43, 2]
sima_site_visit_Bahman_1398=EPG_Bahman_1398.iat[44, 2]

register_user_lenz_Bahman_1398=EPG_Bahman_1398.iat[36, 4]
register_user_aio_Bahman_1398=EPG_Bahman_1398.iat[37, 4]
register_user_anten_Bahman_1398=EPG_Bahman_1398.iat[38, 4]
register_user_tva_Bahman_1398=EPG_Bahman_1398.iat[39, 4]
register_user_fam_Bahman_1398=EPG_Bahman_1398.iat[40, 4]
register_user_televebion_Bahman_1398=EPG_Bahman_1398.iat[41, 4]
register_user_sepehr_Bahman_1398=EPG_Bahman_1398.iat[42, 4]
register_user_shima_Bahman_1398=EPG_Bahman_1398.iat[43, 4]
register_user_site_Bahman_1398=EPG_Bahman_1398.iat[44, 4]

active_user_lenz_Bahman_1398=EPG_Bahman_1398.iat[36, 10]
active_user_aio_Bahman_1398=EPG_Bahman_1398.iat[37, 10]
active_user_anten_Bahman_1398=EPG_Bahman_1398.iat[38, 10]
active_user_tva_Bahman_1398=EPG_Bahman_1398.iat[39, 10]
active_user_fam_Bahman_1398=EPG_Bahman_1398.iat[40, 10]
active_user_televebion_Bahman_1398=EPG_Bahman_1398.iat[41, 10]
active_user_sepehr_Bahman_1398=EPG_Bahman_1398.iat[42, 10]
active_user_shima_Bahman_1398=EPG_Bahman_1398.iat[43, 10]
active_user_site_Bahman_1398=EPG_Bahman_1398.iat[44, 10]

all_visit_Bahman_1398=EPG_Bahman_1398.iat[25, 4]
all_duration_Bahman_1398=sum(EPG_Bahman_1398.iloc[1:24, 6])
all_content_sima_Bahman_1398=EPG_Bahman_1398.iat[25, 2]
all_register_user_Bahman_1398=sum(EPG_Bahman_1398.iloc[36:45, 4])
all_active_user_Bahman_1398=sum(EPG_Bahman_1398.iloc[36:45, 10])

Bahman_1398_sima_visit_channels=pd.DataFrame()
Bahman_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Bahman_1398, sima_2_visit_Bahman_1398, sima_3_visit_Bahman_1398,
                 sima_4_visit_Bahman_1398, sima_5_visit_Bahman_1398, sima_khabar_visit_Bahman_1398,
                 sima_ofogh_visit_Bahman_1398, sima_pooya_visit_Bahman_1398, sima_omid_visit_Bahman_1398,
                 sima_ifilm_visit_Bahman_1398, sima_namayesh_visit_Bahman_1398, sima_tamasha_visit_Bahman_1398,
                 sima_mostanad_visit_Bahman_1398, sima_shoma_visit_Bahman_1398, sima_amozesh_visit_Bahman_1398,
                 sima_varzesh_visit_Bahman_1398, sima_nasim_visit_Bahman_1398, sima_qoran_visit_Bahman_1398,
                 sima_salamat_visit_Bahman_1398, sima_irankala_visit_Bahman_1398, sima_alalam_visit_Bahman_1398,
                 sima_alkosar_visit_Bahman_1398, sima_presstv_visit_Bahman_1398, sima_sepehr_visit_Bahman_1398,],
        'duration': [sima_1_duration_Bahman_1398, sima_2_duration_Bahman_1398, sima_3_duration_Bahman_1398,
                 sima_4_duration_Bahman_1398, sima_5_duration_Bahman_1398, sima_khabar_duration_Bahman_1398,
                 sima_ofogh_duration_Bahman_1398, sima_pooya_duration_Bahman_1398, sima_omid_duration_Bahman_1398,
                 sima_ifilm_duration_Bahman_1398, sima_namayesh_duration_Bahman_1398, sima_tamasha_duration_Bahman_1398,
                 sima_mostanad_duration_Bahman_1398, sima_shoma_duration_Bahman_1398, sima_amozesh_duration_Bahman_1398,
                 sima_varzesh_duration_Bahman_1398, sima_nasim_duration_Bahman_1398, sima_qoran_duration_Bahman_1398,
                 sima_salamat_duration_Bahman_1398, sima_irankala_duration_Bahman_1398, sima_alalam_duration_Bahman_1398,
                 sima_alkosar_duration_Bahman_1398, sima_presstv_duration_Bahman_1398, sima_sepehr_duration_Bahman_1398,],}
Bahman_1398_sima_visit_channels=pd.DataFrame(Bahman_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Bahman_1398_sima_visit_channels=Bahman_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Bahman_1398_operator_data=pd.DataFrame()
Bahman_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها'],
       'visit': [sima_lenz_visit_Bahman_1398, sima_aio_visit_Bahman_1398, sima_anten_visit_Bahman_1398,
                 sima_tva_visit_Bahman_1398, sima_fam_visit_Bahman_1398,sima_televebion_visit_Bahman_1398,
                 sima_sepehr_visit_Bahman_1398, sima_shima_visit_Bahman_1398, sima_site_visit_Bahman_1398,],
       'register': [register_user_lenz_Bahman_1398, register_user_aio_Bahman_1398, register_user_anten_Bahman_1398,
                 register_user_tva_Bahman_1398, register_user_fam_Bahman_1398, register_user_televebion_Bahman_1398,
                 register_user_sepehr_Bahman_1398, register_user_shima_Bahman_1398, register_user_site_Bahman_1398,],
       'active': [active_user_lenz_Bahman_1398, active_user_aio_Bahman_1398, active_user_anten_Bahman_1398,
                 active_user_tva_Bahman_1398, active_user_fam_Bahman_1398, active_user_televebion_Bahman_1398,
                 active_user_sepehr_Bahman_1398, active_user_shima_Bahman_1398, register_user_site_Bahman_1398,],}

Bahman_1398_operator_data=pd.DataFrame(Bahman_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])

Bahman_1398_operator_data=Bahman_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Bahman_1398_all_data_summary=pd.DataFrame()
Bahman_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Bahman_1398, all_duration_Bahman_1398,all_content_sima_Bahman_1398, all_register_user_Bahman_1398,],}

Bahman_1398_all_data_summary=pd.DataFrame(Bahman_1398_all_data_summary, columns=['parameters', 'statistics'])

Bahman_1398_all_data_summary=Bahman_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه بهمن 1398.xlsx', engine='xlsxwriter')
Bahman_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Bahman_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Bahman_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه بهمن')
writer.save()

        ########################### اسفند #############################
print("EPG Esfand 1398")
EPG_Esfand_1398=pd.read_excel('EPG/EPG 1398/EPG Esfand 1398.xlsx', sheet_name='آمار')
EPG_Esfand_1398.fillna(0, inplace=True)
EPG_Esfand_1398['مدت زمان بازدید']=EPG_Esfand_1398.iloc[1:26, 9]*60
sima_1_visit_Esfand_1398=EPG_Esfand_1398.iat[1, 4]
sima_2_visit_Esfand_1398=EPG_Esfand_1398.iat[2, 4]
sima_3_visit_Esfand_1398=EPG_Esfand_1398.iat[3, 4]
sima_4_visit_Esfand_1398=EPG_Esfand_1398.iat[4, 4]
sima_5_visit_Esfand_1398=EPG_Esfand_1398.iat[5, 4]
sima_khabar_visit_Esfand_1398=EPG_Esfand_1398.iat[6, 4]
sima_ofogh_visit_Esfand_1398=EPG_Esfand_1398.iat[7, 4]
sima_pooya_visit_Esfand_1398=EPG_Esfand_1398.iat[8, 4]
sima_omid_visit_Esfand_1398=EPG_Esfand_1398.iat[9, 4]
sima_ifilm_visit_Esfand_1398=EPG_Esfand_1398.iat[10, 4]
sima_namayesh_visit_Esfand_1398=EPG_Esfand_1398.iat[11, 4]
sima_tamasha_visit_Esfand_1398=EPG_Esfand_1398.iat[12, 4]
sima_mostanad_visit_Esfand_1398=EPG_Esfand_1398.iat[13, 4]
sima_shoma_visit_Esfand_1398=EPG_Esfand_1398.iat[14, 4]
sima_amozesh_visit_Esfand_1398=EPG_Esfand_1398.iat[15, 4]
sima_varzesh_visit_Esfand_1398=EPG_Esfand_1398.iat[16, 4]
sima_nasim_visit_Esfand_1398=EPG_Esfand_1398.iat[17, 4]
sima_qoran_visit_Esfand_1398=EPG_Esfand_1398.iat[18, 4]
sima_salamat_visit_Esfand_1398=EPG_Esfand_1398.iat[19, 4]
sima_irankala_visit_Esfand_1398=EPG_Esfand_1398.iat[20, 4]
sima_alalam_visit_Esfand_1398=EPG_Esfand_1398.iat[21, 4]
sima_alkosar_visit_Esfand_1398=EPG_Esfand_1398.iat[22, 4]
sima_presstv_visit_Esfand_1398=EPG_Esfand_1398.iat[23, 4]
sima_sepehr_visit_Esfand_1398=EPG_Esfand_1398.iat[24, 4]

sima_1_duration_Esfand_1398=EPG_Esfand_1398.iat[1, 6]
sima_2_duration_Esfand_1398=EPG_Esfand_1398.iat[2, 6]
sima_3_duration_Esfand_1398=EPG_Esfand_1398.iat[3, 6]
sima_4_duration_Esfand_1398=EPG_Esfand_1398.iat[4, 6]
sima_5_duration_Esfand_1398=EPG_Esfand_1398.iat[5, 6]
sima_khabar_duration_Esfand_1398=EPG_Esfand_1398.iat[6, 6]
sima_ofogh_duration_Esfand_1398=EPG_Esfand_1398.iat[7, 6]
sima_pooya_duration_Esfand_1398=EPG_Esfand_1398.iat[8, 6]
sima_omid_duration_Esfand_1398=EPG_Esfand_1398.iat[9, 6]
sima_ifilm_duration_Esfand_1398=EPG_Esfand_1398.iat[10, 6]
sima_namayesh_duration_Esfand_1398=EPG_Esfand_1398.iat[11, 6]
sima_tamasha_duration_Esfand_1398=EPG_Esfand_1398.iat[12, 6]
sima_mostanad_duration_Esfand_1398=EPG_Esfand_1398.iat[13, 6]
sima_shoma_duration_Esfand_1398=EPG_Esfand_1398.iat[14, 6]
sima_amozesh_duration_Esfand_1398=EPG_Esfand_1398.iat[15, 6]
sima_varzesh_duration_Esfand_1398=EPG_Esfand_1398.iat[16, 6]
sima_nasim_duration_Esfand_1398=EPG_Esfand_1398.iat[17, 6]
sima_qoran_duration_Esfand_1398=EPG_Esfand_1398.iat[18, 6]
sima_salamat_duration_Esfand_1398=EPG_Esfand_1398.iat[19, 6]
sima_irankala_duration_Esfand_1398=EPG_Esfand_1398.iat[20, 6]
sima_alalam_duration_Esfand_1398=EPG_Esfand_1398.iat[21, 6]
sima_alkosar_duration_Esfand_1398=EPG_Esfand_1398.iat[22, 6]
sima_presstv_duration_Esfand_1398=EPG_Esfand_1398.iat[23, 6]
sima_sepehr_duration_Esfand_1398=EPG_Esfand_1398.iat[24, 6]

sima_lenz_visit_Esfand_1398=EPG_Esfand_1398.iat[36, 2]
sima_aio_visit_Esfand_1398=EPG_Esfand_1398.iat[37, 2]
sima_anten_visit_Esfand_1398=EPG_Esfand_1398.iat[38, 2]
sima_tva_visit_Esfand_1398=EPG_Esfand_1398.iat[39, 2]
sima_fam_visit_Esfand_1398=EPG_Esfand_1398.iat[40, 2]
sima_televebion_visit_Esfand_1398=EPG_Esfand_1398.iat[41, 2]
sima_sepehr_visit_Esfand_1398=EPG_Esfand_1398.iat[42, 2]
sima_shima_visit_Esfand_1398=EPG_Esfand_1398.iat[43, 2]
sima_site_visit_Esfand_1398=EPG_Esfand_1398.iat[44, 2]

register_user_lenz_Esfand_1398=EPG_Esfand_1398.iat[36, 4]
register_user_aio_Esfand_1398=EPG_Esfand_1398.iat[37, 4]
register_user_anten_Esfand_1398=EPG_Esfand_1398.iat[38, 4]
register_user_tva_Esfand_1398=EPG_Esfand_1398.iat[39, 4]
register_user_fam_Esfand_1398=EPG_Esfand_1398.iat[40, 4]
register_user_televebion_Esfand_1398=EPG_Esfand_1398.iat[41, 4]
register_user_sepehr_Esfand_1398=EPG_Esfand_1398.iat[42, 4]
register_user_shima_Esfand_1398=EPG_Esfand_1398.iat[43, 4]
register_user_site_Esfand_1398=EPG_Esfand_1398.iat[44, 4]

active_user_lenz_Esfand_1398=EPG_Esfand_1398.iat[36, 10]
active_user_aio_Esfand_1398=EPG_Esfand_1398.iat[37, 10]
active_user_anten_Esfand_1398=EPG_Esfand_1398.iat[38, 10]
active_user_tva_Esfand_1398=EPG_Esfand_1398.iat[39, 10]
active_user_fam_Esfand_1398=EPG_Esfand_1398.iat[40, 10]
active_user_televebion_Esfand_1398=EPG_Esfand_1398.iat[41, 10]
active_user_sepehr_Esfand_1398=EPG_Esfand_1398.iat[42, 10]
active_user_shima_Esfand_1398=EPG_Esfand_1398.iat[43, 10]
active_user_site_Esfand_1398=EPG_Esfand_1398.iat[44, 10]

all_visit_Esfand_1398=EPG_Esfand_1398.iat[25, 4]
all_duration_Esfand_1398=sum(EPG_Esfand_1398.iloc[1:24, 6])
all_content_sima_Esfand_1398=EPG_Esfand_1398.iat[25, 2]
all_register_user_Esfand_1398=sum(EPG_Esfand_1398.iloc[36:45, 4])
all_active_user_Esfand_1398=sum(EPG_Esfand_1398.iloc[36:45, 10])

Esfand_1398_sima_visit_channels=pd.DataFrame()
Esfand_1398_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [sima_1_visit_Esfand_1398, sima_2_visit_Esfand_1398, sima_3_visit_Esfand_1398,
                 sima_4_visit_Esfand_1398, sima_5_visit_Esfand_1398, sima_khabar_visit_Esfand_1398,
                 sima_ofogh_visit_Esfand_1398, sima_pooya_visit_Esfand_1398, sima_omid_visit_Esfand_1398,
                 sima_ifilm_visit_Esfand_1398, sima_namayesh_visit_Esfand_1398, sima_tamasha_visit_Esfand_1398,
                 sima_mostanad_visit_Esfand_1398, sima_shoma_visit_Esfand_1398, sima_amozesh_visit_Esfand_1398,
                 sima_varzesh_visit_Esfand_1398, sima_nasim_visit_Esfand_1398, sima_qoran_visit_Esfand_1398,
                 sima_salamat_visit_Esfand_1398, sima_irankala_visit_Esfand_1398, sima_alalam_visit_Esfand_1398,
                 sima_alkosar_visit_Esfand_1398, sima_presstv_visit_Esfand_1398, sima_sepehr_visit_Esfand_1398,],
        'duration': [sima_1_duration_Esfand_1398, sima_2_duration_Esfand_1398, sima_3_duration_Esfand_1398,
                 sima_4_duration_Esfand_1398, sima_5_duration_Esfand_1398, sima_khabar_duration_Esfand_1398,
                 sima_ofogh_duration_Esfand_1398, sima_pooya_duration_Esfand_1398, sima_omid_duration_Esfand_1398,
                 sima_ifilm_duration_Esfand_1398, sima_namayesh_duration_Esfand_1398, sima_tamasha_duration_Esfand_1398,
                 sima_mostanad_duration_Esfand_1398, sima_shoma_duration_Esfand_1398, sima_amozesh_duration_Esfand_1398,
                 sima_varzesh_duration_Esfand_1398, sima_nasim_duration_Esfand_1398, sima_qoran_duration_Esfand_1398,
                 sima_salamat_duration_Esfand_1398, sima_irankala_duration_Esfand_1398, sima_alalam_duration_Esfand_1398,
                 sima_alkosar_duration_Esfand_1398, sima_presstv_duration_Esfand_1398, sima_sepehr_duration_Esfand_1398,],}
Esfand_1398_sima_visit_channels=pd.DataFrame(Esfand_1398_sima_visit_channels, columns=['channels', 'visit', 'duration'])

Esfand_1398_sima_visit_channels=Esfand_1398_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

Esfand_1398_operator_data=pd.DataFrame()
Esfand_1398_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها'],
       'visit': [sima_lenz_visit_Esfand_1398, sima_aio_visit_Esfand_1398, sima_anten_visit_Esfand_1398,
                 sima_tva_visit_Esfand_1398, sima_fam_visit_Esfand_1398,sima_televebion_visit_Esfand_1398,
                 sima_sepehr_visit_Esfand_1398, sima_shima_visit_Esfand_1398, sima_site_visit_Esfand_1398,],
       'register': [register_user_lenz_Esfand_1398, register_user_aio_Esfand_1398, register_user_anten_Esfand_1398,
                 register_user_tva_Esfand_1398, register_user_fam_Esfand_1398, register_user_televebion_Esfand_1398,
                 register_user_sepehr_Esfand_1398, register_user_shima_Esfand_1398, register_user_site_Esfand_1398,],
       'active': [active_user_lenz_Esfand_1398, active_user_aio_Esfand_1398, active_user_anten_Esfand_1398,
                 active_user_tva_Esfand_1398, active_user_fam_Esfand_1398, active_user_televebion_Esfand_1398,
                 active_user_sepehr_Esfand_1398, active_user_shima_Esfand_1398, register_user_site_Esfand_1398,],}

Esfand_1398_operator_data=pd.DataFrame(Esfand_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])

Esfand_1398_operator_data=Esfand_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})

Esfand_1398_all_data_summary=pd.DataFrame()
Esfand_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_visit_Esfand_1398, all_duration_Esfand_1398,all_content_sima_Esfand_1398, all_register_user_Esfand_1398,],}

Esfand_1398_all_data_summary=pd.DataFrame(Esfand_1398_all_data_summary, columns=['parameters', 'statistics'])

Esfand_1398_all_data_summary=Esfand_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/ماه اسفند 1398.xlsx', engine='xlsxwriter')
Esfand_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
Esfand_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
Esfand_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه اسفند')
writer.save()

  ########################### total 1398 #############################
EPG_1398_total_visit_channels=pd.DataFrame()
EPG_1398_total_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'Farvardin': [sima_1_visit_Farvardin_1398, sima_2_visit_Farvardin_1398, sima_3_visit_Farvardin_1398,
                 sima_4_visit_Farvardin_1398, sima_5_visit_Farvardin_1398, sima_khabar_visit_Farvardin_1398,
                 sima_ofogh_visit_Farvardin_1398, sima_pooya_visit_Farvardin_1398, sima_omid_visit_Farvardin_1398,
                 sima_ifilm_visit_Farvardin_1398, sima_namayesh_visit_Farvardin_1398, sima_tamasha_visit_Farvardin_1398,
                 sima_mostanad_visit_Farvardin_1398, sima_shoma_visit_Farvardin_1398, sima_amozesh_visit_Farvardin_1398,
                 sima_varzesh_visit_Farvardin_1398, sima_nasim_visit_Farvardin_1398, sima_qoran_visit_Farvardin_1398,
                 sima_salamat_visit_Farvardin_1398, sima_irankala_visit_Farvardin_1398, sima_alalam_visit_Farvardin_1398,
                 sima_alkosar_visit_Farvardin_1398, sima_presstv_visit_Farvardin_1398, sima_sepehr_visit_Farvardin_1398,],
        'Ordibehesht': [sima_1_visit_Ordibehesht_1398, sima_2_visit_Ordibehesht_1398, sima_3_visit_Ordibehesht_1398,
                 sima_4_visit_Ordibehesht_1398, sima_5_visit_Ordibehesht_1398, sima_khabar_visit_Ordibehesht_1398,
                 sima_ofogh_visit_Ordibehesht_1398, sima_pooya_visit_Ordibehesht_1398, sima_omid_visit_Ordibehesht_1398,
                 sima_ifilm_visit_Ordibehesht_1398, sima_namayesh_visit_Ordibehesht_1398, sima_tamasha_visit_Ordibehesht_1398,
                 sima_mostanad_visit_Ordibehesht_1398, sima_shoma_visit_Ordibehesht_1398, sima_amozesh_visit_Ordibehesht_1398,
                 sima_varzesh_visit_Ordibehesht_1398, sima_nasim_visit_Ordibehesht_1398, sima_qoran_visit_Ordibehesht_1398,
                 sima_salamat_visit_Ordibehesht_1398, sima_irankala_visit_Ordibehesht_1398, sima_alalam_visit_Ordibehesht_1398,
                 sima_alkosar_visit_Ordibehesht_1398, sima_presstv_visit_Ordibehesht_1398, sima_sepehr_visit_Ordibehesht_1398,],
        'Khordad': [sima_1_visit_Khordad_1398, sima_2_visit_Khordad_1398, sima_3_visit_Khordad_1398,
                 sima_4_visit_Khordad_1398, sima_5_visit_Khordad_1398, sima_khabar_visit_Khordad_1398,
                 sima_ofogh_visit_Khordad_1398, sima_pooya_visit_Khordad_1398, sima_omid_visit_Khordad_1398,
                 sima_ifilm_visit_Khordad_1398, sima_namayesh_visit_Khordad_1398, sima_tamasha_visit_Khordad_1398,
                 sima_mostanad_visit_Khordad_1398, sima_shoma_visit_Khordad_1398, sima_amozesh_visit_Khordad_1398,
                 sima_varzesh_visit_Khordad_1398, sima_nasim_visit_Khordad_1398, sima_qoran_visit_Khordad_1398,
                 sima_salamat_visit_Khordad_1398, sima_irankala_visit_Khordad_1398, sima_alalam_visit_Khordad_1398,
                 sima_alkosar_visit_Khordad_1398, sima_presstv_visit_Khordad_1398, sima_sepehr_visit_Khordad_1398,],
        'Tir': [sima_1_visit_Tir_1398, sima_2_visit_Tir_1398, sima_3_visit_Tir_1398,
                 sima_4_visit_Tir_1398, sima_5_visit_Tir_1398, sima_khabar_visit_Tir_1398,
                 sima_ofogh_visit_Tir_1398, sima_pooya_visit_Tir_1398, sima_omid_visit_Tir_1398,
                 sima_ifilm_visit_Tir_1398, sima_namayesh_visit_Tir_1398, sima_tamasha_visit_Tir_1398,
                 sima_mostanad_visit_Tir_1398, sima_shoma_visit_Tir_1398, sima_amozesh_visit_Tir_1398,
                 sima_varzesh_visit_Tir_1398, sima_nasim_visit_Tir_1398, sima_qoran_visit_Tir_1398,
                 sima_salamat_visit_Tir_1398, sima_irankala_visit_Tir_1398, sima_alalam_visit_Tir_1398,
                 sima_alkosar_visit_Tir_1398, sima_presstv_visit_Tir_1398, sima_sepehr_visit_Tir_1398,],
        'Mordad': [sima_1_visit_Mordad_1398, sima_2_visit_Mordad_1398, sima_3_visit_Mordad_1398,
                 sima_4_visit_Mordad_1398, sima_5_visit_Mordad_1398, sima_khabar_visit_Mordad_1398,
                 sima_ofogh_visit_Mordad_1398, sima_pooya_visit_Mordad_1398, sima_omid_visit_Mordad_1398,
                 sima_ifilm_visit_Mordad_1398, sima_namayesh_visit_Mordad_1398, sima_tamasha_visit_Mordad_1398,
                 sima_mostanad_visit_Mordad_1398, sima_shoma_visit_Mordad_1398, sima_amozesh_visit_Mordad_1398,
                 sima_varzesh_visit_Mordad_1398, sima_nasim_visit_Mordad_1398, sima_qoran_visit_Mordad_1398,
                 sima_salamat_visit_Mordad_1398, sima_irankala_visit_Mordad_1398, sima_alalam_visit_Mordad_1398,
                 sima_alkosar_visit_Mordad_1398, sima_presstv_visit_Mordad_1398, sima_sepehr_visit_Mordad_1398,],
        'Shahrivar': [sima_1_visit_Shahrivar_1398, sima_2_visit_Shahrivar_1398, sima_3_visit_Shahrivar_1398,
                 sima_4_visit_Shahrivar_1398, sima_5_visit_Shahrivar_1398, sima_khabar_visit_Shahrivar_1398,
                 sima_ofogh_visit_Shahrivar_1398, sima_pooya_visit_Shahrivar_1398, sima_omid_visit_Shahrivar_1398,
                 sima_ifilm_visit_Shahrivar_1398, sima_namayesh_visit_Shahrivar_1398, sima_tamasha_visit_Shahrivar_1398,
                 sima_mostanad_visit_Shahrivar_1398, sima_shoma_visit_Shahrivar_1398, sima_amozesh_visit_Shahrivar_1398,
                 sima_varzesh_visit_Shahrivar_1398, sima_nasim_visit_Shahrivar_1398, sima_qoran_visit_Shahrivar_1398,
                 sima_salamat_visit_Shahrivar_1398, sima_irankala_visit_Shahrivar_1398, sima_alalam_visit_Shahrivar_1398,
                 sima_alkosar_visit_Shahrivar_1398, sima_presstv_visit_Shahrivar_1398, sima_sepehr_visit_Shahrivar_1398,],
        'Mehr': [sima_1_visit_Mehr_1398, sima_2_visit_Mehr_1398, sima_3_visit_Mehr_1398,
                 sima_4_visit_Mehr_1398, sima_5_visit_Mehr_1398, sima_khabar_visit_Mehr_1398,
                 sima_ofogh_visit_Mehr_1398, sima_pooya_visit_Mehr_1398, sima_omid_visit_Mehr_1398,
                 sima_ifilm_visit_Mehr_1398, sima_namayesh_visit_Mehr_1398, sima_tamasha_visit_Mehr_1398,
                 sima_mostanad_visit_Mehr_1398, sima_shoma_visit_Mehr_1398, sima_amozesh_visit_Mehr_1398,
                 sima_varzesh_visit_Mehr_1398, sima_nasim_visit_Mehr_1398, sima_qoran_visit_Mehr_1398,
                 sima_salamat_visit_Mehr_1398, sima_irankala_visit_Mehr_1398, sima_alalam_visit_Mehr_1398,
                 sima_alkosar_visit_Mehr_1398, sima_presstv_visit_Mehr_1398, sima_sepehr_visit_Mehr_1398,],
        'Aban': [sima_1_visit_Aban_1398, sima_2_visit_Aban_1398, sima_3_visit_Aban_1398,
                 sima_4_visit_Aban_1398, sima_5_visit_Aban_1398, sima_khabar_visit_Aban_1398,
                 sima_ofogh_visit_Aban_1398, sima_pooya_visit_Aban_1398, sima_omid_visit_Aban_1398,
                 sima_ifilm_visit_Aban_1398, sima_namayesh_visit_Aban_1398, sima_tamasha_visit_Aban_1398,
                 sima_mostanad_visit_Aban_1398, sima_shoma_visit_Aban_1398, sima_amozesh_visit_Aban_1398,
                 sima_varzesh_visit_Aban_1398, sima_nasim_visit_Aban_1398, sima_qoran_visit_Aban_1398,
                 sima_salamat_visit_Aban_1398, sima_irankala_visit_Aban_1398, sima_alalam_visit_Aban_1398,
                 sima_alkosar_visit_Aban_1398, sima_presstv_visit_Aban_1398, sima_sepehr_visit_Aban_1398,],
        'Azar': [sima_1_visit_Azar_1398, sima_2_visit_Azar_1398, sima_3_visit_Azar_1398,
                 sima_4_visit_Azar_1398, sima_5_visit_Azar_1398, sima_khabar_visit_Azar_1398,
                 sima_ofogh_visit_Azar_1398, sima_pooya_visit_Azar_1398, sima_omid_visit_Azar_1398,
                 sima_ifilm_visit_Azar_1398, sima_namayesh_visit_Azar_1398, sima_tamasha_visit_Azar_1398,
                 sima_mostanad_visit_Azar_1398, sima_shoma_visit_Azar_1398, sima_amozesh_visit_Azar_1398,
                 sima_varzesh_visit_Azar_1398, sima_nasim_visit_Azar_1398, sima_qoran_visit_Azar_1398,
                 sima_salamat_visit_Azar_1398, sima_irankala_visit_Azar_1398, sima_alalam_visit_Azar_1398,
                 sima_alkosar_visit_Azar_1398, sima_presstv_visit_Azar_1398, sima_sepehr_visit_Azar_1398,],
        'Dey': [sima_1_visit_Dey_1398, sima_2_visit_Dey_1398, sima_3_visit_Dey_1398,
                 sima_4_visit_Dey_1398, sima_5_visit_Dey_1398, sima_khabar_visit_Dey_1398,
                 sima_ofogh_visit_Dey_1398, sima_pooya_visit_Dey_1398, sima_omid_visit_Dey_1398,
                 sima_ifilm_visit_Dey_1398, sima_namayesh_visit_Dey_1398, sima_tamasha_visit_Dey_1398,
                 sima_mostanad_visit_Dey_1398, sima_shoma_visit_Dey_1398, sima_amozesh_visit_Dey_1398,
                 sima_varzesh_visit_Dey_1398, sima_nasim_visit_Dey_1398, sima_qoran_visit_Dey_1398,
                 sima_salamat_visit_Dey_1398, sima_irankala_visit_Dey_1398, sima_alalam_visit_Dey_1398,
                 sima_alkosar_visit_Dey_1398, sima_presstv_visit_Dey_1398, sima_sepehr_visit_Dey_1398,],
        'Bahman': [sima_1_visit_Bahman_1398, sima_2_visit_Bahman_1398, sima_3_visit_Bahman_1398,
                 sima_4_visit_Bahman_1398, sima_5_visit_Bahman_1398, sima_khabar_visit_Bahman_1398,
                 sima_ofogh_visit_Bahman_1398, sima_pooya_visit_Bahman_1398, sima_omid_visit_Bahman_1398,
                 sima_ifilm_visit_Bahman_1398, sima_namayesh_visit_Bahman_1398, sima_tamasha_visit_Bahman_1398,
                 sima_mostanad_visit_Bahman_1398, sima_shoma_visit_Bahman_1398, sima_amozesh_visit_Bahman_1398,
                 sima_varzesh_visit_Bahman_1398, sima_nasim_visit_Bahman_1398, sima_qoran_visit_Bahman_1398,
                 sima_salamat_visit_Bahman_1398, sima_irankala_visit_Bahman_1398, sima_alalam_visit_Bahman_1398,
                 sima_alkosar_visit_Bahman_1398, sima_presstv_visit_Bahman_1398, sima_sepehr_visit_Bahman_1398,],
        'Esfand': [sima_1_visit_Esfand_1398, sima_2_visit_Esfand_1398, sima_3_visit_Esfand_1398,
                 sima_4_visit_Esfand_1398, sima_5_visit_Esfand_1398, sima_khabar_visit_Esfand_1398,
                 sima_ofogh_visit_Esfand_1398, sima_pooya_visit_Esfand_1398, sima_omid_visit_Esfand_1398,
                 sima_ifilm_visit_Esfand_1398, sima_namayesh_visit_Esfand_1398, sima_tamasha_visit_Esfand_1398,
                 sima_mostanad_visit_Esfand_1398, sima_shoma_visit_Esfand_1398, sima_amozesh_visit_Esfand_1398,
                 sima_varzesh_visit_Esfand_1398, sima_nasim_visit_Esfand_1398, sima_qoran_visit_Esfand_1398,
                 sima_salamat_visit_Esfand_1398, sima_irankala_visit_Esfand_1398, sima_alalam_visit_Esfand_1398,
                 sima_alkosar_visit_Esfand_1398, sima_presstv_visit_Esfand_1398, sima_sepehr_visit_Esfand_1398,],}

EPG_1398_total_visit_channels=pd.DataFrame(EPG_1398_total_visit_channels, columns=['channels', 'Farvardin', 'Ordibehesht', 'Khordad', 'Tir', 'Mordad', 'Shahrivar', 'Mehr', 'Aban', 'Azar', 'Dey', 'Bahman', 'Esfand'])

EPG_1398_total_visit_channels=EPG_1398_total_visit_channels.rename(columns={'channels': 'نام شبکه', 'Farvardin': 'فروردین', 'Ordibehesht': 'اردیبهشت',
                                                                                          'Khordad': 'خرداد', 'Tir': 'تیر', 'Mordad': 'مرداد',
                                                                                          'Shahrivar': 'شهریور', 'Mehr': 'مهر', 'Aban': 'آبان', 
                                                                                          'Azar': 'آذر', 'Dey': 'دی', 'Bahman': 'بهمن', 'Esfand': 'اسفند',})


EPG_1398_total_duration_channels=pd.DataFrame()
EPG_1398_total_duration_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'Farvardin': [sima_1_duration_Farvardin_1398, sima_2_duration_Farvardin_1398, sima_3_duration_Farvardin_1398,
                 sima_4_duration_Farvardin_1398, sima_5_duration_Farvardin_1398, sima_khabar_duration_Farvardin_1398,
                 sima_ofogh_duration_Farvardin_1398, sima_pooya_duration_Farvardin_1398, sima_omid_duration_Farvardin_1398,
                 sima_ifilm_duration_Farvardin_1398, sima_namayesh_duration_Farvardin_1398, sima_tamasha_duration_Farvardin_1398,
                 sima_mostanad_duration_Farvardin_1398, sima_shoma_duration_Farvardin_1398, sima_amozesh_duration_Farvardin_1398,
                 sima_varzesh_duration_Farvardin_1398, sima_nasim_duration_Farvardin_1398, sima_qoran_duration_Farvardin_1398,
                 sima_salamat_duration_Farvardin_1398, sima_irankala_duration_Farvardin_1398, sima_alalam_duration_Farvardin_1398,
                 sima_alkosar_duration_Farvardin_1398, sima_presstv_duration_Farvardin_1398, sima_sepehr_duration_Farvardin_1398,],
        'Ordibehesht': [sima_1_duration_Ordibehesht_1398, sima_2_duration_Ordibehesht_1398, sima_3_duration_Ordibehesht_1398,
                 sima_4_duration_Ordibehesht_1398, sima_5_duration_Ordibehesht_1398, sima_khabar_duration_Ordibehesht_1398,
                 sima_ofogh_duration_Ordibehesht_1398, sima_pooya_duration_Ordibehesht_1398, sima_omid_duration_Ordibehesht_1398,
                 sima_ifilm_duration_Ordibehesht_1398, sima_namayesh_duration_Ordibehesht_1398, sima_tamasha_duration_Ordibehesht_1398,
                 sima_mostanad_duration_Ordibehesht_1398, sima_shoma_duration_Ordibehesht_1398, sima_amozesh_duration_Ordibehesht_1398,
                 sima_varzesh_duration_Ordibehesht_1398, sima_nasim_duration_Ordibehesht_1398, sima_qoran_duration_Ordibehesht_1398,
                 sima_salamat_duration_Ordibehesht_1398, sima_irankala_duration_Ordibehesht_1398, sima_alalam_duration_Ordibehesht_1398,
                 sima_alkosar_duration_Ordibehesht_1398, sima_presstv_duration_Ordibehesht_1398, sima_sepehr_duration_Ordibehesht_1398,],
        'Khordad': [sima_1_duration_Khordad_1398, sima_2_duration_Khordad_1398, sima_3_duration_Khordad_1398,
                 sima_4_duration_Khordad_1398, sima_5_duration_Khordad_1398, sima_khabar_duration_Khordad_1398,
                 sima_ofogh_duration_Khordad_1398, sima_pooya_duration_Khordad_1398, sima_omid_duration_Khordad_1398,
                 sima_ifilm_duration_Khordad_1398, sima_namayesh_duration_Khordad_1398, sima_tamasha_duration_Khordad_1398,
                 sima_mostanad_duration_Khordad_1398, sima_shoma_duration_Khordad_1398, sima_amozesh_duration_Khordad_1398,
                 sima_varzesh_duration_Khordad_1398, sima_nasim_duration_Khordad_1398, sima_qoran_duration_Khordad_1398,
                 sima_salamat_duration_Khordad_1398, sima_irankala_duration_Khordad_1398, sima_alalam_duration_Khordad_1398,
                 sima_alkosar_duration_Khordad_1398, sima_presstv_duration_Khordad_1398, sima_sepehr_duration_Khordad_1398,],
        'Tir': [sima_1_duration_Tir_1398, sima_2_duration_Tir_1398, sima_3_duration_Tir_1398,
                 sima_4_duration_Tir_1398, sima_5_duration_Tir_1398, sima_khabar_duration_Tir_1398,
                 sima_ofogh_duration_Tir_1398, sima_pooya_duration_Tir_1398, sima_omid_duration_Tir_1398,
                 sima_ifilm_duration_Tir_1398, sima_namayesh_duration_Tir_1398, sima_tamasha_duration_Tir_1398,
                 sima_mostanad_duration_Tir_1398, sima_shoma_duration_Tir_1398, sima_amozesh_duration_Tir_1398,
                 sima_varzesh_duration_Tir_1398, sima_nasim_duration_Tir_1398, sima_qoran_duration_Tir_1398,
                 sima_salamat_duration_Tir_1398, sima_irankala_duration_Tir_1398, sima_alalam_duration_Tir_1398,
                 sima_alkosar_duration_Tir_1398, sima_presstv_duration_Tir_1398, sima_sepehr_duration_Tir_1398,],
        'Mordad': [sima_1_duration_Mordad_1398, sima_2_duration_Mordad_1398, sima_3_duration_Mordad_1398,
                 sima_4_duration_Mordad_1398, sima_5_duration_Mordad_1398, sima_khabar_duration_Mordad_1398,
                 sima_ofogh_duration_Mordad_1398, sima_pooya_duration_Mordad_1398, sima_omid_duration_Mordad_1398,
                 sima_ifilm_duration_Mordad_1398, sima_namayesh_duration_Mordad_1398, sima_tamasha_duration_Mordad_1398,
                 sima_mostanad_duration_Mordad_1398, sima_shoma_duration_Mordad_1398, sima_amozesh_duration_Mordad_1398,
                 sima_varzesh_duration_Mordad_1398, sima_nasim_duration_Mordad_1398, sima_qoran_duration_Mordad_1398,
                 sima_salamat_duration_Mordad_1398, sima_irankala_duration_Mordad_1398, sima_alalam_duration_Mordad_1398,
                 sima_alkosar_duration_Mordad_1398, sima_presstv_duration_Mordad_1398, sima_sepehr_duration_Mordad_1398,],
        'Shahrivar': [sima_1_duration_Shahrivar_1398, sima_2_duration_Shahrivar_1398, sima_3_duration_Shahrivar_1398,
                 sima_4_duration_Shahrivar_1398, sima_5_duration_Shahrivar_1398, sima_khabar_duration_Shahrivar_1398,
                 sima_ofogh_duration_Shahrivar_1398, sima_pooya_duration_Shahrivar_1398, sima_omid_duration_Shahrivar_1398,
                 sima_ifilm_duration_Shahrivar_1398, sima_namayesh_duration_Shahrivar_1398, sima_tamasha_duration_Shahrivar_1398,
                 sima_mostanad_duration_Shahrivar_1398, sima_shoma_duration_Shahrivar_1398, sima_amozesh_duration_Shahrivar_1398,
                 sima_varzesh_duration_Shahrivar_1398, sima_nasim_duration_Shahrivar_1398, sima_qoran_duration_Shahrivar_1398,
                 sima_salamat_duration_Shahrivar_1398, sima_irankala_duration_Shahrivar_1398, sima_alalam_duration_Shahrivar_1398,
                 sima_alkosar_duration_Shahrivar_1398, sima_presstv_duration_Shahrivar_1398, sima_sepehr_duration_Shahrivar_1398,],
        'Mehr': [sima_1_duration_Mehr_1398, sima_2_duration_Mehr_1398, sima_3_duration_Mehr_1398,
                 sima_4_duration_Mehr_1398, sima_5_duration_Mehr_1398, sima_khabar_duration_Mehr_1398,
                 sima_ofogh_duration_Mehr_1398, sima_pooya_duration_Mehr_1398, sima_omid_duration_Mehr_1398,
                 sima_ifilm_duration_Mehr_1398, sima_namayesh_duration_Mehr_1398, sima_tamasha_duration_Mehr_1398,
                 sima_mostanad_duration_Mehr_1398, sima_shoma_duration_Mehr_1398, sima_amozesh_duration_Mehr_1398,
                 sima_varzesh_duration_Mehr_1398, sima_nasim_duration_Mehr_1398, sima_qoran_duration_Mehr_1398,
                 sima_salamat_duration_Mehr_1398, sima_irankala_duration_Mehr_1398, sima_alalam_duration_Mehr_1398,
                 sima_alkosar_duration_Mehr_1398, sima_presstv_duration_Mehr_1398, sima_sepehr_duration_Mehr_1398,],
        'Aban': [sima_1_duration_Aban_1398, sima_2_duration_Aban_1398, sima_3_duration_Aban_1398,
                 sima_4_duration_Aban_1398, sima_5_duration_Aban_1398, sima_khabar_duration_Aban_1398,
                 sima_ofogh_duration_Aban_1398, sima_pooya_duration_Aban_1398, sima_omid_duration_Aban_1398,
                 sima_ifilm_duration_Aban_1398, sima_namayesh_duration_Aban_1398, sima_tamasha_duration_Aban_1398,
                 sima_mostanad_duration_Aban_1398, sima_shoma_duration_Aban_1398, sima_amozesh_duration_Aban_1398,
                 sima_varzesh_duration_Aban_1398, sima_nasim_duration_Aban_1398, sima_qoran_duration_Aban_1398,
                 sima_salamat_duration_Aban_1398, sima_irankala_duration_Aban_1398, sima_alalam_duration_Aban_1398,
                 sima_alkosar_duration_Aban_1398, sima_presstv_duration_Aban_1398, sima_sepehr_duration_Aban_1398,],
        'Azar': [sima_1_duration_Azar_1398, sima_2_duration_Azar_1398, sima_3_duration_Azar_1398,
                 sima_4_duration_Azar_1398, sima_5_duration_Azar_1398, sima_khabar_duration_Azar_1398,
                 sima_ofogh_duration_Azar_1398, sima_pooya_duration_Azar_1398, sima_omid_duration_Azar_1398,
                 sima_ifilm_duration_Azar_1398, sima_namayesh_duration_Azar_1398, sima_tamasha_duration_Azar_1398,
                 sima_mostanad_duration_Azar_1398, sima_shoma_duration_Azar_1398, sima_amozesh_duration_Azar_1398,
                 sima_varzesh_duration_Azar_1398, sima_nasim_duration_Azar_1398, sima_qoran_duration_Azar_1398,
                 sima_salamat_duration_Azar_1398, sima_irankala_duration_Azar_1398, sima_alalam_duration_Azar_1398,
                 sima_alkosar_duration_Azar_1398, sima_presstv_duration_Azar_1398, sima_sepehr_duration_Azar_1398,],
        'Dey': [sima_1_duration_Dey_1398, sima_2_duration_Dey_1398, sima_3_duration_Dey_1398,
                 sima_4_duration_Dey_1398, sima_5_duration_Dey_1398, sima_khabar_duration_Dey_1398,
                 sima_ofogh_duration_Dey_1398, sima_pooya_duration_Dey_1398, sima_omid_duration_Dey_1398,
                 sima_ifilm_duration_Dey_1398, sima_namayesh_duration_Dey_1398, sima_tamasha_duration_Dey_1398,
                 sima_mostanad_duration_Dey_1398, sima_shoma_duration_Dey_1398, sima_amozesh_duration_Dey_1398,
                 sima_varzesh_duration_Dey_1398, sima_nasim_duration_Dey_1398, sima_qoran_duration_Dey_1398,
                 sima_salamat_duration_Dey_1398, sima_irankala_duration_Dey_1398, sima_alalam_duration_Dey_1398,
                 sima_alkosar_duration_Dey_1398, sima_presstv_duration_Dey_1398, sima_sepehr_duration_Dey_1398,],
        'Bahman': [sima_1_duration_Bahman_1398, sima_2_duration_Bahman_1398, sima_3_duration_Bahman_1398,
                 sima_4_duration_Bahman_1398, sima_5_duration_Bahman_1398, sima_khabar_duration_Bahman_1398,
                 sima_ofogh_duration_Bahman_1398, sima_pooya_duration_Bahman_1398, sima_omid_duration_Bahman_1398,
                 sima_ifilm_duration_Bahman_1398, sima_namayesh_duration_Bahman_1398, sima_tamasha_duration_Bahman_1398,
                 sima_mostanad_duration_Bahman_1398, sima_shoma_duration_Bahman_1398, sima_amozesh_duration_Bahman_1398,
                 sima_varzesh_duration_Bahman_1398, sima_nasim_duration_Bahman_1398, sima_qoran_duration_Bahman_1398,
                 sima_salamat_duration_Bahman_1398, sima_irankala_duration_Bahman_1398, sima_alalam_duration_Bahman_1398,
                 sima_alkosar_duration_Bahman_1398, sima_presstv_duration_Bahman_1398, sima_sepehr_duration_Bahman_1398,],
        'Esfand': [sima_1_duration_Esfand_1398, sima_2_duration_Esfand_1398, sima_3_duration_Esfand_1398,
                 sima_4_duration_Esfand_1398, sima_5_duration_Esfand_1398, sima_khabar_duration_Esfand_1398,
                 sima_ofogh_duration_Esfand_1398, sima_pooya_duration_Esfand_1398, sima_omid_duration_Esfand_1398,
                 sima_ifilm_duration_Esfand_1398, sima_namayesh_duration_Esfand_1398, sima_tamasha_duration_Esfand_1398,
                 sima_mostanad_duration_Esfand_1398, sima_shoma_duration_Esfand_1398, sima_amozesh_duration_Esfand_1398,
                 sima_varzesh_duration_Esfand_1398, sima_nasim_duration_Esfand_1398, sima_qoran_duration_Esfand_1398,
                 sima_salamat_duration_Esfand_1398, sima_irankala_duration_Esfand_1398, sima_alalam_duration_Esfand_1398,
                 sima_alkosar_duration_Esfand_1398, sima_presstv_duration_Esfand_1398, sima_sepehr_duration_Esfand_1398,],}

EPG_1398_total_duration_channels=pd.DataFrame(EPG_1398_total_duration_channels, columns=['channels', 'Farvardin', 'Ordibehesht', 'Khordad', 'Tir', 'Mordad', 'Shahrivar', 'Mehr', 'Aban', 'Azar', 'Dey', 'Bahman', 'Esfand'])

EPG_1398_total_duration_channels=EPG_1398_total_duration_channels.rename(columns={'channels': 'نام شبکه', 'Farvardin': 'فروردین', 'Ordibehesht': 'اردیبهشت',
                                                                                          'Khordad': 'خرداد', 'Tir': 'تیر', 'Mordad': 'مرداد',
                                                                                          'Shahrivar': 'شهریور', 'Mehr': 'مهر', 'Aban': 'آبان', 
                                                                                          'Azar': 'آذر', 'Dey': 'دی', 'Bahman': 'بهمن', 'Esfand': 'اسفند',})

EPG_1398_total_operator_visit=pd.DataFrame()
EPG_1398_total_operator_visit={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها'],
       'Farvardin': [sima_lenz_visit_Farvardin_1398, sima_aio_visit_Farvardin_1398, sima_anten_visit_Farvardin_1398,
                 sima_tva_visit_Farvardin_1398, sima_fam_visit_Farvardin_1398,0,
                 0, 0, 0,],
       'Ordibehesht': [sima_lenz_visit_Ordibehesht_1398, sima_aio_visit_Ordibehesht_1398, sima_anten_visit_Ordibehesht_1398,
                 sima_tva_visit_Ordibehesht_1398, sima_fam_visit_Ordibehesht_1398,0,
                 0, 0, 0,],
       'Khordad': [sima_lenz_visit_Khordad_1398, sima_aio_visit_Khordad_1398, sima_anten_visit_Khordad_1398,
                 sima_tva_visit_Khordad_1398, sima_fam_visit_Khordad_1398,0,
                 0, 0, 0,],
       'Tir': [sima_lenz_visit_Tir_1398, sima_aio_visit_Tir_1398, sima_anten_visit_Tir_1398,
                 sima_tva_visit_Tir_1398, sima_fam_visit_Tir_1398,sima_televebion_visit_Tir_1398,
                 0, 0, 0,],
       'Mordad': [sima_lenz_visit_Mordad_1398, sima_aio_visit_Mordad_1398, sima_anten_visit_Mordad_1398,
                 sima_tva_visit_Mordad_1398, sima_fam_visit_Mordad_1398,sima_televebion_visit_Mordad_1398,
                 sima_sepehr_visit_Mordad_1398, 0, 0,],
       'Shahrivar': [sima_lenz_visit_Shahrivar_1398, sima_aio_visit_Shahrivar_1398, sima_anten_visit_Shahrivar_1398,
                 sima_tva_visit_Shahrivar_1398, sima_fam_visit_Shahrivar_1398,sima_televebion_visit_Shahrivar_1398,
                 sima_sepehr_visit_Shahrivar_1398, sima_shima_visit_Shahrivar_1398, 0,],
       'Mehr': [sima_lenz_visit_Mehr_1398, sima_aio_visit_Mehr_1398, sima_anten_visit_Mehr_1398,
                 sima_tva_visit_Mehr_1398, sima_fam_visit_Mehr_1398,sima_televebion_visit_Mehr_1398,
                 sima_sepehr_visit_Mehr_1398, sima_shima_visit_Mehr_1398, sima_site_visit_Mehr_1398,],
       'Aban': [sima_lenz_visit_Aban_1398, sima_aio_visit_Aban_1398, sima_anten_visit_Aban_1398,
                 sima_tva_visit_Aban_1398, sima_fam_visit_Aban_1398,sima_televebion_visit_Aban_1398,
                 sima_sepehr_visit_Aban_1398, sima_shima_visit_Aban_1398, sima_site_visit_Aban_1398,],
       'Azar': [sima_lenz_visit_Azar_1398, sima_aio_visit_Azar_1398, sima_anten_visit_Azar_1398,
                 sima_tva_visit_Azar_1398, sima_fam_visit_Azar_1398,sima_televebion_visit_Azar_1398,
                 sima_sepehr_visit_Azar_1398, sima_shima_visit_Azar_1398, sima_site_visit_Azar_1398,],
       'Dey': [sima_lenz_visit_Dey_1398, sima_aio_visit_Dey_1398, sima_anten_visit_Dey_1398,
                 sima_tva_visit_Dey_1398, sima_fam_visit_Dey_1398,sima_televebion_visit_Dey_1398,
                 sima_sepehr_visit_Dey_1398, sima_shima_visit_Dey_1398, sima_site_visit_Dey_1398,],
       'Bahman': [sima_lenz_visit_Bahman_1398, sima_aio_visit_Bahman_1398, sima_anten_visit_Bahman_1398,
                 sima_tva_visit_Bahman_1398, sima_fam_visit_Bahman_1398,sima_televebion_visit_Bahman_1398,
                 sima_sepehr_visit_Bahman_1398, sima_shima_visit_Bahman_1398, sima_site_visit_Bahman_1398,],
       'Esfand': [sima_lenz_visit_Esfand_1398, sima_aio_visit_Esfand_1398, sima_anten_visit_Esfand_1398,
                 sima_tva_visit_Esfand_1398, sima_fam_visit_Esfand_1398,sima_televebion_visit_Esfand_1398,
                 sima_sepehr_visit_Esfand_1398, sima_shima_visit_Esfand_1398, sima_site_visit_Esfand_1398,],}

EPG_1398_total_operator_visit=pd.DataFrame(EPG_1398_total_operator_visit, columns=['operators', 'Farvardin', 'Ordibehesht', 'Khordad', 'Tir', 'Mordad', 'Shahrivar', 'Mehr', 'Aban', 'Azar', 'Dey', 'Bahman', 'Esfand'])

EPG_1398_total_operator_visit=EPG_1398_total_operator_visit.rename(columns={'operators': 'اپراتور', 'Farvardin': 'فروردین', 'Ordibehesht': 'اردیبهشت',
                                                                                          'Khordad': 'خرداد', 'Tir': 'تیر', 'Mordad': 'مرداد',
                                                                                          'Shahrivar': 'شهریور', 'Mehr': 'مهر', 'Aban': 'آبان', 
                                                                                          'Azar': 'آذر', 'Dey': 'دی', 'Bahman': 'بهمن', 'Esfand': 'اسفند',})

all_visit_1398=all_visit_Farvardin_1398+all_visit_Ordibehesht_1398+all_visit_Khordad_1398+ \
all_visit_Tir_1398+all_visit_Mordad_1398+all_visit_Shahrivar_1398+ \
all_visit_Mehr_1398+all_visit_Aban_1398+all_visit_Azar_1398+ \
all_visit_Dey_1398+all_visit_Bahman_1398+all_visit_Esfand_1398

all_duration_1398=all_duration_Farvardin_1398+all_duration_Ordibehesht_1398+all_duration_Khordad_1398+ \
all_duration_Tir_1398+all_duration_Mordad_1398+all_duration_Shahrivar_1398+ \
all_duration_Mehr_1398+all_duration_Aban_1398+all_duration_Azar_1398+ \
all_duration_Dey_1398+all_duration_Bahman_1398+all_duration_Esfand_1398

EPG_1398_total_summary=pd.DataFrame()
EPG_1398_total_summary={'parameters': ['تعداد بازدید', 'مدت مان بازدید (به دقیقه)', 'تعداد کاربران ثبت نامی'],
       'statistics': [all_visit_1398, all_duration_1398, all_register_user_Esfand_1398],}

EPG_1398_total_summary=pd.DataFrame(EPG_1398_total_summary, columns=['parameters', 'statistics'])
EPG_1398_total_summary=EPG_1398_total_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})


writer = pd.ExcelWriter('output/خلاصه آمار سال 1398.xlsx', engine='xlsxwriter')
EPG_1398_total_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
EPG_1398_total_duration_channels.to_excel(writer, 'زمان بازدید شبکه ها (به دقیقه)')
EPG_1398_total_operator_visit.to_excel(writer, 'آمار اپراتورها')
EPG_1398_total_summary.to_excel(writer, 'خلاصه آمار سال 1398')
writer.save()

  
########################### EPG 1397 #############################
EPG_2=pd.read_excel('EPG/EPG 1397/EPG 2.xlsx', sheet_name='ALL')  
EPG_3=pd.read_excel('EPG/EPG 1397/EPG 3.xlsx', sheet_name='کل') 
EPG_4=pd.read_excel('EPG/EPG 1397/EPG 4.xlsx', sheet_name='کل') 
EPG_5=pd.read_excel('EPG/EPG 1397/EPG 5.xlsx', sheet_name='کل') 
EPG_esfand=pd.read_excel('EPG/EPG 1397/EPG Esfand.xlsx', sheet_name='کل')
del EPG_2['ساعت']
del EPG_2['تاریخ']
del EPG_2['Unnamed: 9']
del EPG_2['Unnamed: 10']
del EPG_2['Unnamed: 11']
del EPG_3['ساعت']
del EPG_3['تاریخ']
del EPG_3['نوع محتوا']
del EPG_3['Unnamed: 10']
del EPG_3['Unnamed: 11']
del EPG_4['ساعت']
del EPG_4['تاریخ']
del EPG_4['جنس']
del EPG_4['ماه']
del EPG_4['روز']
del EPG_4['Unnamed: 12']
del EPG_4['Unnamed: 13']
del EPG_4['Unnamed: 14']
del EPG_5['ساعت']
del EPG_5['تاریخ']
del EPG_5['جنس']
del EPG_5['ماه']
del EPG_5['روز']
del EPG_5['Unnamed: 12']
del EPG_5['Unnamed: 13']
del EPG_5['Unnamed: 14']
del EPG_esfand['ساعت']
del EPG_esfand['تاریخ']
del EPG_esfand['جنس']
del EPG_esfand['ماه']
del EPG_esfand['روز']
del EPG_esfand['ساعت.1']
del EPG_esfand['Unnamed: 13']
del EPG_esfand['Unnamed: 14']

all_data_EPG_1397 = EPG_2.append([EPG_3, EPG_4, EPG_5, EPG_esfand])
duration_1397=all_data_EPG_1397['مدت بازدید']   # convert duration to minute
duration_1397=duration_1397*60                   # convert duration to minute
all_data_EPG_1397['مدت بازدید']=duration_1397   # convert duration to minute

all_data_EPG_1397=all_data_EPG_1397.rename(columns={'نام شبکه': 'channels', 'نام برنامه': 'title',
                                                    'تاریخ شروع': 'start date', 'تاریخ پایان': 'end date',
                                                    'مدت بازدید': 'duration', 'تعداد بازدید': 'visit',
                                                    'نام اپراتور': 'operators' })

all_data_EPG_1397_total_visit=all_data_EPG_1397['visit'].sum()
all_data_EPG_1397_total_duration=all_data_EPG_1397['duration'].sum()
all_data_EPG_1397_total_register_user=15134898

EPG_tva_1397=all_data_EPG_1397.query("operators == 'تیوا'")
EPG_fam_1397=all_data_EPG_1397.query("operators == 'فام'")
EPG_anten_1397=all_data_EPG_1397.query("operators == 'آنتن'")
EPG_aio_1397=all_data_EPG_1397.query("operators == 'آیو'")
EPG_lenz_1397=all_data_EPG_1397.query("operators == 'لنز'")

EPG_tva_1397_visit=EPG_tva_1397['visit'].sum()
EPG_fam_1397_visit=EPG_fam_1397['visit'].sum()
EPG_anten_1397_visit=EPG_anten_1397['visit'].sum()
EPG_aio_1397_visit=EPG_aio_1397['visit'].sum()
EPG_lenz_1397_visit=EPG_lenz_1397['visit'].sum()

EPG_tva_1397_duration=EPG_tva_1397['duration'].sum()
EPG_fam_1397_duration=EPG_fam_1397['duration'].sum()
EPG_anten_1397_duration=EPG_anten_1397['duration'].sum()
EPG_aio_1397_duration=EPG_aio_1397['duration'].sum()
EPG_lenz_1397_duration=EPG_lenz_1397['duration'].sum()

all_data_EPG_1397_operators=pd.DataFrame()
all_data_EPG_1397_operators={'operators': ['تیوا', 'فام', 'آنتن', 'آیو', 'لنز',],
       'visit': [EPG_tva_1397_visit, EPG_fam_1397_visit, EPG_anten_1397_visit,
                 EPG_aio_1397_visit, EPG_lenz_1397_visit,],
        'duration': [EPG_tva_1397_duration, EPG_fam_1397_duration, EPG_anten_1397_duration,
                 EPG_aio_1397_duration, EPG_lenz_1397_duration,],}
all_data_EPG_1397_operators=pd.DataFrame(all_data_EPG_1397_operators, columns=['operators', 'visit', 'duration'])

all_data_EPG_1397_operators=all_data_EPG_1397_operators.rename(columns={'operators': 'اپراتور', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

EPG_shabake_1_1397=all_data_EPG_1397.query("channels == 'شبکه 1'")
EPG_shabake_2_1397=all_data_EPG_1397.query("channels == 'شبکه 2'")
EPG_shabake_3_1397=all_data_EPG_1397.query("channels == 'شبکه 3'")
EPG_shabake_4_1397=all_data_EPG_1397.query("channels == 'شبکه 4'")
EPG_shabake_5_1397=all_data_EPG_1397.query("channels == 'شبکه 5'")
EPG_shabake_khabar_1397=all_data_EPG_1397.query("channels == 'خبر'")
EPG_shabake_ofogh_1397=all_data_EPG_1397.query("channels == 'افق'")
EPG_shabake_pooya_1397=all_data_EPG_1397.query("channels == 'پویا'")
EPG_shabake_omid_1397=all_data_EPG_1397.query("channels == 'امید'")
EPG_shabake_ifilm_1397=all_data_EPG_1397.query("channels == 'آی فیلم'")
EPG_shabake_namayesh_1397=all_data_EPG_1397.query("channels == 'نمایش'")
EPG_shabake_tamasha_1397=all_data_EPG_1397.query("channels == 'تماشا'")
EPG_shabake_mostanad_1397=all_data_EPG_1397.query("channels == 'مستند'")
EPG_shabake_shoma_1397=all_data_EPG_1397.query("channels == 'شما'")
EPG_shabake_amozesh_1397=all_data_EPG_1397.query("channels == 'آموزش'")
EPG_shabake_varzesh_1397=all_data_EPG_1397.query("channels == 'ورزش'")
EPG_shabake_nasim_1397=all_data_EPG_1397.query("channels == 'نسیم'")
EPG_shabake_qoran_1397=all_data_EPG_1397.query("channels == 'قرآن'")
EPG_shabake_salamat_1397=all_data_EPG_1397.query("channels == 'سلامت'")
EPG_shabake_irankala_1397=all_data_EPG_1397.query("channels == 'ایران کالا'")
EPG_shabake_alalam_1397=all_data_EPG_1397.query("channels == 'العالم'")
EPG_shabake_alkosar_1397=all_data_EPG_1397.query("channels == 'الکوثر'")
EPG_shabake_presstv_1397=all_data_EPG_1397.query("channels == 'پرس تی وی'")
EPG_shabake_sepehr_1397=all_data_EPG_1397.query("channels == 'سپهر'")
EPG_shabake_jamejam_1397=all_data_EPG_1397.query("channels == 'جام جم'")

EPG_shabake_1_1397_visit=EPG_shabake_1_1397['visit'].sum()
EPG_shabake_2_1397_visit=EPG_shabake_2_1397['visit'].sum()
EPG_shabake_3_1397_visit=EPG_shabake_3_1397['visit'].sum()
EPG_shabake_4_1397_visit=EPG_shabake_4_1397['visit'].sum()
EPG_shabake_5_1397_visit=EPG_shabake_5_1397['visit'].sum()
EPG_shabake_khabar_1397_visit=EPG_shabake_khabar_1397['visit'].sum()
EPG_shabake_ofogh_1397_visit=EPG_shabake_ofogh_1397['visit'].sum()
EPG_shabake_pooya_1397_visit=EPG_shabake_pooya_1397['visit'].sum()
EPG_shabake_omid_1397_visit=EPG_shabake_omid_1397['visit'].sum()
EPG_shabake_ifilm_1397_visit=EPG_shabake_ifilm_1397['visit'].sum()
EPG_shabake_namayesh_1397_visit=EPG_shabake_namayesh_1397['visit'].sum()
EPG_shabake_tamasha_1397_visit=EPG_shabake_tamasha_1397['visit'].sum()
EPG_shabake_mostanad_1397_visit=EPG_shabake_mostanad_1397['visit'].sum()
EPG_shabake_shoma_1397_visit=EPG_shabake_shoma_1397['visit'].sum()
EPG_shabake_amozesh_1397_visit=EPG_shabake_amozesh_1397['visit'].sum()
EPG_shabake_varzesh_1397_visit=EPG_shabake_varzesh_1397['visit'].sum()
EPG_shabake_nasim_1397_visit=EPG_shabake_nasim_1397['visit'].sum()
EPG_shabake_qoran_1397_visit=EPG_shabake_qoran_1397['visit'].sum()
EPG_shabake_salamat_1397_visit=EPG_shabake_salamat_1397['visit'].sum()
EPG_shabake_irankala_1397_visit=EPG_shabake_irankala_1397['visit'].sum()
EPG_shabake_alalam_1397_visit=EPG_shabake_alalam_1397['visit'].sum()
EPG_shabake_alkosar_1397_visit=EPG_shabake_alkosar_1397['visit'].sum()
EPG_shabake_presstv_1397_visit=EPG_shabake_presstv_1397['visit'].sum()
EPG_shabake_sepehr_1397_visit=EPG_shabake_sepehr_1397['visit'].sum()
EPG_shabake_jamejam_1397_visit=EPG_shabake_jamejam_1397['visit'].sum()

EPG_shabake_1_1397_duration=EPG_shabake_1_1397['duration'].sum()
EPG_shabake_2_1397_duration=EPG_shabake_2_1397['duration'].sum()
EPG_shabake_3_1397_duration=EPG_shabake_3_1397['duration'].sum()
EPG_shabake_4_1397_duration=EPG_shabake_4_1397['duration'].sum()
EPG_shabake_5_1397_duration=EPG_shabake_5_1397['duration'].sum()
EPG_shabake_khabar_1397_duration=EPG_shabake_khabar_1397['duration'].sum()
EPG_shabake_ofogh_1397_duration=EPG_shabake_ofogh_1397['duration'].sum()
EPG_shabake_pooya_1397_duration=EPG_shabake_pooya_1397['duration'].sum()
EPG_shabake_omid_1397_duration=EPG_shabake_omid_1397['duration'].sum()
EPG_shabake_ifilm_1397_duration=EPG_shabake_ifilm_1397['duration'].sum()
EPG_shabake_namayesh_1397_duration=EPG_shabake_namayesh_1397['duration'].sum()
EPG_shabake_tamasha_1397_duration=EPG_shabake_tamasha_1397['duration'].sum()
EPG_shabake_mostanad_1397_duration=EPG_shabake_mostanad_1397['duration'].sum()
EPG_shabake_shoma_1397_duration=EPG_shabake_shoma_1397['duration'].sum()
EPG_shabake_amozesh_1397_duration=EPG_shabake_amozesh_1397['duration'].sum()
EPG_shabake_varzesh_1397_duration=EPG_shabake_varzesh_1397['duration'].sum()
EPG_shabake_nasim_1397_duration=EPG_shabake_nasim_1397['duration'].sum()
EPG_shabake_qoran_1397_duration=EPG_shabake_qoran_1397['duration'].sum()
EPG_shabake_salamat_1397_duration=EPG_shabake_salamat_1397['duration'].sum()
EPG_shabake_irankala_1397_duration=EPG_shabake_irankala_1397['duration'].sum()
EPG_shabake_alalam_1397_duration=EPG_shabake_alalam_1397['duration'].sum()
EPG_shabake_alkosar_1397_duration=EPG_shabake_alkosar_1397['duration'].sum()
EPG_shabake_presstv_1397_duration=EPG_shabake_presstv_1397['duration'].sum()
EPG_shabake_sepehr_1397_duration=EPG_shabake_sepehr_1397['duration'].sum()
EPG_shabake_jamejam_1397_duration=EPG_shabake_jamejam_1397['duration'].sum()

all_data_1397_channels=pd.DataFrame()
all_data_1397_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                     'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                     'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                     'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                     'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
       'visit': [EPG_shabake_1_1397_visit, EPG_shabake_2_1397_visit, EPG_shabake_3_1397_visit,
                 EPG_shabake_4_1397_visit, EPG_shabake_5_1397_visit, EPG_shabake_khabar_1397_visit,
                 EPG_shabake_ofogh_1397_visit, EPG_shabake_pooya_1397_visit, EPG_shabake_omid_1397_visit,
                 EPG_shabake_ifilm_1397_visit, EPG_shabake_namayesh_1397_visit, EPG_shabake_tamasha_1397_visit,
                 EPG_shabake_mostanad_1397_visit, EPG_shabake_shoma_1397_visit, EPG_shabake_amozesh_1397_visit,
                 EPG_shabake_varzesh_1397_visit, EPG_shabake_nasim_1397_visit, EPG_shabake_qoran_1397_visit,
                 EPG_shabake_salamat_1397_visit, EPG_shabake_irankala_1397_visit, EPG_shabake_alalam_1397_visit,
                 EPG_shabake_alkosar_1397_visit, EPG_shabake_presstv_1397_visit, EPG_shabake_sepehr_1397_visit,],
        'duration': [EPG_shabake_1_1397_duration, EPG_shabake_2_1397_duration, EPG_shabake_3_1397_duration,
                 EPG_shabake_4_1397_duration, EPG_shabake_5_1397_duration, EPG_shabake_khabar_1397_duration,
                 EPG_shabake_ofogh_1397_duration, EPG_shabake_pooya_1397_duration, EPG_shabake_omid_1397_duration,
                 EPG_shabake_ifilm_1397_duration, EPG_shabake_namayesh_1397_duration, EPG_shabake_tamasha_1397_duration,
                 EPG_shabake_mostanad_1397_duration, EPG_shabake_shoma_1397_duration, EPG_shabake_amozesh_1397_duration,
                 EPG_shabake_varzesh_1397_duration, EPG_shabake_nasim_1397_duration, EPG_shabake_qoran_1397_duration,
                 EPG_shabake_salamat_1397_duration, EPG_shabake_irankala_1397_duration, EPG_shabake_alalam_1397_duration,
                 EPG_shabake_alkosar_1397_duration, EPG_shabake_presstv_1397_duration, EPG_shabake_sepehr_1397_duration,],}
all_data_1397_channels=pd.DataFrame(all_data_1397_channels, columns=['channels', 'visit', 'duration'])

all_data_1397_channels=all_data_1397_channels.rename(columns={'channels': 'نام شبکه', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})

all_data_EPG_1397_summary=pd.DataFrame()
all_data_EPG_1397_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد کاربران ثبت نامی',],
       'statistics': [all_data_EPG_1397_total_visit, all_data_EPG_1397_total_duration, all_data_EPG_1397_total_register_user,],}
all_data_EPG_1397_summary=pd.DataFrame(all_data_EPG_1397_summary, columns=['parameters', 'statistics'])

all_data_EPG_1397_summary=all_data_EPG_1397_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})

writer = pd.ExcelWriter('output/خلاصه آمار سال 1397.xlsx', engine='xlsxwriter')
all_data_1397_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
all_data_EPG_1397_operators.to_excel(writer, 'آمار اپراتورها')
all_data_EPG_1397_summary.to_excel(writer, 'خلاصه سال 1397')
writer.save()







   ########################### General Data #############################

   ########################### popular contents #############################
print("popular contents") 
sima_popular_content=sima.copy()
sima_popular_content=sima_popular_content.groupby(['نام برنامه','channel']).sum().reset_index()
sima_popular_content.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
sima_popular_content_visit=sima_popular_content.iloc[0:10 , [0, 3]]
sima_popular_content_visit.to_excel('busy/sima_popular_content_visit.xlsx')
sima_popular_content_visit=pd.read_excel('busy/sima_popular_content_visit.xlsx')
del sima_popular_content_visit['Unnamed: 0']

sima_popular_content.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
sima_popular_content_duration=sima_popular_content.iloc[0:10 , [0, 2]]
sima_popular_content_duration.to_excel('busy/sima_popular_content_duration.xlsx')
sima_popular_content_duration=pd.read_excel('busy/sima_popular_content_duration.xlsx')
del sima_popular_content_duration['Unnamed: 0']

sima_popular_content=pd.DataFrame()
sima_popular_content=pd.concat([sima_popular_content_visit, sima_popular_content_duration], axis=1)

   ########################### Time visit #############################
print("Time visit") 


all_data_Time=all_data_Time.groupby(['ساعت']).sum().reset_index()
del all_data_Time['ردیف']
del all_data_Time['تاریخ']
del all_data_Time['میانگین']
del all_data_Time['tag']

   ########################### Daily visit #############################

print("Daily visit") 

all_data_Daily=all_data.copy()
all_data_Daily=all_data_Daily.query("operator != 'تلوبیون'")
all_data_Daily.insert(15, 'day', '')

date_start_time=[]
time_start_time=[]
month_start_time=[]
day_start_time=[]
hour_start_time=[]
jajli_date=[]
jajli_date_year=[]
jajli_date_month=[]
jajli_date_day=[]

for i in all_data_Daily.iloc[:,3]:   
    date_obj = dt.strptime(str(i), '%Y-%m-%d  %H:%M:%S')
    date_start_time.append(dt.strftime(date_obj,'%m/%d/%Y'))
    time_start_time.append(dt.strftime(date_obj,'%I:%M:%S'))
    month_start_time.append(dt.strftime(date_obj,'%m'))
    day_start_time.append(dt.strftime(date_obj,'%d'))
    hour_start_time.append(dt.strftime(date_obj,'%H'))
    jajli_date.append(dt.strftime(date_obj,'%Y,%m,%d'))

for j in jajli_date:
    jajli_date_year.append(jalali.Gregorian(j).persian_string("{0}"))
    jajli_date_month.append(jalali.Gregorian(j).persian_string("{1}"))
    jajli_date_day.append(jalali.Gregorian(j).persian_string("{2}"))

#print(jalali.Gregorian(jajli_date[0]).persian_string("{0}"))
#print(jalali.Gregorian(jajli_date[0]).persian_string("{1}"))
#print(jalali.Gregorian(jajli_date[0]).persian_string("{2}"))

length_all_data_Daily=len(all_data_Daily['تاریخ شروع'])
for i in range(0, length_all_data_Daily):
    all_data_Daily.loc[i, 'day']=jalali.Gregorian(jajli_date[i]).persian_string("{2}")

#all_data_Daily['day'] = all_data_Daily['day'].astype(int)
all_data_Daily.sort_values('day', axis = 0, ascending = False, inplace = True, na_position ='last')
all_data_Daily=all_data_Daily.groupby(['day']).sum().reset_index()

del all_data_Daily['ردیف']
del all_data_Daily['تاریخ']
del all_data_Daily['میانگین']
del all_data_Daily['tag']
del all_data_Daily['ساعت']
















print("convert to excel")





























