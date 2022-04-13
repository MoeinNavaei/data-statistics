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
import time
import openpyxl

start = time.time()
########################## PROGRAM START ##########################

########################## Get Data ##########################
print("get data")
all_data=pd.read_excel(r'E:\hard\report\1400\Esfand\input\EPG Esfand 1400.xlsx')
site_channels_primary=pd.read_excel(r'E:\hard\report\1400\Esfand\input\سایت شبکه ها.xlsx')
RegisterActiveUsers=pd.read_excel(r'E:\hard\report\1400\Esfand\input\RegisterActiveUsers.xlsx')
vod_lenz=pd.read_csv(r'E:\hard\report\1400\Esfand\input\LenzEsfand1400.csv')
vod_tva=pd.read_excel(r'E:\hard\report\1400\Esfand\input\TvaEsfand1400.xlsx', sheet_name='Videos')
vod_aio=pd.read_excel(r'E:\hard\report\1400\Esfand\input\AioEsfand1400.xlsx')
########################## Sepehr $ Channels Site ##########################
print("SepehrChannelsSite")
from SepehrChannelsSite import *
site_channels_final = SepehrChannelsSite(site_channels_primary)
all_data=all_data.append([site_channels_final])
all_data = all_data.reset_index()
del all_data['index']
########################## convert hour time to minute time ##########################
print("duration")
all_data['مدت بازدید']=all_data['مدت بازدید'] * 60   # convert hour to minute
########################## change columns name ##########################
all_data=all_data.rename(columns={"نوع":"type"})
all_data=all_data.rename(columns={"اپراتور":"operator"})
all_data=all_data.rename(columns={"نام شبکه":"channel"})
all_data=all_data.rename(columns={"جنس":"content_type"})
########################## Unify Channel Names ##########################
all_data['channel'] = all_data['channel'].str.replace('قران', 'قرآن')
all_data['channel'] = all_data['channel'].str.replace('جام جم 1', 'جام جم')
all_data['channel'] = all_data['channel'].str.replace('ي', 'ی')
all_data['channel'] = all_data['channel'].str.replace('ؤ','و')
all_data['channel'] = all_data['channel'].str.replace('ك','ک')
all_data['نام برنامه'] = all_data['نام برنامه'].str.replace('ي', 'ی')
all_data['نام برنامه'] = all_data['نام برنامه'].str.replace('ؤ','و')
all_data['نام برنامه'] = all_data['نام برنامه'].str.replace('ك','ک')
all_data['type'] = all_data['type'].str.replace('ي', 'ی')
all_data['type'] = all_data['type'].str.replace('ؤ','و')
all_data['type'] = all_data['type'].str.replace('ك','ک')
all_data['operator'] = all_data['operator'].str.replace('ي', 'ی')
all_data['operator'] = all_data['operator'].str.replace('ؤ','و')
all_data['operator'] = all_data['operator'].str.replace('ك','ک')
########################## distinct of services ##########################
print("query of type")
sima=all_data.query("type == 'سراسری'")
radio=all_data.query("type == 'رادیویی'")
ostani=all_data.query("type == 'استانی'")
ekhtesasi=all_data.query("type == 'اختصاصی'")
boronmarzi=all_data.query("type == 'برون مرزی'")
all_data_Time=all_data.copy()
all_data_Time=all_data_Time.query("operator != 'تلوبیون'")
########################## input of alexa data ##########################
#sites_name_alexa = pd.read_excel(r'D:\hard\report\1400\ordibehesht\input\sites_name_alexa.xlsx')
########################## input of 1397 data ##########################
print("get data 1397")
EPG_2=pd.read_excel('EPG/EPG 1397/EPG 2.xlsx', sheet_name='ALL')  
EPG_3=pd.read_excel('EPG/EPG 1397/EPG 3.xlsx', sheet_name='کل') 
EPG_4=pd.read_excel('EPG/EPG 1397/EPG 4.xlsx', sheet_name='کل') 
EPG_5=pd.read_excel('EPG/EPG 1397/EPG 5.xlsx', sheet_name='کل') 
EPG_esfand=pd.read_excel('EPG/EPG 1397/EPG Esfand.xlsx', sheet_name='کل')

EPG_1397_sima=pd.DataFrame()
EPG_1397_sima=EPG_1397_sima.append([EPG_3, EPG_4, EPG_5, EPG_esfand])

print("get data 1397 ekhtesasi")
EPG_3_ekhtesasi=pd.read_excel('EPG/EPG 1397/EPG 3.xlsx', sheet_name='اختصاصی') 
EPG_4_ekhtesasi=pd.read_excel('EPG/EPG 1397/EPG 4.xlsx', sheet_name='اختصاصی') 
EPG_5_ekhtesasi=pd.read_excel('EPG/EPG 1397/EPG 5.xlsx', sheet_name='اختصاصی') 
EPG_esfand_ekhtesasi=pd.read_excel('EPG/EPG 1397/EPG Esfand.xlsx', sheet_name='اختصاصی')

EPG_1397_ekhtesasi=pd.DataFrame()
EPG_1397_ekhtesasi=EPG_1397_ekhtesasi.append([EPG_3_ekhtesasi, EPG_4_ekhtesasi, EPG_5_ekhtesasi, EPG_esfand_ekhtesasi])
########################## input of 1398 data ##########################
print("get data 1398")
EPG_Farvardin_1398=pd.read_excel('EPG/EPG 1398/EPG Farvardin 1398.xlsx', sheet_name='آمار')
EPG_Ordibehesht_1398=pd.read_excel('EPG/EPG 1398/EPG Ordibehesht 1398.xlsx', sheet_name='آمار')
EPG_Khordad_1398=pd.read_excel('EPG/EPG 1398/EPG Khordad 1398.xlsx', sheet_name='آمار')
EPG_Tir_1398=pd.read_excel('EPG/EPG 1398/EPG Tir 1398.xlsx', sheet_name='آمار')
EPG_Mordad_1398=pd.read_excel('EPG/EPG 1398/EPG Mordad 1398.xlsx', sheet_name='آمار')
EPG_Shahrivar_1398=pd.read_excel('EPG/EPG 1398/EPG Shahrivar 1398.xlsx', sheet_name='آمار')
EPG_Mehr_1398=pd.read_excel('EPG/EPG 1398/EPG Mehr 1398.xlsx', sheet_name='آمار')
EPG_Aban_1398=pd.read_excel('EPG/EPG 1398/EPG Aban 1398.xlsx', sheet_name='آمار')
EPG_Azar_1398=pd.read_excel('EPG/EPG 1398/EPG Azar 1398.xlsx', sheet_name='آمار')
EPG_Dey_1398=pd.read_excel('EPG/EPG 1398/EPG Dey 1398.xlsx', sheet_name='آمار')
EPG_Bahman_1398=pd.read_excel('EPG/EPG 1398/EPG Bahman 1398.xlsx', sheet_name='آمار')
EPG_Esfand_1398=pd.read_excel('EPG/EPG 1398/EPG Esfand 1398.xlsx', sheet_name='آمار')
print("get data 1398 ekhtesasi")
EPG_Farvardin_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Farvardin 1398.xlsx', sheet_name='اختصاصی بدون قسمت')
EPG_Ordibehesht_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Ordibehesht 1398.xlsx', sheet_name='اختصاصی بدون قسمت')
EPG_Khordad_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Khordad 1398.xlsx', sheet_name='اختصاصی بدون قسمت')
EPG_Tir_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Tir 1398.xlsx', sheet_name='اختصاصی بدون قسمت')
EPG_Mordad_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Mordad 1398.xlsx', sheet_name='اختصاصی بدون قسمت')
EPG_Shahrivar_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Shahrivar 1398.xlsx', sheet_name='اختصاصی بدون قسمت')
EPG_Mehr_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Mehr 1398.xlsx', sheet_name='اختصاصی بدون قسمت')
EPG_Aban_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Aban 1398.xlsx', sheet_name='اختصاصی')
EPG_Azar_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Azar 1398.xlsx', sheet_name='اختصاصی')
EPG_Dey_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Dey 1398.xlsx', sheet_name='اختصاصی')
EPG_Bahman_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Bahman 1398.xlsx', sheet_name='اختصاصی')
EPG_Esfand_1398_ekhtesasi=pd.read_excel('EPG/EPG 1398/EPG Esfand 1398.xlsx', sheet_name='اختصاصی')

EPG_1398_ekhtesasi=pd.DataFrame()
EPG_1398_ekhtesasi=EPG_1398_ekhtesasi.append([EPG_Farvardin_1398_ekhtesasi, EPG_Ordibehesht_1398_ekhtesasi, EPG_Khordad_1398_ekhtesasi,
                                              EPG_Tir_1398_ekhtesasi, EPG_Mordad_1398_ekhtesasi, EPG_Shahrivar_1398_ekhtesasi,
                                              EPG_Mehr_1398_ekhtesasi, EPG_Aban_1398_ekhtesasi, EPG_Azar_1398_ekhtesasi,
                                              EPG_Dey_1398_ekhtesasi, EPG_Bahman_1398_ekhtesasi, EPG_Esfand_1398_ekhtesasi,])
########################## input of 1399 data ##########################
print("get data 1399")
EPG_Farvardin_1399=pd.read_excel('EPG/EPG 1399/EPG Farvardin 1399.xlsx', sheet_name='آمار')
EPG_Ordibehesht_1399=pd.read_excel('EPG/EPG 1399/EPG Ordibehesht 1399.xlsx', sheet_name='آمار')
EPG_Khordad_1399=pd.read_excel('EPG/EPG 1399/EPG Khordad 1399.xlsx', sheet_name='آمار')
EPG_Tir_1399=pd.read_excel('EPG/EPG 1399/EPG Tir 1399.xlsx', sheet_name='آمار')
EPG_Mordad_1399=pd.read_excel('EPG/EPG 1399/EPG Mordad 1399.xlsx', sheet_name='آمار')
EPG_Shahrivar_1399=pd.read_excel('EPG/EPG 1399/EPG Shahrivar 1399.xlsx', sheet_name='آمار')
EPG_Mehr_1399=pd.read_excel('EPG/EPG 1399/EPG Mehr 1399.xlsx', sheet_name='آمار')
EPG_Aban_1399=pd.read_excel('EPG/EPG 1399/EPG Aban 1399.xlsx', sheet_name='آمار')
EPG_Azar_1399=pd.read_excel('EPG/EPG 1399/EPG Azar 1399.xlsx', sheet_name='آمار')
EPG_Dey_1399=pd.read_excel('EPG/EPG 1399/EPG Dey 1399.xlsx', sheet_name='آمار')
EPG_Bahman_1399=pd.read_excel('EPG/EPG 1399/EPG Bahman 1399.xlsx', sheet_name='آمار')
EPG_Esfand_1399=pd.read_excel('EPG/EPG 1399/EPG Esfand 1399.xlsx', sheet_name='آمار')
print("get data 1399 ekhtesasi")
EPG_Farvardin_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Farvardin 1399.xlsx', sheet_name='اختصاصی')
EPG_Ordibehesht_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Ordibehesht 1399.xlsx', sheet_name='اختصاصی')
EPG_Khordad_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Khordad 1399.xlsx', sheet_name='اختصاصی')
EPG_Tir_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Tir 1399.xlsx', sheet_name='اختصاصی')
EPG_Mordad_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Mordad 1399.xlsx', sheet_name='اختصاصی')
EPG_Shahrivar_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Shahrivar 1399.xlsx', sheet_name='اختصاصی')
EPG_Mehr_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Mehr 1399.xlsx', sheet_name='اختصاصی')
EPG_Aban_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Aban 1399.xlsx', sheet_name='اختصاصی')
EPG_Azar_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Azar 1399.xlsx', sheet_name='اختصاصی')
EPG_Dey_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Dey 1399.xlsx', sheet_name='اختصاصی')
EPG_Bahman_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Bahman 1399.xlsx', sheet_name='اختصاصی')
EPG_Esfand_1399_ekhtesasi=pd.read_excel('EPG/EPG 1399/EPG Esfand 1399.xlsx', sheet_name='اختصاصی')

EPG_1399_ekhtesasi=pd.DataFrame()
EPG_1399_ekhtesasi=EPG_1399_ekhtesasi.append([EPG_Farvardin_1399_ekhtesasi, EPG_Ordibehesht_1399_ekhtesasi, EPG_Khordad_1399_ekhtesasi,
                                              EPG_Tir_1399_ekhtesasi, EPG_Mordad_1399_ekhtesasi, EPG_Shahrivar_1399_ekhtesasi,
                                              EPG_Mehr_1399_ekhtesasi, EPG_Aban_1399_ekhtesasi, EPG_Azar_1399_ekhtesasi,
                                              EPG_Dey_1399_ekhtesasi, EPG_Bahman_1399_ekhtesasi, EPG_Bahman_1399_ekhtesasi,])
########################## Run of Functions ##########################
from EPG_1397 import * 
from EPG_1398 import *
from EPG_1399 import *
from EPG_1400 import *
from summary import *
from sima_data import *
from radio_data import *
from ostani_data import *
from ekhtesasi_data import *
from boronmarzi_data import *
from Daily_Time import *
from popular import *
from vod import *
from content_type import *
#from alexa import *
from total_all_data import *

########################## Run of EPG_1397 Function ##########################
print("EPG_1397")
[all_data_EPG_1397_operators, all_data_1397_channels, all_data_EPG_1397_summary]=EPG_1397(EPG_2, EPG_3, EPG_4, EPG_5, EPG_esfand)
########################## Run of EPG_1398 Function ##########################
print("EPG_1398")
[Farvardin_1398_sima_visit_channels, Farvardin_1398_operator_data, Farvardin_1398_all_data_summary, \
Ordibehesht_1398_sima_visit_channels, Ordibehesht_1398_operator_data, Ordibehesht_1398_all_data_summary, \
Khordad_1398_sima_visit_channels, Khordad_1398_operator_data, Khordad_1398_all_data_summary, \
Tir_1398_sima_visit_channels, Tir_1398_operator_data, Tir_1398_all_data_summary, \
Mordad_1398_sima_visit_channels, Mordad_1398_operator_data, Mordad_1398_all_data_summary, \
Shahrivar_1398_sima_visit_channels, Shahrivar_1398_operator_data, Shahrivar_1398_all_data_summary, \
Mehr_1398_sima_visit_channels, Mehr_1398_operator_data, Mehr_1398_all_data_summary, \
Aban_1398_sima_visit_channels, Aban_1398_operator_data, Aban_1398_all_data_summary, \
Azar_1398_sima_visit_channels, Azar_1398_operator_data, Azar_1398_all_data_summary, \
Dey_1398_sima_visit_channels, Dey_1398_operator_data, Dey_1398_all_data_summary, \
Bahman_1398_sima_visit_channels, Bahman_1398_operator_data, Bahman_1398_all_data_summary, \
Esfand_1398_sima_visit_channels, Esfand_1398_operator_data, Esfand_1398_all_data_summary]=EPG_1398(EPG_Farvardin_1398, EPG_Ordibehesht_1398, EPG_Khordad_1398, \
EPG_Tir_1398, EPG_Mordad_1398, EPG_Shahrivar_1398, \
EPG_Mehr_1398, EPG_Aban_1398, EPG_Azar_1398, \
EPG_Dey_1398, EPG_Bahman_1398, EPG_Esfand_1398)
########################## Run of EPG_1399 Function ##########################
print("EPG_1399")
[Farvardin_1399_sima_visit_channels, Farvardin_1399_operator_data, Farvardin_1399_all_data_summary, \
Ordibehesht_1399_sima_visit_channels, Ordibehesht_1399_operator_data, Ordibehesht_1399_all_data_summary, \
Khordad_1399_sima_visit_channels, Khordad_1399_operator_data, Khordad_1399_all_data_summary, \
Tir_1399_sima_visit_channels, Tir_1399_operator_data, Tir_1399_all_data_summary, \
Mordad_1399_sima_visit_channels, Mordad_1399_operator_data, Mordad_1399_all_data_summary, \
Shahrivar_1399_sima_visit_channels, Shahrivar_1399_operator_data, Shahrivar_1399_all_data_summary, \
Mehr_1399_sima_visit_channels, Mehr_1399_operator_data, Mehr_1399_all_data_summary, \
Aban_1399_sima_visit_channels, Aban_1399_operator_data, Aban_1399_all_data_summary, \
Azar_1399_sima_visit_channels, Azar_1399_operator_data, Azar_1399_all_data_summary, \
Dey_1399_sima_visit_channels, Dey_1399_operator_data, Dey_1399_all_data_summary, \
Bahman_1399_sima_visit_channels, Bahman_1399_operator_data, Bahman_1399_all_data_summary, \
Esfand_1399_sima_visit_channels, Esfand_1399_operator_data, Esfand_1399_all_data_summary]=EPG_1399(EPG_Farvardin_1399, EPG_Ordibehesht_1399, EPG_Khordad_1399, \
EPG_Tir_1399, EPG_Mordad_1399, EPG_Shahrivar_1399, \
EPG_Mehr_1399, EPG_Aban_1399, EPG_Azar_1399, \
EPG_Dey_1399, EPG_Bahman_1399, EPG_Esfand_1399)
########################## Run of EPG_1400 Function ##########################
print("EPG_1400")
[current_month_sima_visit_channels, current_month_operator_data, current_month_all_data_summary]=EPG_1400(all_data, sima, RegisterActiveUsers)
########################## Run of summary Function ##########################
print("summary")
[data_summary_service, data_summary_operator, data_summary_service_operator]=summary(all_data)
########################## Run of sima Function ##########################
print("sima")
[sima_channels_statistics, sima_channels_popular_content, sima_channels_Time_visit]=sima_data(sima, all_data_Time)
########################## Run of radio Function ##########################
print("radio")
[radio_channels_statistics, radio_channels_popular_content]=radio_data(radio)
########################## Run of ostani Function ##########################
print("ostani")
[ostani_channels_statistics, ostani_channels_popular_content]=ostani_data(ostani)
########################## Run of ekhtesasi Function ##########################
print("ekhtesasi")
[ekhtesasi_channels_statistics, ekhtesasi_channels_popular_content]=ekhtesasi_data(ekhtesasi)
########################## Run of boronmarzi Function ##########################
print("boronmarzi")
[boronmarzi_channels_statistics, boronmarzi_channels_popular_content]=boronmarzi_data(boronmarzi)
########################## Run of Daily_Time Function ##########################
print("Daily_Time")
[all_data_Time_statistics_sima, all_data_Daily_statistics_sima, \
all_data_Time_statistics_radio, all_data_Daily_statistics_radio, \
all_data_Time_statistics_ostani, all_data_Daily_statistics_ostani, \
all_data_Time_statistics_ekhtesasi, all_data_Daily_statistics_ekhtesasi, \
all_data_Time_statistics_boronmarzi, all_data_Daily_statistics_boronmarzi]=Daily_Time(all_data, all_data_Time)
########################## Run of popular Function ##########################
print("popular")
[sima_popular_content, ekhtesasi_popular_content, ostani_popular_content, radio_popular_content, boronmarzi_popular_content]=popular(sima, ekhtesasi, ostani, radio, boronmarzi)
########################## Run of vod Function ##########################
print("vod")
[vod_statistics_summary, vod_popular_content]=vod(vod_tva, vod_lenz, vod_aio)
########################## Run of content_type Function ##########################
print("content_type")
[Akhbar_statistics, Akhbar_popular_visit, Akhbar_popular_duration, \
FilmSinamaei_statistics, FilmSinamaei_popular_visit, FilmSinamaei_popular_duration, \
Kodak_statistics, Kodak_popular_visit, Kodak_popular_duration, \
MajmoeTV_statistics, MajmoeTV_popular_visit, MajmoeTV_popular_duration, \
Mostanad_statistics, Mostanad_popular_visit, Mostanad_popular_duration, \
Varzeshi_statistics, Varzeshi_popular_visit, Varzeshi_popular_duration]=content_type(sima)
########################## Run of alexa Function ##########################
print("alexa")
#alexa_data=alexa(sites_name_alexa)
########################## input of 1400 data ##########################
EPG_Farvardin_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Farvardin_1400_first['Unnamed: 0']
EPG_Farvardin_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Farvardin_1400_second['Unnamed: 0']
EPG_Farvardin_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Farvardin_1400_third['Unnamed: 0']
RegisterActiveUsers_Farvardin_1400=pd.read_excel(r'E:\hard\report\1400\farvardin\input\RegisterActiveUsers.xlsx')

EPG_Ordibehesht_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Ordibehesht_1400_first['Unnamed: 0']
EPG_Ordibehesht_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Ordibehesht_1400_second['Unnamed: 0']
EPG_Ordibehesht_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Ordibehesht_1400_third['Unnamed: 0']
RegisterActiveUsers_Ordibehesht_1400=pd.read_excel(r'E:\hard\report\1400\ordibehesht\input\RegisterActiveUsers.xlsx')

EPG_Khordad_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Khordad_1400_first['Unnamed: 0']
EPG_Khordad_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Khordad_1400_second['Unnamed: 0']
EPG_Khordad_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Khordad_1400_third['Unnamed: 0']
RegisterActiveUsers_Khordad_1400=pd.read_excel(r'E:\hard\report\1400\Khordad\input\RegisterActiveUsers.xlsx')

EPG_Tir_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Tir_1400_first['Unnamed: 0']
EPG_Tir_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Tir_1400_second['Unnamed: 0']
EPG_Tir_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Tir_1400_third['Unnamed: 0']
RegisterActiveUsers_Tir_1400=pd.read_excel(r'E:\hard\report\1400\Tir\input\RegisterActiveUsers.xlsx')

EPG_Mordad_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Mordad_1400_first['Unnamed: 0']
EPG_Mordad_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Mordad_1400_second['Unnamed: 0']
EPG_Mordad_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Mordad_1400_third['Unnamed: 0']
RegisterActiveUsers_Mordad_1400=pd.read_excel(r'E:\hard\report\1400\Mordad\input\RegisterActiveUsers.xlsx')

EPG_Shahrivar_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Shahrivar_1400_first['Unnamed: 0']
EPG_Shahrivar_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Shahrivar_1400_second['Unnamed: 0']
EPG_Shahrivar_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Shahrivar_1400_third['Unnamed: 0']
RegisterActiveUsers_Shahrivar_1400=pd.read_excel(r'E:\hard\report\1400\Shahrivar\input\RegisterActiveUsers.xlsx')

EPG_Mehr_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Mehr_1400_first['Unnamed: 0']
EPG_Mehr_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Mehr_1400_second['Unnamed: 0']
EPG_Mehr_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Mehr_1400_third['Unnamed: 0']
RegisterActiveUsers_Mehr_1400=pd.read_excel(r'E:\hard\report\1400\Mehr\input\RegisterActiveUsers.xlsx')

EPG_Aban_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Aban_1400_first['Unnamed: 0']
EPG_Aban_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Aban_1400_second['Unnamed: 0']
EPG_Aban_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Aban_1400_third['Unnamed: 0']
RegisterActiveUsers_Aban_1400=pd.read_excel(r'E:\hard\report\1400\Aban\input\RegisterActiveUsers.xlsx')

EPG_Azar_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Azar_1400_first['Unnamed: 0']
EPG_Azar_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Azar_1400_second['Unnamed: 0']
EPG_Azar_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Azar_1400_third['Unnamed: 0']
RegisterActiveUsers_Azar_1400=pd.read_excel(r'E:\hard\report\1400\Azar\input\RegisterActiveUsers.xlsx')

EPG_Dey_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Dey_1400_first['Unnamed: 0']
EPG_Dey_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Dey_1400_second['Unnamed: 0']
EPG_Dey_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Dey_1400_third['Unnamed: 0']
RegisterActiveUsers_Dey_1400=pd.read_excel(r'E:\hard\report\1400\Dey\input\RegisterActiveUsers.xlsx')

EPG_Bahman_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Bahman_1400_first['Unnamed: 0']
EPG_Bahman_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Bahman_1400_second['Unnamed: 0']
EPG_Bahman_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Bahman_1400_third['Unnamed: 0']
RegisterActiveUsers_Bahman_1400=pd.read_excel(r'E:\hard\report\1400\Bahman\input\RegisterActiveUsers.xlsx')

EPG_Esfand_1400_first=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار انواع سرویس های سازمان')
#del EPG_Esfand_1400_first['Unnamed: 0']
EPG_Esfand_1400_second=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='آمار اپراتورها')
#del EPG_Esfand_1400_second['Unnamed: 0']
EPG_Esfand_1400_third=pd.read_excel(r'E:\python codes\data statistics\output\output.sending.hard\خلاصه آمار.xlsx', sheet_name='اپراتورها و سرویسهای سازمان')
#del EPG_Esfand_1400_third['Unnamed: 0']
RegisterActiveUsers_Esfand_1400=pd.read_excel(r'E:\hard\report\1400\Esfand\input\RegisterActiveUsers.xlsx')
########################## Run of total_all_data Function ##########################
print("total_all_data")
[EPG_1397_total, EPG_1398_total, EPG_1399_total, EPG_1400_total]=total_all_data(EPG_1397_sima, EPG_1397_ekhtesasi, EPG_1398_ekhtesasi, EPG_1399_ekhtesasi, \
                   Farvardin_1398_all_data_summary, Ordibehesht_1398_all_data_summary, Khordad_1398_all_data_summary, \
                   Tir_1398_all_data_summary, Mordad_1398_all_data_summary, Shahrivar_1398_all_data_summary, \
                   Mehr_1398_all_data_summary, Aban_1398_all_data_summary, Azar_1398_all_data_summary, \
                   Dey_1398_all_data_summary, Bahman_1398_all_data_summary, Esfand_1398_all_data_summary, \
                   Farvardin_1399_all_data_summary, Ordibehesht_1399_all_data_summary, Khordad_1399_all_data_summary, 
                   Tir_1399_all_data_summary, Mordad_1399_all_data_summary, Shahrivar_1399_all_data_summary, 
                   Mehr_1399_all_data_summary, Aban_1399_all_data_summary, Azar_1399_all_data_summary, 
                   Dey_1399_all_data_summary, Bahman_1399_all_data_summary, Esfand_1399_all_data_summary,
                   EPG_Farvardin_1400_first, EPG_Farvardin_1400_second, EPG_Farvardin_1400_third, RegisterActiveUsers_Farvardin_1400,
                   EPG_Ordibehesht_1400_first, EPG_Ordibehesht_1400_second, EPG_Ordibehesht_1400_third, RegisterActiveUsers_Ordibehesht_1400,
                   EPG_Khordad_1400_first, EPG_Khordad_1400_second, EPG_Khordad_1400_third, RegisterActiveUsers_Khordad_1400,
                   EPG_Tir_1400_first, EPG_Tir_1400_second, EPG_Tir_1400_third, RegisterActiveUsers_Tir_1400,
                   EPG_Mordad_1400_first, EPG_Mordad_1400_second, EPG_Mordad_1400_third, RegisterActiveUsers_Mordad_1400,
                   EPG_Shahrivar_1400_first, EPG_Shahrivar_1400_second, EPG_Shahrivar_1400_third, RegisterActiveUsers_Shahrivar_1400,
                   EPG_Mehr_1400_first, EPG_Mehr_1400_second, EPG_Mehr_1400_third, RegisterActiveUsers_Mehr_1400,
                   EPG_Aban_1400_first, EPG_Aban_1400_second, EPG_Aban_1400_third, RegisterActiveUsers_Aban_1400,
                   EPG_Azar_1400_first, EPG_Azar_1400_second, EPG_Azar_1400_third, RegisterActiveUsers_Azar_1400,
                   EPG_Dey_1400_first, EPG_Dey_1400_second, EPG_Dey_1400_third, RegisterActiveUsers_Dey_1400,
                   EPG_Bahman_1400_first, EPG_Bahman_1400_second, EPG_Bahman_1400_third, RegisterActiveUsers_Bahman_1400,
                   EPG_Esfand_1400_first, EPG_Esfand_1400_second, EPG_Esfand_1400_third, RegisterActiveUsers_Esfand_1400)




########################## PROGRAM END ##########################

print("--- %s seconds ---" % (time.time() - start))






