
def total_all_data(EPG_1397_sima, EPG_1397_ekhtesasi, EPG_1398_ekhtesasi, EPG_1399_ekhtesasi, \
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
                   EPG_Bahman_1400_first, EPG_Bahman_1400_second, EPG_Bahman_1400_third, RegisterActiveUsers_Bahman_1400):
    
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
    ##################################### year of 1397 #######################################
    print("sima 1397")
    
    EPG_1397_visit=EPG_1397_sima['تعداد بازدید'].sum()
    EPG_1397_duration=EPG_1397_sima['مدت بازدید'].sum()
    EPG_1397_channels_1=EPG_1397_sima.copy()
    EPG_1397_channels=EPG_1397_channels_1.groupby(['نام شبکه']).sum().reset_index()
    EPG_1397_channels=len(EPG_1397_channels['نام شبکه'])
    EPG_1397_contents_1=EPG_1397_sima.copy()
    EPG_1397_contents=EPG_1397_contents_1.groupby(['نام برنامه']).sum().reset_index()
    EPG_1397_contents=len(EPG_1397_contents['نام برنامه'])
    EPG_1397_operators_1=EPG_1397_sima.copy()
    EPG_1397_operators=EPG_1397_operators_1.groupby(['نام اپراتور']).sum().reset_index()
    EPG_1397_operators=len(EPG_1397_operators['نام اپراتور'])    
    
    print("ekhtesasi 1397")
    
    EPG_1397_ekhtesasi_visit=EPG_1397_ekhtesasi['تعداد بازدید'].sum()
    EPG_1397_ekhtesasi_duration=EPG_1397_ekhtesasi['مدت بازدید'].sum()
    EPG_1397_ekhtesasi_channels_1=EPG_1397_ekhtesasi.copy()
    EPG_1397_ekhtesasi_channels=EPG_1397_ekhtesasi_channels_1.groupby(['نام شبکه']).sum().reset_index()
    EPG_1397_ekhtesasi_channels=len(EPG_1397_ekhtesasi_channels['نام شبکه'])
    EPG_1397_ekhtesasi_contents_1=EPG_1397_ekhtesasi.copy()
    EPG_1397_ekhtesasi_contents=EPG_1397_ekhtesasi_contents_1.groupby(['نام برنامه']).sum().reset_index()
    EPG_1397_ekhtesasi_contents=len(EPG_1397_ekhtesasi_contents['نام برنامه'])
    EPG_1397_ekhtesasi_operators_1=EPG_1397_ekhtesasi.copy()
    EPG_1397_ekhtesasi_operators=EPG_1397_ekhtesasi_operators_1.groupby(['نام اپراتور']).sum().reset_index()
    EPG_1397_ekhtesasi_operators=len(EPG_1397_ekhtesasi_operators['نام اپراتور'])
    
    EPG_1397_total=pd.DataFrame()
    EPG_1397_total={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد اپراتور', 'تعداد شبکه',],
           'sima': [EPG_1397_contents, EPG_1397_visit, EPG_1397_duration, EPG_1397_ekhtesasi_operators, EPG_1397_ekhtesasi_channels,],
           'ekhtesasi': [EPG_1397_ekhtesasi_contents, EPG_1397_ekhtesasi_visit, EPG_1397_ekhtesasi_duration, EPG_1397_ekhtesasi_operators, EPG_1397_ekhtesasi_channels,],}
    EPG_1397_total=pd.DataFrame(EPG_1397_total, columns=['parameters', 'sima', 'ekhtesasi'])
    EPG_1397_total=EPG_1397_total.rename(columns={'parameters': 'پارامترها', 'sima': 'شبکه های سیما', 'ekhtesasi': 'شبکه های اختصاصی'})
    ######################################## year of 1398 ####################################
    print("sima 1398")
    
    EPG_1398_sima_visit=Farvardin_1398_all_data_summary.iat[0, 1]+Ordibehesht_1398_all_data_summary.iat[0, 1]+Khordad_1398_all_data_summary.iat[0, 1]+ \
                         Tir_1398_all_data_summary.iat[0, 1]+Mordad_1398_all_data_summary.iat[0, 1]+Shahrivar_1398_all_data_summary.iat[0, 1]+ \
                         Mehr_1398_all_data_summary.iat[0, 1]+Aban_1398_all_data_summary.iat[0, 1]+Azar_1398_all_data_summary.iat[0, 1]+ \
                         Dey_1398_all_data_summary.iat[0, 1]+Bahman_1398_all_data_summary.iat[0, 1]+Esfand_1398_all_data_summary.iat[0, 1]

    EPG_1398_sima_duration=Farvardin_1398_all_data_summary.iat[1, 1]+Ordibehesht_1398_all_data_summary.iat[1, 1]+Khordad_1398_all_data_summary.iat[1, 1]+ \
                         Tir_1398_all_data_summary.iat[1, 1]+Mordad_1398_all_data_summary.iat[1, 1]+Shahrivar_1398_all_data_summary.iat[1, 1]+ \
                         Mehr_1398_all_data_summary.iat[1, 1]+Aban_1398_all_data_summary.iat[1, 1]+Azar_1398_all_data_summary.iat[1, 1]+ \
                         Dey_1398_all_data_summary.iat[1, 1]+Bahman_1398_all_data_summary.iat[1, 1]+Esfand_1398_all_data_summary.iat[1, 1]

    EPG_1398_sima_register=Esfand_1398_all_data_summary.iat[3, 1]
    
    print("ekhtesasi 1398")
    
    EPG_1398_ekhtesasi_visit=EPG_1398_ekhtesasi['تعداد بازدید'].sum()
    EPG_1398_ekhtesasi_duration=EPG_1398_ekhtesasi['مدت بازدید'].sum()
    EPG_1398_ekhtesasi_channels_1=EPG_1398_ekhtesasi.copy()
    EPG_1398_ekhtesasi_channels=EPG_1398_ekhtesasi_channels_1.groupby(['نام شبکه']).sum().reset_index()
    EPG_1398_ekhtesasi_channels=len(EPG_1398_ekhtesasi_channels['نام شبکه'])
    EPG_1398_ekhtesasi_contents_1=EPG_1398_ekhtesasi.copy()
    EPG_1398_ekhtesasi_contents=EPG_1398_ekhtesasi_contents_1.groupby(['نام برنامه']).sum().reset_index()
    EPG_1398_ekhtesasi_contents=len(EPG_1398_ekhtesasi_contents['نام برنامه'])
    EPG_1398_ekhtesasi_operators_1=EPG_1398_ekhtesasi.copy()
    EPG_1398_ekhtesasi_operators=EPG_1398_ekhtesasi_operators_1.groupby(['نام اپراتور']).sum().reset_index()
    EPG_1398_ekhtesasi_operators=len(EPG_1398_ekhtesasi_operators['نام اپراتور'])
    
    EPG_1398_total=pd.DataFrame()
    EPG_1398_total={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد اپراتور', 'تعداد شبکه', 'کل کاربران ثبت نامی',],
           'sima': ["-", EPG_1398_sima_visit, EPG_1398_sima_duration, "9", "24", EPG_1398_sima_register,],
           'ekhtesasi': [EPG_1398_ekhtesasi_contents, EPG_1398_ekhtesasi_visit, EPG_1398_ekhtesasi_duration, EPG_1398_ekhtesasi_operators, EPG_1398_ekhtesasi_channels, "-",],}
    EPG_1398_total=pd.DataFrame(EPG_1398_total, columns=['parameters', 'sima', 'ekhtesasi'])
    EPG_1398_total=EPG_1398_total.rename(columns={'parameters': 'پارامترها', 'sima': 'شبکه های سیما', 'ekhtesasi': 'شبکه های اختصاصی'})
   ######################################## year of 1399 ####################################
    print("sima 1399")
    
    EPG_1399_sima_visit=Farvardin_1399_all_data_summary.iat[0, 1]+Ordibehesht_1399_all_data_summary.iat[0, 1]+Khordad_1399_all_data_summary.iat[0, 1]+ \
                         Tir_1399_all_data_summary.iat[0, 1]+Mordad_1399_all_data_summary.iat[0, 1]+Shahrivar_1399_all_data_summary.iat[0, 1]+ \
                         Mehr_1399_all_data_summary.iat[0, 1]+Aban_1399_all_data_summary.iat[0, 1]+Azar_1399_all_data_summary.iat[0, 1]+ \
                         Dey_1399_all_data_summary.iat[0, 1]

    EPG_1399_sima_duration=Farvardin_1399_all_data_summary.iat[1, 1]+Ordibehesht_1399_all_data_summary.iat[1, 1]+Khordad_1399_all_data_summary.iat[1, 1]+ \
                         Tir_1399_all_data_summary.iat[1, 1]+Mordad_1399_all_data_summary.iat[1, 1]+Shahrivar_1399_all_data_summary.iat[1, 1]+ \
                         Mehr_1399_all_data_summary.iat[1, 1]+Aban_1399_all_data_summary.iat[1, 1]+Azar_1399_all_data_summary.iat[1, 1]+ \
                         Dey_1399_all_data_summary.iat[1, 1]

    EPG_1399_sima_register=Dey_1399_all_data_summary.iat[3, 1]
    
    print("ekhtesasi 1399")
    
    EPG_1399_ekhtesasi_visit=EPG_1399_ekhtesasi['تعداد بازدید'].sum()
    EPG_1399_ekhtesasi_duration=EPG_1399_ekhtesasi['مدت بازدید'].sum()
    EPG_1399_ekhtesasi_channels_1=EPG_1399_ekhtesasi.copy()
    EPG_1399_ekhtesasi_channels=EPG_1399_ekhtesasi_channels_1.groupby(['نام شبکه']).sum().reset_index()
    EPG_1399_ekhtesasi_channels=len(EPG_1399_ekhtesasi_channels['نام شبکه'])
    EPG_1399_ekhtesasi_contents_1=EPG_1399_ekhtesasi.copy()
    EPG_1399_ekhtesasi_contents=EPG_1399_ekhtesasi_contents_1.groupby(['نام برنامه']).sum().reset_index()
    EPG_1399_ekhtesasi_contents=len(EPG_1399_ekhtesasi_contents['نام برنامه'])
    EPG_1399_ekhtesasi_operators_1=EPG_1399_ekhtesasi.copy()
    EPG_1399_ekhtesasi_operators=EPG_1399_ekhtesasi_operators_1.groupby(['نام اپراتور']).sum().reset_index()
    EPG_1399_ekhtesasi_operators=len(EPG_1399_ekhtesasi_operators['نام اپراتور'])
    
    EPG_1399_total=pd.DataFrame()
    EPG_1399_total={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد اپراتور', 'تعداد شبکه', 'کل کاربران ثبت نامی',],
           'sima': ["-", EPG_1399_sima_visit, EPG_1399_sima_duration, "7", "24", EPG_1399_sima_register,],
           'ekhtesasi': [EPG_1399_ekhtesasi_contents, EPG_1399_ekhtesasi_visit, EPG_1399_ekhtesasi_duration, EPG_1399_ekhtesasi_operators, EPG_1399_ekhtesasi_channels, "-",],}
    EPG_1399_total=pd.DataFrame(EPG_1399_total, columns=['parameters', 'sima', 'ekhtesasi'])
    EPG_1399_total=EPG_1399_total.rename(columns={'parameters': 'پارامترها', 'sima': 'شبکه های سیما', 'ekhtesasi': 'شبکه های اختصاصی'})
    ######################################## year of 1400 ####################################
    print("sima 1400")
    EPG_1400_sima_content=EPG_Farvardin_1400_first.loc[0, 'سیما'] + \
                          EPG_Ordibehesht_1400_first.loc[0, 'سیما'] + \
                          EPG_Khordad_1400_first.loc[0, 'سیما'] + \
                          EPG_Tir_1400_first.loc[0, 'سیما'] + \
                          EPG_Mordad_1400_first.loc[0, 'سیما'] + \
                          EPG_Shahrivar_1400_first.loc[0, 'سیما'] + \
                          EPG_Mehr_1400_first.loc[0, 'سیما'] + \
                          EPG_Aban_1400_first.loc[0, 'سیما'] + \
                          EPG_Azar_1400_first.loc[0, 'سیما'] + \
                          EPG_Dey_1400_first.loc[0, 'سیما'] + \
                          EPG_Bahman_1400_first.loc[0, 'سیما']
    EPG_1400_sima_visit=EPG_Farvardin_1400_first.loc[1, 'سیما'] + \
                        EPG_Ordibehesht_1400_first.loc[1, 'سیما'] + \
                        EPG_Khordad_1400_first.loc[1, 'سیما'] + \
                        EPG_Tir_1400_first.loc[1, 'سیما'] + \
                        EPG_Mordad_1400_first.loc[1, 'سیما'] + \
                        EPG_Shahrivar_1400_first.loc[1, 'سیما'] + \
                        EPG_Mehr_1400_first.loc[1, 'سیما'] + \
                        EPG_Aban_1400_first.loc[1, 'سیما'] + \
                        EPG_Azar_1400_first.loc[1, 'سیما'] + \
                        EPG_Dey_1400_first.loc[1, 'سیما'] + \
                        EPG_Bahman_1400_first.loc[1, 'سیما']
    EPG_1400_sima_duration=EPG_Farvardin_1400_first.loc[2, 'سیما'] + \
                           EPG_Ordibehesht_1400_first.loc[2, 'سیما'] + \
                           EPG_Khordad_1400_first.loc[2, 'سیما'] + \
                           EPG_Tir_1400_first.loc[2, 'سیما'] + \
                           EPG_Mordad_1400_first.loc[2, 'سیما'] + \
                           EPG_Shahrivar_1400_first.loc[2, 'سیما'] + \
                           EPG_Mehr_1400_first.loc[2, 'سیما'] + \
                           EPG_Aban_1400_first.loc[2, 'سیما'] + \
                           EPG_Azar_1400_first.loc[2, 'سیما'] + \
                           EPG_Dey_1400_first.loc[2, 'سیما'] + \
                           EPG_Bahman_1400_first.loc[2, 'سیما']
    EPG_1400_sima_channels=EPG_Farvardin_1400_first.loc[3, 'سیما'] + \
                           EPG_Ordibehesht_1400_first.loc[3, 'سیما'] + \
                           EPG_Khordad_1400_first.loc[3, 'سیما'] + \
                           EPG_Tir_1400_first.loc[3, 'سیما'] + \
                           EPG_Mordad_1400_first.loc[3, 'سیما'] + \
                           EPG_Shahrivar_1400_first.loc[3, 'سیما'] + \
                           EPG_Mehr_1400_first.loc[3, 'سیما'] + \
                           EPG_Aban_1400_first.loc[3, 'سیما'] + \
                           EPG_Azar_1400_first.loc[3, 'سیما'] + \
                           EPG_Dey_1400_first.loc[3, 'سیما']+ \
                           EPG_Bahman_1400_first.loc[3, 'سیما']
    EPG_1400_sima_operators=0
    EPG_1400_sima_operators = EPG_Farvardin_1400_first.loc[4, 'سیما'] + \
                              EPG_Ordibehesht_1400_first.loc[4, 'سیما'] + \
                              EPG_Khordad_1400_first.loc[4, 'سیما'] + \
                              EPG_Tir_1400_first.loc[4, 'سیما'] + \
                              EPG_Mordad_1400_first.loc[4, 'سیما'] + \
                              EPG_Shahrivar_1400_first.loc[4, 'سیما'] + \
                              EPG_Mehr_1400_first.loc[4, 'سیما'] + \
                              EPG_Aban_1400_first.loc[4, 'سیما'] + \
                              EPG_Azar_1400_first.loc[4, 'سیما'] + \
                              EPG_Dey_1400_first.loc[4, 'سیما'] + \
                              EPG_Bahman_1400_first.loc[4, 'سیما']
    EPG_1400_RegisterUsers=RegisterActiveUsers_Farvardin_1400['register users'].sum() + \
                           RegisterActiveUsers_Ordibehesht_1400['register users'].sum() + \
                           RegisterActiveUsers_Khordad_1400['register users'].sum() + \
                           RegisterActiveUsers_Tir_1400['register users'].sum() + \
                           RegisterActiveUsers_Mordad_1400['register users'].sum() + \
                           RegisterActiveUsers_Shahrivar_1400['register users'].sum() + \
                           RegisterActiveUsers_Mehr_1400['register users'].sum() + \
                           RegisterActiveUsers_Aban_1400['register users'].sum() + \
                           RegisterActiveUsers_Azar_1400['register users'].sum() + \
                           RegisterActiveUsers_Dey_1400['register users'].sum() + \
                           RegisterActiveUsers_Bahman_1400['register users'].sum()
    
    print("radio 1400")
    EPG_1400_radio_content=EPG_Farvardin_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Ordibehesht_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Khordad_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Tir_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Mordad_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Shahrivar_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Mehr_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Aban_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Azar_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Dey_1400_first.loc[0, 'رادیویی'] + \
                           EPG_Bahman_1400_first.loc[0, 'رادیویی']
    EPG_1400_radio_visit=EPG_Farvardin_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Ordibehesht_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Khordad_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Tir_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Mordad_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Shahrivar_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Mehr_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Aban_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Azar_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Dey_1400_first.loc[1, 'رادیویی'] + \
                         EPG_Bahman_1400_first.loc[1, 'رادیویی']
    EPG_1400_radio_duration=EPG_Farvardin_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Ordibehesht_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Khordad_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Tir_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Mordad_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Shahrivar_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Mehr_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Aban_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Azar_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Dey_1400_first.loc[2, 'رادیویی'] + \
                            EPG_Bahman_1400_first.loc[2, 'رادیویی']
    EPG_1400_radio_channels=EPG_Farvardin_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Ordibehesht_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Khordad_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Tir_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Mordad_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Shahrivar_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Mehr_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Aban_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Azar_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Dey_1400_first.loc[3, 'رادیویی'] + \
                            EPG_Bahman_1400_first.loc[3, 'رادیویی']
    EPG_1400_radio_operators=0
    EPG_1400_radio_operators = EPG_Farvardin_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Ordibehesht_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Khordad_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Tir_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Mordad_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Shahrivar_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Mehr_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Aban_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Azar_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Dey_1400_first.loc[4, 'رادیویی'] + \
                               EPG_Bahman_1400_first.loc[4, 'رادیویی']
    EPG_1400_RegisterUsers=RegisterActiveUsers_Farvardin_1400['register users'].sum() + \
                           RegisterActiveUsers_Ordibehesht_1400['register users'].sum() + \
                           RegisterActiveUsers_Khordad_1400['register users'].sum() + \
                           RegisterActiveUsers_Tir_1400['register users'].sum() + \
                           RegisterActiveUsers_Mordad_1400['register users'].sum() + \
                           RegisterActiveUsers_Shahrivar_1400['register users'].sum() + \
                           RegisterActiveUsers_Mehr_1400['register users'].sum() + \
                           RegisterActiveUsers_Aban_1400['register users'].sum() + \
                           RegisterActiveUsers_Azar_1400['register users'].sum() + \
                           RegisterActiveUsers_Dey_1400['register users'].sum() + \
                           RegisterActiveUsers_Bahman_1400['register users'].sum()
    
    print("ostani 1400")
    EPG_1400_ostani_content=EPG_Farvardin_1400_first.loc[0, 'استانی'] + \
                            EPG_Ordibehesht_1400_first.loc[0, 'استانی'] + \
                            EPG_Khordad_1400_first.loc[0, 'استانی'] + \
                            EPG_Tir_1400_first.loc[0, 'استانی'] + \
                            EPG_Mordad_1400_first.loc[0, 'استانی'] + \
                            EPG_Shahrivar_1400_first.loc[0, 'استانی'] + \
                            EPG_Mehr_1400_first.loc[0, 'استانی'] + \
                            EPG_Aban_1400_first.loc[0, 'استانی'] + \
                            EPG_Azar_1400_first.loc[0, 'استانی'] + \
                            EPG_Dey_1400_first.loc[0, 'استانی'] + \
                            EPG_Bahman_1400_first.loc[0, 'استانی']
    EPG_1400_ostani_visit=EPG_Farvardin_1400_first.loc[1, 'استانی'] + \
                          EPG_Ordibehesht_1400_first.loc[1, 'استانی'] + \
                          EPG_Khordad_1400_first.loc[1, 'استانی'] + \
                          EPG_Tir_1400_first.loc[1, 'استانی'] + \
                          EPG_Mordad_1400_first.loc[1, 'استانی'] + \
                          EPG_Shahrivar_1400_first.loc[1, 'استانی'] + \
                          EPG_Mehr_1400_first.loc[1, 'استانی'] + \
                          EPG_Aban_1400_first.loc[1, 'استانی'] + \
                          EPG_Azar_1400_first.loc[1, 'استانی'] + \
                          EPG_Dey_1400_first.loc[1, 'استانی'] + \
                          EPG_Bahman_1400_first.loc[1, 'استانی']
    EPG_1400_ostani_duration=EPG_Farvardin_1400_first.loc[2, 'استانی'] + \
                             EPG_Ordibehesht_1400_first.loc[2, 'استانی'] + \
                             EPG_Khordad_1400_first.loc[2, 'استانی'] + \
                             EPG_Tir_1400_first.loc[2, 'استانی'] + \
                             EPG_Mordad_1400_first.loc[2, 'استانی'] + \
                             EPG_Shahrivar_1400_first.loc[2, 'استانی'] + \
                             EPG_Mehr_1400_first.loc[2, 'استانی'] + \
                             EPG_Aban_1400_first.loc[2, 'استانی'] + \
                             EPG_Azar_1400_first.loc[2, 'استانی'] + \
                             EPG_Dey_1400_first.loc[2, 'استانی'] + \
                             EPG_Bahman_1400_first.loc[2, 'استانی']
    EPG_1400_ostani_channels=EPG_Farvardin_1400_first.loc[3, 'استانی'] + \
                             EPG_Ordibehesht_1400_first.loc[3, 'استانی'] + \
                             EPG_Khordad_1400_first.loc[3, 'استانی'] + \
                             EPG_Tir_1400_first.loc[3, 'استانی'] + \
                             EPG_Mordad_1400_first.loc[3, 'استانی'] + \
                             EPG_Shahrivar_1400_first.loc[3, 'استانی'] + \
                             EPG_Mehr_1400_first.loc[3, 'استانی'] + \
                             EPG_Aban_1400_first.loc[3, 'استانی'] + \
                             EPG_Azar_1400_first.loc[3, 'استانی'] + \
                             EPG_Dey_1400_first.loc[3, 'استانی'] + \
                             EPG_Bahman_1400_first.loc[3, 'استانی']
    EPG_1400_ostani_operators=0
    EPG_1400_ostani_operators = EPG_Farvardin_1400_first.loc[4, 'استانی'] + \
                                EPG_Ordibehesht_1400_first.loc[4, 'استانی'] + \
                                EPG_Khordad_1400_first.loc[4, 'استانی'] + \
                                EPG_Tir_1400_first.loc[4, 'استانی'] + \
                                EPG_Mordad_1400_first.loc[4, 'استانی'] + \
                                EPG_Shahrivar_1400_first.loc[4, 'استانی'] + \
                                EPG_Mehr_1400_first.loc[4, 'استانی'] + \
                                EPG_Aban_1400_first.loc[4, 'استانی'] + \
                                EPG_Azar_1400_first.loc[4, 'استانی'] + \
                                EPG_Dey_1400_first.loc[4, 'استانی'] + \
                                EPG_Bahman_1400_first.loc[4, 'استانی']
    EPG_1400_RegisterUsers=RegisterActiveUsers_Farvardin_1400['register users'].sum() + \
                           RegisterActiveUsers_Ordibehesht_1400['register users'].sum() + \
                           RegisterActiveUsers_Khordad_1400['register users'].sum() + \
                           RegisterActiveUsers_Tir_1400['register users'].sum() + \
                           RegisterActiveUsers_Mordad_1400['register users'].sum() + \
                           RegisterActiveUsers_Shahrivar_1400['register users'].sum() + \
                           RegisterActiveUsers_Mehr_1400['register users'].sum() + \
                           RegisterActiveUsers_Aban_1400['register users'].sum() + \
                           RegisterActiveUsers_Azar_1400['register users'].sum() + \
                           RegisterActiveUsers_Dey_1400['register users'].sum() + \
                           RegisterActiveUsers_Bahman_1400['register users'].sum()
    
    print("ekhtesasi 1400")
    EPG_1400_ekhtesasi_content=EPG_Farvardin_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Ordibehesht_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Khordad_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Tir_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Mordad_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Shahrivar_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Mehr_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Aban_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Azar_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Dey_1400_first.loc[0, 'اختصاصی'] + \
                               EPG_Bahman_1400_first.loc[0, 'اختصاصی']
    EPG_1400_ekhtesasi_visit=EPG_Farvardin_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Ordibehesht_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Khordad_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Tir_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Mordad_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Shahrivar_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Mehr_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Aban_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Azar_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Dey_1400_first.loc[1, 'اختصاصی'] + \
                             EPG_Bahman_1400_first.loc[1, 'اختصاصی']
    EPG_1400_ekhtesasi_duration=EPG_Farvardin_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Ordibehesht_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Khordad_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Tir_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Mordad_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Shahrivar_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Mehr_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Aban_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Azar_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Dey_1400_first.loc[2, 'اختصاصی'] + \
                                EPG_Bahman_1400_first.loc[2, 'اختصاصی']
    EPG_1400_ekhtesasi_channels=EPG_Farvardin_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Ordibehesht_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Khordad_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Tir_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Mordad_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Shahrivar_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Mehr_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Aban_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Azar_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Dey_1400_first.loc[3, 'اختصاصی'] + \
                                EPG_Bahman_1400_first.loc[3, 'اختصاصی']
    EPG_1400_ekhtesasi_operators=0
    EPG_1400_ekhtesasi_operators = EPG_Farvardin_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Ordibehesht_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Khordad_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Tir_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Mordad_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Shahrivar_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Mehr_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Aban_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Azar_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Dey_1400_first.loc[4, 'اختصاصی'] + \
                                   EPG_Bahman_1400_first.loc[4, 'اختصاصی']
    EPG_1400_RegisterUsers=RegisterActiveUsers_Farvardin_1400['register users'].sum() + \
                           RegisterActiveUsers_Ordibehesht_1400['register users'].sum() + \
                           RegisterActiveUsers_Khordad_1400['register users'].sum() + \
                           RegisterActiveUsers_Tir_1400['register users'].sum() + \
                           RegisterActiveUsers_Mordad_1400['register users'].sum() + \
                           RegisterActiveUsers_Shahrivar_1400['register users'].sum() + \
                           RegisterActiveUsers_Mehr_1400['register users'].sum() + \
                           RegisterActiveUsers_Aban_1400['register users'].sum() + \
                           RegisterActiveUsers_Azar_1400['register users'].sum() + \
                           RegisterActiveUsers_Dey_1400['register users'].sum() + \
                           RegisterActiveUsers_Bahman_1400['register users'].sum()
    
    EPG_1400_total=pd.DataFrame()
    EPG_1400_total={'parameters': ['تعداد محتوا', 'تعداد بازدید', 'زمان بازدید (به دقیقه)', 'تعداد اپراتور', 'تعداد شبکه', 'کل کاربران ثبت نامی',],
           'sima': [EPG_1400_sima_content, EPG_1400_sima_visit, EPG_1400_sima_duration, EPG_1400_sima_operators, EPG_1400_sima_channels, EPG_1400_RegisterUsers,],
           'radio': [EPG_1400_radio_content, EPG_1400_radio_visit, EPG_1400_radio_duration, EPG_1400_radio_operators, EPG_1400_radio_channels, EPG_1400_RegisterUsers,],
           'ostani': [EPG_1400_ostani_content, EPG_1400_ostani_visit, EPG_1400_ostani_duration, EPG_1400_ostani_operators, EPG_1400_ostani_channels, EPG_1400_RegisterUsers,],
           'ekhtesasi': [EPG_1400_ekhtesasi_content, EPG_1400_ekhtesasi_visit, EPG_1400_ekhtesasi_duration, EPG_1400_ekhtesasi_operators, EPG_1400_ekhtesasi_channels, EPG_1400_RegisterUsers,],}
    EPG_1400_total=pd.DataFrame(EPG_1400_total, columns=['parameters', 'sima', 'radio', 'ostani', 'ekhtesasi'])
    EPG_1400_total=EPG_1400_total.rename(columns={'parameters': 'پارامترها', 'sima': 'شبکه های سیما', 'radio': 'شبکه های رادیویی', 'ostani': 'شبکه های استانی', 'ekhtesasi': 'شبکه های اختصاصی'})
    
    
    writer = pd.ExcelWriter('E:/hard/report/total EPG/آمار کلی سال ها.xlsx', engine='xlsxwriter')
    EPG_1397_total.to_excel(writer, 'سال 1397')
    EPG_1398_total.to_excel(writer, 'سال 1398')
    EPG_1399_total.to_excel(writer, 'سال 1399')
    EPG_1400_total.to_excel(writer, 'سال 1400')
    writer.save()
    
    
    return EPG_1397_total, EPG_1398_total, EPG_1399_total, EPG_1400_total
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    