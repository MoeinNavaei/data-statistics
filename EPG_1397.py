
def EPG_1397(EPG_2, EPG_3, EPG_4, EPG_5, EPG_esfand):
    
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
#    EPG_shabake_jamejam_1397=all_data_EPG_1397.query("channels == 'جام جم'")
    
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
#    EPG_shabake_jamejam_1397_visit=EPG_shabake_jamejam_1397['visit'].sum()
    
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
#    EPG_shabake_jamejam_1397_duration=EPG_shabake_jamejam_1397['duration'].sum()
    
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
    
    writer = pd.ExcelWriter('output/EPG 1397/خلاصه آمار سال 1397.xlsx', engine='xlsxwriter')
    all_data_1397_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    all_data_EPG_1397_operators.to_excel(writer, 'آمار اپراتورها')
    all_data_EPG_1397_summary.to_excel(writer, 'خلاصه سال 1397')
    writer.save()
    
    return all_data_EPG_1397_operators, all_data_1397_channels, all_data_EPG_1397_summary




























