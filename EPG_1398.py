    
def EPG_1398(EPG_Farvardin_1398, EPG_Ordibehesht_1398, EPG_Khordad_1398, \
 EPG_Tir_1398, EPG_Mordad_1398, EPG_Shahrivar_1398, \
 EPG_Mehr_1398, EPG_Aban_1398, EPG_Azar_1398, \
 EPG_Dey_1398, EPG_Bahman_1398, EPG_Esfand_1398):
        
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
        
    print("EPG Farvardin 1398")
#    EPG_Farvardin_1398=pd.read_excel('EPG/EPG 1398/EPG Farvardin 1398.xlsx', sheet_name='آمار')
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
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه فروردین 1398.xlsx', engine='xlsxwriter')
    Farvardin_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Farvardin_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Farvardin_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه فروردین')
    writer.save()
    
            ########################### اردیبهشت #############################
    print("EPG Ordibehesht 1398")
#    EPG_Ordibehesht_1398=pd.read_excel('EPG/EPG 1398/EPG Ordibehesht 1398.xlsx', sheet_name='آمار')
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
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه اردیبهشت 1398.xlsx', engine='xlsxwriter')
    Ordibehesht_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Ordibehesht_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Ordibehesht_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه اردیبهشت')
    writer.save()
    
            ########################### خرداد #############################
    print("EPG Khordad 1398")
#    EPG_Khordad_1398=pd.read_excel('EPG/EPG 1398/EPG Khordad 1398.xlsx', sheet_name='آمار')
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
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه خرداد 1398.xlsx', engine='xlsxwriter')
    Khordad_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Khordad_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Khordad_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه خرداد')
    writer.save()
    
            ########################### تیر #############################
    print("EPG Tir 1398")
#    EPG_Tir_1398=pd.read_excel('EPG/EPG 1398/EPG Tir 1398.xlsx', sheet_name='آمار')
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
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه تیر 1398.xlsx', engine='xlsxwriter')
    Tir_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Tir_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Tir_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه تیر')
    writer.save()
    
            ########################### مرداد #############################
    print("EPG Mordad 1398")
#    EPG_Mordad_1398=pd.read_excel('EPG/EPG 1398/EPG Mordad 1398.xlsx', sheet_name='آمار')
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
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه مرداد 1398.xlsx', engine='xlsxwriter')
    Mordad_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Mordad_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Mordad_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه مرداد')
    writer.save()
    
            ########################### شهریور #############################
    print("EPG Shahrivar 1398")
#    EPG_Shahrivar_1398=pd.read_excel('EPG/EPG 1398/EPG Shahrivar 1398.xlsx', sheet_name='آمار')
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
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه شهریور 1398.xlsx', engine='xlsxwriter')
    Shahrivar_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Shahrivar_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Shahrivar_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه شهریور')
    writer.save()
    
            ########################### مهر #############################
    print("EPG Mehr 1398")
#    EPG_Mehr_1398=pd.read_excel('EPG/EPG 1398/EPG Mehr 1398.xlsx', sheet_name='آمار')
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
                     active_user_sepehr_Mehr_1398, active_user_shima_Mehr_1398, active_user_site_Mehr_1398,],}
    
    Mehr_1398_operator_data=pd.DataFrame(Mehr_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])
    
    Mehr_1398_operator_data=Mehr_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})
    
    Mehr_1398_all_data_summary=pd.DataFrame()
    Mehr_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
           'statistics': [all_visit_Mehr_1398, all_duration_Mehr_1398,all_content_sima_Mehr_1398, all_register_user_Mehr_1398,],}
    
    Mehr_1398_all_data_summary=pd.DataFrame(Mehr_1398_all_data_summary, columns=['parameters', 'statistics'])
    
    Mehr_1398_all_data_summary=Mehr_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه مهر 1398.xlsx', engine='xlsxwriter')
    Mehr_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Mehr_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Mehr_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه مهر')
    writer.save()
    
            ########################### آبان #############################
    print("EPG Aban 1398")
#    EPG_Aban_1398=pd.read_excel('EPG/EPG 1398/EPG Aban 1398.xlsx', sheet_name='آمار')
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
                     active_user_sepehr_Aban_1398, active_user_shima_Aban_1398, active_user_site_Aban_1398,],}
    
    Aban_1398_operator_data=pd.DataFrame(Aban_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])
    
    Aban_1398_operator_data=Aban_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})
    
    Aban_1398_all_data_summary=pd.DataFrame()
    Aban_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
           'statistics': [all_visit_Aban_1398, all_duration_Aban_1398,all_content_sima_Aban_1398, all_register_user_Aban_1398,],}
    
    Aban_1398_all_data_summary=pd.DataFrame(Aban_1398_all_data_summary, columns=['parameters', 'statistics'])
    
    Aban_1398_all_data_summary=Aban_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه آبان 1398.xlsx', engine='xlsxwriter')
    Aban_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Aban_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Aban_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه آبان')
    writer.save()
    
            ########################### آذر #############################
    print("EPG Azar 1398")
#    EPG_Azar_1398=pd.read_excel('EPG/EPG 1398/EPG Azar 1398.xlsx', sheet_name='آمار')
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
                     active_user_sepehr_Azar_1398, active_user_shima_Azar_1398, active_user_site_Azar_1398,],}
    
    Azar_1398_operator_data=pd.DataFrame(Azar_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])
    
    Azar_1398_operator_data=Azar_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})
    
    Azar_1398_all_data_summary=pd.DataFrame()
    Azar_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
           'statistics': [all_visit_Azar_1398, all_duration_Azar_1398,all_content_sima_Azar_1398, all_register_user_Azar_1398,],}
    
    Azar_1398_all_data_summary=pd.DataFrame(Azar_1398_all_data_summary, columns=['parameters', 'statistics'])
    
    Azar_1398_all_data_summary=Azar_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه آذر 1398.xlsx', engine='xlsxwriter')
    Azar_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Azar_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Azar_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه آذر')
    writer.save()
    
            ########################### دی #############################
    print("EPG Dey 1398")
#    EPG_Dey_1398=pd.read_excel('EPG/EPG 1398/EPG Dey 1398.xlsx', sheet_name='آمار')
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
                     active_user_sepehr_Dey_1398, active_user_shima_Dey_1398, active_user_site_Dey_1398,],}
    
    Dey_1398_operator_data=pd.DataFrame(Dey_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])
    
    Dey_1398_operator_data=Dey_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})
    
    Dey_1398_all_data_summary=pd.DataFrame()
    Dey_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
           'statistics': [all_visit_Dey_1398, all_duration_Dey_1398,all_content_sima_Dey_1398, all_register_user_Dey_1398,],}
    
    Dey_1398_all_data_summary=pd.DataFrame(Dey_1398_all_data_summary, columns=['parameters', 'statistics'])
    
    Dey_1398_all_data_summary=Dey_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه دی 1398.xlsx', engine='xlsxwriter')
    Dey_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Dey_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Dey_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه دی')
    writer.save()
    
            ########################### بهمن #############################
    print("EPG Bahman 1398")
#    EPG_Bahman_1398=pd.read_excel('EPG/EPG 1398/EPG Bahman 1398.xlsx', sheet_name='آمار')
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
                     active_user_sepehr_Bahman_1398, active_user_shima_Bahman_1398, active_user_site_Bahman_1398,],}
    
    Bahman_1398_operator_data=pd.DataFrame(Bahman_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])
    
    Bahman_1398_operator_data=Bahman_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})
    
    Bahman_1398_all_data_summary=pd.DataFrame()
    Bahman_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
           'statistics': [all_visit_Bahman_1398, all_duration_Bahman_1398,all_content_sima_Bahman_1398, all_register_user_Bahman_1398,],}
    
    Bahman_1398_all_data_summary=pd.DataFrame(Bahman_1398_all_data_summary, columns=['parameters', 'statistics'])
    
    Bahman_1398_all_data_summary=Bahman_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه بهمن 1398.xlsx', engine='xlsxwriter')
    Bahman_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Bahman_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Bahman_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه بهمن')
    writer.save()
    
            ########################### اسفند #############################
    print("EPG Esfand 1398")
#    EPG_Esfand_1398=pd.read_excel('EPG/EPG 1398/EPG Esfand 1398.xlsx', sheet_name='آمار')
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
                     active_user_sepehr_Esfand_1398, active_user_shima_Esfand_1398, active_user_site_Esfand_1398,],}
    
    Esfand_1398_operator_data=pd.DataFrame(Esfand_1398_operator_data, columns=['operators', 'visit', 'register', 'active'])
    
    Esfand_1398_operator_data=Esfand_1398_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال'})
    
    Esfand_1398_all_data_summary=pd.DataFrame()
    Esfand_1398_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی',],
           'statistics': [all_visit_Esfand_1398, all_duration_Esfand_1398,all_content_sima_Esfand_1398, all_register_user_Esfand_1398,],}
    
    Esfand_1398_all_data_summary=pd.DataFrame(Esfand_1398_all_data_summary, columns=['parameters', 'statistics'])
    
    Esfand_1398_all_data_summary=Esfand_1398_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    writer = pd.ExcelWriter('output/EPG 1398/ماه اسفند 1398.xlsx', engine='xlsxwriter')
    Esfand_1398_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    Esfand_1398_operator_data.to_excel(writer, 'آمار اپراتورها')
    Esfand_1398_all_data_summary.to_excel(writer, 'خلاصه آمار ماه اسفند')
    writer.save()
    
    
    return Farvardin_1398_sima_visit_channels, Farvardin_1398_operator_data, Farvardin_1398_all_data_summary, \
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
    Esfand_1398_sima_visit_channels, Esfand_1398_operator_data, Esfand_1398_all_data_summary
    
        
        
        
        
        
        
        
        
        
        
        
        
        
