    
def EPG_1400(all_data, sima, RegisterActiveUsers):
            
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
        
            ########################### current month #############################
    print("EPG Farvardin 1400")
#    EPG_current_month=pd.read_excel('EPG/EPG 1400/EPG Farvardin 1400.xlsx', sheet_name='آمار')
#    EPG_current_month.fillna(0, inplace=True)
    
    sima_all=sima.copy()
    
    shabake_1=sima_all.query("channel == 'شبکه 1'")
    sima_1_visit_current_month=shabake_1['تعداد بازدید'].sum()
    sima_1_duration_current_month=shabake_1['مدت بازدید'].sum()
    sima_1_duration_current_month=round(sima_1_duration_current_month, 0)
    sima_1_content=shabake_1.copy()
    sima_1_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_1_content=len(sima_1_content)
    
    shabake_2=sima_all.query("channel == 'شبکه 2'")
    sima_2_visit_current_month=shabake_2['تعداد بازدید'].sum()
    sima_2_duration_current_month=shabake_2['مدت بازدید'].sum()
    sima_2_duration_current_month=round(sima_2_duration_current_month, 0)
    sima_2_content=shabake_2.copy()
    sima_2_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_2_content=len(sima_2_content)
    
    shabake_3=sima_all.query("channel == 'شبکه 3'")
    sima_3_visit_current_month=shabake_3['تعداد بازدید'].sum()
    sima_3_duration_current_month=shabake_3['مدت بازدید'].sum()
    sima_3_duration_current_month=round(sima_3_duration_current_month, 0)
    sima_3_content=shabake_3.copy()
    sima_3_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_3_content=len(sima_3_content)
    
    shabake_4=sima_all.query("channel == 'شبکه 4'")
    sima_4_visit_current_month=shabake_4['تعداد بازدید'].sum()
    sima_4_duration_current_month=shabake_4['مدت بازدید'].sum()
    sima_4_duration_current_month=round(sima_4_duration_current_month, 0)
    sima_4_content=shabake_4.copy()
    sima_4_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_4_content=len(sima_4_content)
    
    shabake_5=sima_all.query("channel == 'شبکه 5'")
    sima_5_visit_current_month=shabake_5['تعداد بازدید'].sum()
    sima_5_duration_current_month=shabake_5['مدت بازدید'].sum()
    sima_5_duration_current_month=round(sima_5_duration_current_month, 0)
    sima_5_content=shabake_5.copy()
    sima_5_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_5_content=len(sima_5_content)
    
    shabake_khabar=sima_all.query("channel == 'خبر'")
    sima_khabar_visit_current_month=shabake_khabar['تعداد بازدید'].sum()
    sima_khabar_duration_current_month=shabake_khabar['مدت بازدید'].sum()
    sima_khabar_duration_current_month=round(sima_khabar_duration_current_month, 0)
    sima_khabar_content=shabake_khabar.copy()
    sima_khabar_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_khabar_content=len(sima_khabar_content)
    
    shabake_Ofogh=sima_all.query("channel == 'افق'")
    sima_Ofogh_visit_current_month=shabake_Ofogh['تعداد بازدید'].sum()
    sima_Ofogh_duration_current_month=shabake_Ofogh['مدت بازدید'].sum()
    sima_Ofogh_duration_current_month=round(sima_Ofogh_duration_current_month, 0)
    sima_Ofogh_content=shabake_Ofogh.copy()
    sima_Ofogh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_Ofogh_content=len(sima_Ofogh_content)
    
    shabake_pooya=sima_all.query("channel == 'پویا'")
    sima_pooya_visit_current_month=shabake_pooya['تعداد بازدید'].sum()
    sima_pooya_duration_current_month=shabake_pooya['مدت بازدید'].sum()
    sima_pooya_duration_current_month=round(sima_pooya_duration_current_month, 0)
    sima_pooya_content=shabake_pooya.copy()
    sima_pooya_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_pooya_content=len(sima_pooya_content)
    
    shabake_omid=sima_all.query("channel == 'امید'")
    sima_omid_visit_current_month=shabake_omid['تعداد بازدید'].sum()
    sima_omid_duration_current_month=shabake_omid['مدت بازدید'].sum()
    sima_omid_duration_current_month=round(sima_omid_duration_current_month, 0)
    sima_omid_content=shabake_omid.copy()
    sima_omid_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_omid_content=len(sima_omid_content)
    
    shabake_ifilm=sima_all.query("channel == 'آی فیلم'")
    sima_ifilm_visit_current_month=shabake_ifilm['تعداد بازدید'].sum()
    sima_ifilm_duration_current_month=shabake_ifilm['مدت بازدید'].sum()
    sima_ifilm_duration_current_month=round(sima_ifilm_duration_current_month, 0)
    sima_ifilm_content=shabake_ifilm.copy()
    sima_ifilm_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_ifilm_content=len(sima_ifilm_content)
    
    shabake_namayesh=sima_all.query("channel == 'نمایش'")
    sima_namayesh_visit_current_month=shabake_namayesh['تعداد بازدید'].sum()
    sima_namayesh_duration_current_month=shabake_namayesh['مدت بازدید'].sum()
    sima_namayesh_duration_current_month=round(sima_namayesh_duration_current_month, 0)
    sima_namayesh_content=shabake_namayesh.copy()
    sima_namayesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_namayesh_content=len(sima_namayesh_content)
    
    shabake_tamasha=sima_all.query("channel == 'تماشا'")
    sima_tamasha_visit_current_month=shabake_tamasha['تعداد بازدید'].sum()
    sima_tamasha_duration_current_month=shabake_tamasha['مدت بازدید'].sum()
    sima_tamasha_duration_current_month=round(sima_tamasha_duration_current_month, 0)
    sima_tamasha_content=shabake_tamasha.copy()
    sima_tamasha_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_tamasha_content=len(sima_tamasha_content)
    
    shabake_mostanad=sima_all.query("channel == 'مستند'")
    sima_mostanad_visit_current_month=shabake_mostanad['تعداد بازدید'].sum()
    sima_mostanad_duration_current_month=shabake_mostanad['مدت بازدید'].sum()
    sima_mostanad_duration_current_month=round(sima_mostanad_duration_current_month, 0)
    sima_mostanad_content=shabake_mostanad.copy()
    sima_mostanad_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_mostanad_content=len(sima_mostanad_content)
    
    shabake_shoma=sima_all.query("channel == 'شما'")
    sima_shoma_visit_current_month=shabake_mostanad['تعداد بازدید'].sum()
    sima_shoma_duration_current_month=shabake_mostanad['مدت بازدید'].sum()
    sima_shoma_duration_current_month=round(sima_shoma_duration_current_month, 0)
    sima_shoma_content=shabake_shoma.copy()
    sima_shoma_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_shoma_content=len(sima_shoma_content)
    
    shabake_amozesh=sima_all.query("channel == 'آموزش'")
    sima_amozesh_visit_current_month=shabake_amozesh['تعداد بازدید'].sum()
    sima_amozesh_duration_current_month=shabake_amozesh['مدت بازدید'].sum()
    sima_amozesh_duration_current_month=round(sima_amozesh_duration_current_month, 0)
    sima_amozesh_content=shabake_amozesh.copy()
    sima_amozesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_amozesh_content=len(sima_amozesh_content)
    
    shabake_varzesh=sima_all.query("channel == 'ورزش'")
    sima_varzesh_visit_current_month=shabake_varzesh['تعداد بازدید'].sum()
    sima_varzesh_duration_current_month=shabake_varzesh['مدت بازدید'].sum()
    sima_varzesh_duration_current_month=round(sima_varzesh_duration_current_month, 0)
    sima_varzesh_content=shabake_varzesh.copy()
    sima_varzesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_varzesh_content=len(sima_varzesh_content)
    
    shabake_nasim=sima_all.query("channel == 'نسیم'")
    sima_nasim_visit_current_month=shabake_nasim['تعداد بازدید'].sum()
    sima_nasim_duration_current_month=shabake_nasim['مدت بازدید'].sum()
    sima_nasim_duration_current_month=round(sima_nasim_duration_current_month, 0)
    sima_nasim_content=shabake_nasim.copy()
    sima_nasim_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_nasim_content=len(sima_nasim_content)
    
    shabake_qoran=sima_all.query("channel == 'قرآن'")
    sima_qoran_visit_current_month=shabake_qoran['تعداد بازدید'].sum()
    sima_qoran_duration_current_month=shabake_qoran['مدت بازدید'].sum()
    sima_qoran_duration_current_month=round(sima_qoran_duration_current_month, 0)
    sima_qoran_content=shabake_qoran.copy()
    sima_qoran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_qoran_content=len(sima_qoran_content)
    
    shabake_salamat=sima_all.query("channel == 'سلامت'")
    sima_salamat_visit_current_month=shabake_salamat['تعداد بازدید'].sum()
    sima_salamat_duration_current_month=shabake_salamat['مدت بازدید'].sum()
    sima_salamat_duration_current_month=round(sima_salamat_duration_current_month, 0)
    sima_salamat_content=shabake_salamat.copy()
    sima_salamat_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_salamat_content=len(sima_salamat_content)
    
    shabake_irankala=sima_all.query("channel == 'ایران کالا'")
    sima_irankala_visit_current_month=shabake_irankala['تعداد بازدید'].sum()
    sima_irankala_duration_current_month=shabake_irankala['مدت بازدید'].sum()
    sima_irankala_duration_current_month=round(sima_irankala_duration_current_month, 0)
    sima_irankala_content=shabake_irankala.copy()
    sima_irankala_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_irankala_content=len(sima_irankala_content)
    
    shabake_alalam=sima_all.query("channel == 'العالم'")
    sima_alalam_visit_current_month=shabake_alalam['تعداد بازدید'].sum()
    sima_alalam_duration_current_month=shabake_alalam['مدت بازدید'].sum()
    sima_alalam_duration_current_month=round(sima_alalam_duration_current_month, 0)
    sima_alalam_content=shabake_alalam.copy()
    sima_alalam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_alalam_content=len(sima_alalam_content)
    
    shabake_alkosar=sima_all.query("channel == 'الکوثر'")
    sima_alkosar_visit_current_month=shabake_alkosar['تعداد بازدید'].sum()
    sima_alkosar_duration_current_month=shabake_alkosar['مدت بازدید'].sum()
    sima_alkosar_duration_current_month=round(sima_alkosar_duration_current_month, 0)
    sima_alkosar_content=shabake_alkosar.copy()
    sima_alkosar_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_alkosar_content=len(sima_alkosar_content)
    
    shabake_presstv=sima_all.query("channel == 'پرس تی وی'")
    sima_presstv_visit_current_month=shabake_presstv['تعداد بازدید'].sum()
    sima_presstv_duration_current_month=shabake_presstv['مدت بازدید'].sum()
    sima_presstv_duration_current_month=round(sima_presstv_duration_current_month, 0)
    sima_presstv_content=shabake_presstv.copy()
    sima_presstv_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_presstv_content=len(sima_presstv_content)
    
    shabake_sepehr=sima_all.query("channel == 'سپهر'")
    sima_sepehr_visit_current_month=shabake_sepehr['تعداد بازدید'].sum()
    sima_sepehr_duration_current_month=shabake_sepehr['مدت بازدید'].sum()
    sima_sepehr_duration_current_month=round(sima_sepehr_duration_current_month, 0)
    sima_sepehr_content=shabake_sepehr.copy()
    sima_sepehr_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_sepehr_content=len(sima_sepehr_content)
    
#    shabake_jamejam=sima_all.query("channel == 'جام جم'")
#    sima_jamejam_visit_Khordad_1400=shabake_jamejam['تعداد بازدید'].sum()
#    sima_jamejam_duration_Khordad_1400=shabake_jamejam['مدت بازدید'].sum()
#    sima_jamejam_duration_Khordad_1400=round(sima_jamejam_duration_Khordad_1400, 0)
#    sima_jamejam_content=shabake_jamejam.copy()
#    sima_jamejam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
#    sima_jamejam_content=len(sima_jamejam_content)
    
    
    lenz_all=all_data.query("operator == 'لنز'")
    lenz_visit_current_month=lenz_all['تعداد بازدید'].sum()
    lenz_duration_current_month=lenz_all['مدت بازدید'].sum()
    lenz_duration_current_month=round(lenz_duration_current_month, 0)
    lenz_sima=lenz_all.query("type == 'سراسری'")
    lenz_radio=lenz_all.query("type == 'رادیو'")
    lenz_ostani=lenz_all.query("type == 'استانی'")
    lenz_ekhtesasi=lenz_all.query("type == 'اختصاصی'")
    lenz_sima_visit=lenz_sima['تعداد بازدید'].sum()
    lenz_sima_duration=lenz_sima['مدت بازدید'].sum()
    lenz_sima_content=lenz_sima.copy()
    lenz_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    lenz_sima_content=len(lenz_sima_content)
    lenz_radio_visit=lenz_radio['تعداد بازدید'].sum()
    lenz_radio_duration=lenz_radio['مدت بازدید'].sum()
    lenz_radio_content=lenz_radio.copy()
    lenz_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    lenz_radio_content=len(lenz_radio_content)
    lenz_ostani_visit=lenz_ostani['تعداد بازدید'].sum()
    lenz_ostani_duration=lenz_ostani['مدت بازدید'].sum()
    lenz_ostani_content=lenz_ostani.copy()
    lenz_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    lenz_ostani_content=len(lenz_ostani_content)
    lenz_ekhtesasi_visit=lenz_ekhtesasi['تعداد بازدید'].sum()
    lenz_ekhtesasi_duration=lenz_ekhtesasi['مدت بازدید'].sum()
    lenz_ekhtesasi_content=lenz_ekhtesasi.copy()
    lenz_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    lenz_ekhtesasi_content=len(lenz_ekhtesasi_content)
    
    aio_all=all_data.query("operator == 'آیو'")
    aio_visit_current_month=aio_all['تعداد بازدید'].sum()
    aio_duration_current_month=aio_all['مدت بازدید'].sum()
    aio_duration_current_month=round(aio_duration_current_month, 0)
    aio_sima=aio_all.query("type == 'سراسری'")
    aio_radio=aio_all.query("type == 'رادیو'")
    aio_ostani=aio_all.query("type == 'استانی'")
    aio_ekhtesasi=aio_all.query("type == 'اختصاصی'")
    aio_sima_visit=aio_sima['تعداد بازدید'].sum()
    aio_sima_duration=aio_sima['مدت بازدید'].sum()
    aio_sima_content=aio_sima.copy()
    aio_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    aio_sima_content=len(aio_sima_content)
    aio_radio_visit=aio_radio['تعداد بازدید'].sum()
    aio_radio_duration=aio_radio['مدت بازدید'].sum()
    aio_radio_content=aio_radio.copy()
    aio_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    aio_radio_content=len(aio_radio_content)
    aio_ostani_visit=aio_ostani['تعداد بازدید'].sum()
    aio_ostani_duration=aio_ostani['مدت بازدید'].sum()
    aio_ostani_content=aio_ostani.copy()
    aio_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    aio_ostani_content=len(aio_ostani_content)
    aio_ekhtesasi_visit=aio_ekhtesasi['تعداد بازدید'].sum()
    aio_ekhtesasi_duration=aio_ekhtesasi['مدت بازدید'].sum()
    aio_ekhtesasi_content=aio_ekhtesasi.copy()
    aio_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    aio_ekhtesasi_content=len(aio_ekhtesasi_content)
    
    anten_all=all_data.query("operator == 'آنتن'")
    anten_visit_current_month=anten_all['تعداد بازدید'].sum()
    anten_duration_current_month=anten_all['مدت بازدید'].sum()
    anten_duration_current_month=round(anten_duration_current_month, 0)
    anten_sima=anten_all.query("type == 'سراسری'")
    anten_radio=anten_all.query("type == 'رادیو'")
    anten_ostani=anten_all.query("type == 'استانی'")
    anten_ekhtesasi=anten_all.query("type == 'اختصاصی'")
    anten_sima_visit=anten_sima['تعداد بازدید'].sum()
    anten_sima_duration=anten_sima['مدت بازدید'].sum()
    anten_sima_content=anten_sima.copy()
    anten_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    anten_sima_content=len(anten_sima_content)
    anten_radio_visit=anten_radio['تعداد بازدید'].sum()
    anten_radio_duration=anten_radio['مدت بازدید'].sum()
    anten_radio_content=anten_radio.copy()
    anten_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    anten_radio_content=len(anten_radio_content)
    anten_ostani_visit=anten_ostani['تعداد بازدید'].sum()
    anten_ostani_duration=anten_ostani['مدت بازدید'].sum()
    anten_ostani_content=anten_ostani.copy()
    anten_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    anten_ostani_content=len(anten_ostani_content)
    anten_ekhtesasi_visit=anten_ekhtesasi['تعداد بازدید'].sum()
    anten_ekhtesasi_duration=anten_ekhtesasi['مدت بازدید'].sum()
    anten_ekhtesasi_content=anten_ekhtesasi.copy()
    anten_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    anten_ekhtesasi_content=len(anten_ekhtesasi_content)
    
    tva_all=all_data.query("operator == 'تیوا'")
    tva_visit_current_month=tva_all['تعداد بازدید'].sum()
    tva_duration_current_month=tva_all['مدت بازدید'].sum()
    tva_duration_current_month=round(tva_duration_current_month, 0)
    tva_sima=tva_all.query("type == 'سراسری'")
    tva_radio=tva_all.query("type == 'رادیو'")
    tva_ostani=tva_all.query("type == 'استانی'")
    tva_ekhtesasi=tva_all.query("type == 'اختصاصی'")
    tva_sima_visit=tva_sima['تعداد بازدید'].sum()
    tva_sima_duration=tva_sima['مدت بازدید'].sum()
    tva_sima_content=tva_sima.copy()
    tva_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    tva_sima_content=len(tva_sima_content)
    tva_radio_visit=tva_radio['تعداد بازدید'].sum()
    tva_radio_duration=tva_radio['مدت بازدید'].sum()
    tva_radio_content=tva_radio.copy()
    tva_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    tva_radio_content=len(tva_radio_content)
    tva_ostani_visit=tva_ostani['تعداد بازدید'].sum()
    tva_ostani_duration=tva_ostani['مدت بازدید'].sum()
    tva_ostani_content=tva_ostani.copy()
    tva_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    tva_ostani_content=len(tva_ostani_content)
    tva_ekhtesasi_visit=tva_ekhtesasi['تعداد بازدید'].sum()
    tva_ekhtesasi_duration=tva_ekhtesasi['مدت بازدید'].sum()
    tva_ekhtesasi_content=tva_ekhtesasi.copy()
    tva_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    tva_ekhtesasi_content=len(tva_ekhtesasi_content)
    
    fam_all=all_data.query("operator == 'فام'")
    fam_visit_current_month=fam_all['تعداد بازدید'].sum()
    fam_duration_current_month=fam_all['مدت بازدید'].sum()
    fam_duration_current_month=round(fam_duration_current_month, 0)
    fam_sima=fam_all.query("type == 'سراسری'")
    fam_radio=fam_all.query("type == 'رادیو'")
    fam_ostani=fam_all.query("type == 'استانی'")
    fam_ekhtesasi=fam_all.query("type == 'اختصاصی'")
    fam_sima_visit=fam_sima['تعداد بازدید'].sum()
    fam_sima_duration=fam_sima['مدت بازدید'].sum()
    fam_sima_content=fam_sima.copy()
    fam_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    fam_sima_content=len(fam_sima_content)
    fam_radio_visit=fam_radio['تعداد بازدید'].sum()
    fam_radio_duration=fam_radio['مدت بازدید'].sum()
    fam_radio_content=fam_radio.copy()
    fam_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    fam_radio_content=len(fam_radio_content)
    fam_ostani_visit=fam_radio['تعداد بازدید'].sum()
    fam_ostani_duration=fam_radio['مدت بازدید'].sum()
    fam_ostani_content=fam_ostani.copy()
    fam_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    fam_ostani_content=len(fam_ostani_content)
    fam_ekhtesasi_visit=fam_ekhtesasi['تعداد بازدید'].sum()
    fam_ekhtesasi_duration=fam_ekhtesasi['مدت بازدید'].sum()
    fam_ekhtesasi_content=fam_ekhtesasi.copy()
    fam_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    fam_ekhtesasi_content=len(fam_ekhtesasi_content)
    
    telewebion_all=all_data.query("operator == 'تلوبیون'")
    telewebion_visit_current_month=telewebion_all['تعداد بازدید'].sum()
    telewebion_duration_current_month=telewebion_all['مدت بازدید'].sum()
    telewebion_duration_current_month=round(telewebion_duration_current_month, 0)
    telewebion_sima=telewebion_all.query("type == 'سراسری'")
    telewebion_radio=telewebion_all.query("type == 'رادیو'")
    telewebion_ostani=telewebion_all.query("type == 'استانی'")
    telewebion_ekhtesasi=telewebion_all.query("type == 'اختصاصی'")
    telewebion_sima_visit=telewebion_sima['تعداد بازدید'].sum()
    telewebion_sima_duration=telewebion_sima['مدت بازدید'].sum()
    telewebion_sima_content=telewebion_sima.copy()
    telewebion_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    telewebion_sima_content=len(telewebion_sima_content)
    telewebion_radio_visit=telewebion_radio['تعداد بازدید'].sum()
    telewebion_radio_duration=telewebion_radio['مدت بازدید'].sum()
    telewebion_radio_content=telewebion_radio.copy()
    telewebion_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    telewebion_radio_content=len(telewebion_radio_content)
    telewebion_ostani_visit=telewebion_ostani['تعداد بازدید'].sum()
    telewebion_ostani_duration=telewebion_ostani['مدت بازدید'].sum()
    telewebion_ostani_content=telewebion_ostani.copy()
    telewebion_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    telewebion_ostani_content=len(telewebion_ostani_content)
    telewebion_ekhtesasi_visit=telewebion_ekhtesasi['تعداد بازدید'].sum()
    telewebion_ekhtesasi_duration=telewebion_ekhtesasi['مدت بازدید'].sum()
    telewebion_ekhtesasi_content=telewebion_ekhtesasi.copy()
    telewebion_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    telewebion_ekhtesasi_content=len(telewebion_ekhtesasi_content)
    
    sepehr_all=all_data.query("operator == 'سپهر'")
    sepehr_visit_current_month=sepehr_all['تعداد بازدید'].sum()
    sepehr_duration_current_month=sepehr_all['مدت بازدید'].sum()
    sepehr_duration_current_month=round(sepehr_duration_current_month, 0)
    sepehr_sima=sepehr_all.query("type == 'سراسری'")
    sepehr_radio=sepehr_all.query("type == 'رادیو'")
    sepehr_ostani=sepehr_all.query("type == 'استانی'")
    sepehr_ekhtesasi=sepehr_all.query("type == 'اختصاصی'")
    sepehr_sima_visit=sepehr_sima['تعداد بازدید'].sum()
    sepehr_sima_duration=sepehr_sima['مدت بازدید'].sum()
    sepehr_sima_content=sepehr_sima.copy()
    sepehr_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sepehr_sima_content=len(sepehr_sima_content)
    sepehr_radio_visit=sepehr_radio['تعداد بازدید'].sum()
    sepehr_radio_duration=sepehr_radio['مدت بازدید'].sum()
    sepehr_radio_content=sepehr_radio.copy()
    sepehr_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sepehr_radio_content=len(sepehr_radio_content)
    sepehr_ostani_visit=sepehr_ostani['تعداد بازدید'].sum()
    sepehr_ostani_duration=sepehr_ostani['مدت بازدید'].sum()
    sepehr_ostani_content=sepehr_ostani.copy()
    sepehr_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sepehr_ostani_content=len(sepehr_ostani_content)
    sepehr_ekhtesasi_visit=sepehr_ekhtesasi['تعداد بازدید'].sum()
    sepehr_ekhtesasi_duration=sepehr_ekhtesasi['مدت بازدید'].sum()
    sepehr_ekhtesasi_content=sepehr_ekhtesasi.copy()
    sepehr_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sepehr_ekhtesasi_content=len(sepehr_ekhtesasi_content)
    
    shima_all=all_data.query("operator == 'شیما'")
    shima_visit_current_month=shima_all['تعداد بازدید'].sum()
    shima_duration_current_month=shima_all['مدت بازدید'].sum()
    shima_duration_current_month=round(shima_duration_current_month, 0)
    shima_sima=shima_all.query("type == 'سراسری'")
    shima_radio=shima_all.query("type == 'رادیو'")
    shima_ostani=shima_all.query("type == 'استانی'")
    shima_ekhtesasi=shima_all.query("type == 'اختصاصی'")
    shima_sima_visit=shima_sima['تعداد بازدید'].sum()
    shima_sima_duration=shima_sima['مدت بازدید'].sum()
    shima_sima_content=shima_sima.copy()
    shima_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    shima_sima_content=len(shima_sima_content)
    shima_radio_visit=shima_radio['تعداد بازدید'].sum()
    shima_radio_duration=shima_radio['مدت بازدید'].sum()
    shima_radio_content=shima_radio.copy()
    shima_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    shima_radio_content=len(shima_radio_content)
    shima_ostani_visit=shima_ostani['تعداد بازدید'].sum()
    shima_ostani_duration=shima_ostani['مدت بازدید'].sum()
    shima_ostani_content=shima_ostani.copy()
    shima_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    shima_ostani_content=len(shima_ostani_content)
    shima_ekhtesasi_visit=shima_ekhtesasi['تعداد بازدید'].sum()
    shima_ekhtesasi_duration=shima_ekhtesasi['مدت بازدید'].sum()
    shima_ekhtesasi_content=shima_ekhtesasi.copy()
    shima_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    shima_ekhtesasi_content=len(shima_ekhtesasi_content)
    
    site_all=all_data.query("operator == 'سایت شبکه ها'")
    site_visit_current_month=site_all['تعداد بازدید'].sum()
    site_duration_current_month=site_all['مدت بازدید'].sum()
    site_duration_current_month=round(site_duration_current_month, 0)
    site_sima=site_all.query("type == 'سراسری'")
    site_radio=site_all.query("type == 'رادیو'")
    site_ostani=site_all.query("type == 'استانی'")
    site_ekhtesasi=site_all.query("type == 'اختصاصی'")
    site_sima_visit=site_sima['تعداد بازدید'].sum()
    site_sima_duration=site_sima['مدت بازدید'].sum()
    site_sima_content=site_sima.copy()
    site_sima_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    site_sima_content=len(site_sima_content)
    site_radio_visit=site_radio['تعداد بازدید'].sum()
    site_radio_duration=site_radio['مدت بازدید'].sum()
    site_radio_content=site_radio.copy()
    site_radio_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    site_radio_content=len(site_radio_content)
    site_ostani_visit=site_ostani['تعداد بازدید'].sum()
    site_ostani_duration=site_ostani['مدت بازدید'].sum()
    site_ostani_content=site_ostani.copy()
    site_ostani_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    site_ostani_content=len(site_ostani_content)
    site_ekhtesasi_visit=site_ekhtesasi['تعداد بازدید'].sum()
    site_ekhtesasi_duration=site_ekhtesasi['مدت بازدید'].sum()
    site_ekhtesasi_content=site_ekhtesasi.copy()
    site_ekhtesasi_content.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    site_ekhtesasi_content=len(site_ekhtesasi_content)

    
    register_user_lenz_current_month=RegisterActiveUsers.iat[0, 1]
    register_user_aio_current_month=RegisterActiveUsers.iat[1, 1]
    register_user_anten_current_month=RegisterActiveUsers.iat[2, 1]
    register_user_tva_current_month=RegisterActiveUsers.iat[3, 1]
    register_user_fam_current_month=RegisterActiveUsers.iat[4, 1]
    register_user_telewebion_current_month=RegisterActiveUsers.iat[5, 1]
    register_user_sepehr_current_month=RegisterActiveUsers.iat[6, 1]
    register_user_shima_current_month=RegisterActiveUsers.iat[7, 1]
    register_user_site_current_month=RegisterActiveUsers.iat[8, 1]
    
    active_user_lenz_current_month=RegisterActiveUsers.iat[0, 2]
    active_user_aio_current_month=RegisterActiveUsers.iat[1, 2]
    active_user_anten_current_month=RegisterActiveUsers.iat[2, 2]
    active_user_tva_current_month=RegisterActiveUsers.iat[3, 2]
    active_user_fam_current_month=RegisterActiveUsers.iat[4, 2]
    active_user_telewebion_current_month=RegisterActiveUsers.iat[5, 2]
    active_user_sepehr_current_month=RegisterActiveUsers.iat[6, 2]
    active_user_shima_current_month=RegisterActiveUsers.iat[7, 2]
    active_user_site_current_month=RegisterActiveUsers.iat[8, 2]
    
    
    all_visit_current_month=all_data['تعداد بازدید'].sum()
    all_duration_current_month=all_data['مدت بازدید'].sum()
    all_duration_current_month=round(all_duration_current_month, 0)
    all_data_content=all_data.copy()
    all_data_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    all_content_sima_current_month=len(all_data_content)
    all_register_user_current_month=RegisterActiveUsers['register users'].sum()
    all_active_user_current_month=RegisterActiveUsers['active users'].sum()
    
    
    current_month_sima_visit_channels=pd.DataFrame()
    current_month_sima_visit_channels={'channels': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                         'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                         'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                         'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                         'شبکه العالم', 'شبکه الکوثر', 'شبکه پرس تی وی', 'شبکه سپهر',],
           'content': [sima_1_content, sima_2_content, sima_3_content,
                     sima_4_content, sima_5_content, sima_khabar_content,
                     sima_Ofogh_content, sima_pooya_content, sima_omid_content,
                     sima_ifilm_content, sima_namayesh_content, sima_tamasha_content,
                     sima_mostanad_content, sima_shoma_content, sima_amozesh_content,
                     sima_varzesh_content, sima_nasim_content, sima_qoran_content,
                     sima_salamat_content, sima_irankala_content, sima_alalam_content,
                     sima_alkosar_content, sima_presstv_content, sima_sepehr_content,],
            'visit': [sima_1_visit_current_month, sima_2_visit_current_month, sima_3_visit_current_month,
                     sima_4_visit_current_month, sima_5_visit_current_month, sima_khabar_visit_current_month,
                     sima_Ofogh_visit_current_month, sima_pooya_visit_current_month, sima_omid_visit_current_month,
                     sima_ifilm_visit_current_month, sima_namayesh_visit_current_month, sima_tamasha_visit_current_month,
                     sima_mostanad_visit_current_month, sima_shoma_visit_current_month, sima_amozesh_visit_current_month,
                     sima_varzesh_visit_current_month, sima_nasim_visit_current_month, sima_qoran_visit_current_month,
                     sima_salamat_visit_current_month, sima_irankala_visit_current_month, sima_alalam_visit_current_month,
                     sima_alkosar_visit_current_month, sima_presstv_visit_current_month, sima_sepehr_visit_current_month,],
            'duration': [sima_1_duration_current_month, sima_2_duration_current_month, sima_3_duration_current_month,
                     sima_4_duration_current_month, sima_5_duration_current_month, sima_khabar_duration_current_month,
                     sima_Ofogh_duration_current_month, sima_pooya_duration_current_month, sima_omid_duration_current_month,
                     sima_ifilm_duration_current_month, sima_namayesh_duration_current_month, sima_tamasha_duration_current_month,
                     sima_mostanad_duration_current_month, sima_shoma_duration_current_month, sima_amozesh_duration_current_month,
                     sima_varzesh_duration_current_month, sima_nasim_duration_current_month, sima_qoran_duration_current_month,
                     sima_salamat_duration_current_month, sima_irankala_duration_current_month, sima_alalam_duration_current_month,
                     sima_alkosar_duration_current_month, sima_presstv_duration_current_month, sima_sepehr_duration_current_month,],}
    current_month_sima_visit_channels=pd.DataFrame(current_month_sima_visit_channels, columns=['channels', 'content', 'visit', 'duration'])
    
    current_month_sima_visit_channels=current_month_sima_visit_channels.rename(columns={'channels': 'نام شبکه', 'content': 'تعداد محتوا', 'visit': 'تعداد بازدید', 'duration': 'مدت زمان بازدید (به دقیقه)'})
    
    current_month_operator_data=pd.DataFrame()
    current_month_operator_data={'operators': ['لنز', 'آیو', 'آنتن', 'تیوا', 'فام', 'تلوبیون', 'سپهر', 'شیما', 'سایت شبکه ها',],
           'visit': [lenz_visit_current_month, aio_visit_current_month, anten_visit_current_month,
                     tva_visit_current_month, fam_visit_current_month, telewebion_visit_current_month,
                     sepehr_visit_current_month, shima_visit_current_month, site_visit_current_month,],
           'register': [register_user_lenz_current_month, register_user_aio_current_month, register_user_anten_current_month,
                     register_user_tva_current_month, register_user_fam_current_month, register_user_telewebion_current_month,
                     register_user_sepehr_current_month, register_user_shima_current_month, register_user_site_current_month,],
           'active': [active_user_lenz_current_month, active_user_aio_current_month, active_user_anten_current_month,
                     active_user_tva_current_month, active_user_fam_current_month, active_user_telewebion_current_month,
                     active_user_sepehr_current_month, active_user_shima_current_month, active_user_site_current_month,],
            'sima_channels': [lenz_sima_content, aio_sima_content, anten_sima_content,
                     tva_sima_content, fam_sima_content, telewebion_sima_content,
                     sepehr_sima_content, shima_sima_content, site_sima_content,],
            'radio_channels': [lenz_radio_content, aio_radio_content, anten_radio_content,
                     tva_radio_content, fam_radio_content, telewebion_radio_content,
                     sepehr_radio_content, shima_radio_content, site_radio_content,],
            'ostani_channels': [lenz_ostani_content, aio_ostani_content, anten_ostani_content,
                     tva_ostani_content, fam_ostani_content, telewebion_ostani_content,
                     sepehr_ostani_content, shima_ostani_content, site_ostani_content,],
            'ekhtesasi_channels': [lenz_ekhtesasi_content, aio_ekhtesasi_content, anten_ekhtesasi_content,
                     tva_ekhtesasi_content, fam_ekhtesasi_content, telewebion_ekhtesasi_content,
                     sepehr_ekhtesasi_content, shima_ekhtesasi_content, site_ekhtesasi_content,],
            'sima_visit': [lenz_sima_visit, aio_sima_visit, anten_sima_visit,
                     tva_sima_visit, fam_sima_visit, telewebion_sima_visit,
                     sepehr_sima_visit, shima_sima_visit, site_sima_visit,],
            'radio_visit': [lenz_radio_visit, aio_radio_visit, anten_radio_visit,
                     tva_radio_visit, fam_radio_visit, telewebion_radio_visit,
                     sepehr_radio_visit, shima_radio_visit, site_radio_visit,],
            'ostani_visit': [lenz_ostani_visit, aio_ostani_visit, anten_ostani_visit,
                     tva_ostani_visit, fam_ostani_visit, telewebion_ostani_visit,
                     sepehr_ostani_visit, shima_ostani_visit, site_ostani_visit,],
            'ekhtesasi_visit': [lenz_ekhtesasi_visit, aio_ekhtesasi_visit, anten_ekhtesasi_visit,
                     tva_ekhtesasi_visit, fam_ekhtesasi_visit, telewebion_ekhtesasi_visit,
                     sepehr_ekhtesasi_visit, shima_ekhtesasi_visit, site_ekhtesasi_visit,],
            'sima_duration': [lenz_sima_duration, aio_sima_duration, anten_sima_duration,
                     tva_sima_duration, fam_sima_duration, telewebion_sima_duration,
                     sepehr_sima_duration, shima_sima_duration, site_sima_duration,],
            'radio_duration': [lenz_radio_duration, aio_radio_duration, anten_radio_duration,
                     tva_radio_duration, fam_radio_duration, telewebion_radio_duration,
                     sepehr_radio_duration, shima_radio_duration, site_radio_duration,],
            'ostani_duration': [lenz_ostani_duration, aio_ostani_duration, anten_ostani_duration,
                     tva_ostani_duration, fam_ostani_duration, telewebion_ostani_duration,
                     sepehr_ostani_duration, shima_ostani_duration, site_ostani_duration,],
            'ekhtesasi_duration': [lenz_ekhtesasi_duration, aio_ekhtesasi_duration, anten_ekhtesasi_duration,
                     tva_ekhtesasi_duration, fam_ekhtesasi_duration, telewebion_ekhtesasi_duration,
                     sepehr_ekhtesasi_duration, shima_ekhtesasi_duration, site_ekhtesasi_duration,],}
    
    current_month_operator_data=pd.DataFrame(current_month_operator_data, columns=['operators', 'visit', 'register', 'active',
                                                                                     'sima_channels', 'radio_channels', 'ostani_channels', 'ekhtesasi_channels',
                                                                                     'sima_visit', 'radio_visit', 'ostani_visit', 'ekhtesasi_visit',
                                                                                     'sima_duration', 'radio_duration', 'ostani_duration', 'ekhtesasi_duration'])
    
    current_month_operator_data=current_month_operator_data.rename(columns={'operators': 'اپراتورها', 'visit': 'تعداد بازدید','register': 'تعداد کاربران ثبت نامی', 'active': 'تعداد کاربران فعال',
                                                                              'sima_channels': 'تعداد شبکه های سیما', 'radio_channels': 'تعداد شبکه های رادیو','ostani_channels': 'تعداد شبکه های استانی', 'ekhtesasi_channels': 'تعداد شبکه های اختصاصی',
                                                                              'sima_visit': 'تعداد بازدید از سیما', 'radio_visit': 'تعداد بازدید از رادیو','ostani_visit': 'تعداد بازدید از استانی', 'ekhtesasi_visit': 'تعداد بازدید از اختصاصی',
                                                                              'sima_duration': 'مدت زمان بازدید از سیما (به دقیقه)', 'radio_duration': 'مدت زمان بازدید از رادیو (به دقیقه)','ostani_duration': 'مدت زمان بازدید از استانی (به دقیقه)', 'ekhtesasi_duration': 'مدت زمان بازدید از اختصاصی (به دقیقه)'})
    
    current_month_all_data_summary=pd.DataFrame()
    current_month_all_data_summary={'parameters': ['تعداد بازدید', 'مدت زمان بازدید (به دقیقه)', 'تعداد محتوا', 'تعداد کاربران ثبت نامی' , 'تعداد کاربران فعال',],
           'statistics': [all_visit_current_month, all_duration_current_month,all_content_sima_current_month,
                          all_register_user_current_month, all_active_user_current_month,],}
    
    current_month_all_data_summary=pd.DataFrame(current_month_all_data_summary, columns=['parameters', 'statistics'])
    
    current_month_all_data_summary=current_month_all_data_summary.rename(columns={'parameters': 'پارامترها', 'statistics': 'آمار'})
    
    writer = pd.ExcelWriter('E:/hard/report/total EPG/EPG 1400/ماه اسفند 1400.xlsx', engine='xlsxwriter')
    current_month_sima_visit_channels.to_excel(writer, 'آمار بازدید شبکه های سیما')
    current_month_operator_data.to_excel(writer, 'آمار اپراتورها')
    current_month_all_data_summary.to_excel(writer, 'خلاصه آمار ماه اسفند')
    writer.save()
    
    return current_month_sima_visit_channels, current_month_operator_data, current_month_all_data_summary













