
def summary(all_data):
    
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
    import openpyxl
    
    print("start summary")
    all_visit=all_data['تعداد بازدید'].sum()
    all_duration=all_data['مدت بازدید'].sum()
    all_duration=round(all_duration, 0)
    all_data['channel'] = all_data['channel'].str.replace('قران', 'قرآن')
    all_data['channel'] = all_data['channel'].str.replace('جام جم 1', 'جام جم')
    all_data_pivot=all_data.groupby(['نام برنامه','channel', 'type', 'operator']).sum().reset_index()
    
    sima=all_data.query("type == 'سراسری'")
    radio=all_data.query("type == 'رادیویی'")
    ostani=all_data.query("type == 'استانی'")
    ekhtesasi=all_data.query("type == 'اختصاصی'")
    
    lenz=all_data.query("operator == 'لنز'")
    tva=all_data.query("operator == 'تیوا'")
    televebion=all_data.query("operator == 'تلوبیون'")
    aio=all_data.query("operator == 'آیو'")
    sepehr=all_data.query("operator == 'سپهر'")
    site_channels=all_data.query("operator == 'سایت شبکه ها'")
    
    print("sima summary")
    sima_all_visit=sima['تعداد بازدید'].sum()
    sima_all_duration=sima['مدت بازدید'].sum()
    sima_all_duration=round(sima_all_duration, 0)
    sima_content=sima.copy()
    sima_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    sima_all_content=len(sima_content)
    sima_channel=sima.copy()
    sima_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sima_all_channel=len(sima_channel)
    sima_operators = sima.copy()
    sima_operators.drop_duplicates(subset =['operator'], keep = 'first', inplace = True)
    sima_operators=len(sima_operators)
    
    print("radio summary")
    radio_all_visit=radio['تعداد بازدید'].sum()
    radio_all_duration=radio['مدت بازدید'].sum()
    radio_all_duration=round(radio_all_duration, 0)
    radio_content=radio.copy()
    radio_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_all_content=len(radio_content)
    radio_channel=radio.copy()
    radio_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    radio_all_channel=len(radio_channel)
    radio_operators = radio.copy()
    radio_operators.drop_duplicates(subset =['operator'], keep = 'first', inplace = True)
    radio_operators=len(radio_operators)
    
    print("ostani summary")
    ostani_all_visit=ostani['تعداد بازدید'].sum()
    ostani_all_duration=ostani['مدت بازدید'].sum()
    ostani_all_duration=round(ostani_all_duration, 0)
    ostani_content=ostani.copy()
    ostani_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_all_content=len(ostani_content)
    ostani_channel=ostani.copy()
    ostani_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ostani_all_channel=len(ostani_channel)
    ostani_operators = ostani.copy()
    ostani_operators.drop_duplicates(subset =['operator'], keep = 'first', inplace = True)
    ostani_operators=len(ostani_operators)
    
    print("ekhtesasi summary")
    ekhtesasi_all_visit=ekhtesasi['تعداد بازدید'].sum()
    ekhtesasi_all_duration=ekhtesasi['مدت بازدید'].sum()
    ekhtesasi_all_duration=round(ekhtesasi_all_duration, 0)
    ekhtesasi_content=ekhtesasi.copy()
    ekhtesasi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_all_content=len(ekhtesasi_content)
    ekhtesasi_channel=ekhtesasi.copy()
    ekhtesasi_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ekhtesasi_all_channel=len(ekhtesasi_channel)
    ekhtesasi_operators = ekhtesasi.copy()
    ekhtesasi_operators.drop_duplicates(subset =['operator'], keep = 'first', inplace = True)
    ekhtesasi_operators=len(ekhtesasi_operators)
    
    print("lenz summary")
    lenz_all_visit=lenz['تعداد بازدید'].sum()
    lenz_all_duration=lenz['مدت بازدید'].sum()
    lenz_all_duration=round(lenz_all_duration, 0)
    lenz_channel=lenz.copy()
    lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    lenz_all_channel=len(lenz_channel)
    
    print("tva summary")
    tva_all_visit=tva['تعداد بازدید'].sum()
    tva_all_duration=tva['مدت بازدید'].sum()
    tva_all_duration=round(tva_all_duration, 0)
    tva_channel=tva.copy()
    tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    tva_all_channel=len(tva_channel)
    
    print("televebion summary")
    televebion_all_visit=televebion['تعداد بازدید'].sum()
    televebion_all_duration=televebion['مدت بازدید'].sum()
    televebion_all_duration=round(televebion_all_duration, 0)
    televebion_channel=televebion.copy()
    televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    televebion_all_channel=len(televebion_channel)
    
    print("aio summary")
    aio_all_visit=aio['تعداد بازدید'].sum()
    aio_all_duration=aio['مدت بازدید'].sum()
    aio_all_duration=round(aio_all_duration, 0)
    aio_channel=aio.copy()
    aio_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    aio_all_channel=len(aio_channel)
    
    print("sepehr summary")
    sepehr_all_visit=sepehr['تعداد بازدید'].sum()
    sepehr_all_duration=sepehr['مدت بازدید'].sum()
    sepehr_all_duration=round(sepehr_all_duration, 0)
    sepehr_channel=sepehr.copy()
    sepehr_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sepehr_all_channel=len(sepehr_channel)
    
    print("site channels summary")
    site_channels_all_visit=site_channels['تعداد بازدید'].sum()
    site_channels_all_duration=site_channels['مدت بازدید'].sum()
    site_channels_all_duration=round(site_channels_all_duration, 0)
    site_channels_channel=site_channels.copy()
    site_channels_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    site_channels_all_channel=len(site_channels_channel)
    
    print("sima_operator summary")
    sima_lenz=sima.query("operator == 'لنز'")
    sima_lenz_all_visit=sima_lenz['تعداد بازدید'].sum()
    sima_lenz_all_duration=sima_lenz['مدت بازدید'].sum()
    sima_lenz_all_duration=round(sima_lenz_all_duration, 0)
    sima_lenz_channel=sima_lenz.copy()
    sima_lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sima_lenz_all_channel=len(sima_lenz_channel)
    
    sima_tva=sima.query("operator == 'تیوا'")
    sima_tva_all_visit=sima_tva['تعداد بازدید'].sum()
    sima_tva_all_duration=sima_tva['مدت بازدید'].sum()
    sima_tva_all_duration=round(sima_tva_all_duration, 0)
    sima_tva_channel=sima_tva.copy()
    sima_tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sima_tva_all_channel=len(sima_tva_channel)
    
    sima_televebion=sima.query("operator == 'تلوبیون'")
    sima_televebion_all_visit=sima_televebion['تعداد بازدید'].sum()
    sima_televebion_all_duration=sima_televebion['مدت بازدید'].sum()
    sima_televebion_all_duration=round(sima_televebion_all_duration, 0)
    sima_televebion_channel=sima_televebion.copy()
    sima_televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sima_televebion_all_channel=len(sima_televebion_channel)
    
    sima_aio=sima.query("operator == 'آیو'")
    sima_aio_all_visit=sima_aio['تعداد بازدید'].sum()
    sima_aio_all_duration=sima_aio['مدت بازدید'].sum()
    sima_aio_all_duration=round(sima_aio_all_duration, 0)
    sima_aio_channel=sima_aio.copy()
    sima_aio_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sima_aio_all_channel=len(sima_aio_channel)
    
    sima_sepehr=sima.query("operator == 'سپهر'")
    sima_sepehr_all_visit=sima_sepehr['تعداد بازدید'].sum()
    sima_sepehr_all_duration=sima_sepehr['مدت بازدید'].sum()
    sima_sepehr_all_duration=round(sima_sepehr_all_duration, 0)
    sima_sepehr_channel=sima_sepehr.copy()
    sima_sepehr_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sima_sepehr_all_channel=len(sima_sepehr_channel)
    
    sima_site_channels=sima.query("operator == 'سایت شبکه ها'")
    sima_site_channels_all_visit=sima_site_channels['تعداد بازدید'].sum()
    sima_site_channels_all_duration=sima_site_channels['مدت بازدید'].sum()
    sima_site_channels_all_duration=round(sima_site_channels_all_duration, 0)
    sima_site_channels_channel=sima_site_channels.copy()
    sima_site_channels_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    sima_site_channels_all_channel=len(sima_site_channels_channel)
    
    print("radio_operator summary")
    radio_lenz=radio.query("operator == 'لنز'")
    radio_lenz_all_visit=radio_lenz['تعداد بازدید'].sum()
    radio_lenz_all_duration=radio_lenz['مدت بازدید'].sum()
    radio_lenz_all_duration=round(radio_lenz_all_duration, 0)
    radio_lenz_channel=radio_lenz.copy()
    radio_lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    radio_lenz_all_channel=len(radio_lenz_channel)
    
    radio_tva=radio.query("operator == 'تیوا'")
    radio_tva_all_visit=radio_tva['تعداد بازدید'].sum()
    radio_tva_all_duration=radio_tva['مدت بازدید'].sum()
    radio_tva_all_duration=round(radio_tva_all_duration, 0)
    radio_tva_channel=radio_tva.copy()
    radio_tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    radio_tva_all_channel=len(radio_tva_channel)
    
    radio_televebion=radio.query("operator == 'تلوبیون'")
    radio_televebion_all_visit=radio_televebion['تعداد بازدید'].sum()
    radio_televebion_all_duration=radio_televebion['مدت بازدید'].sum()
    radio_televebion_all_duration=round(radio_televebion_all_duration, 0)
    radio_televebion_channel=radio_televebion.copy()
    radio_televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    radio_televebion_all_channel=len(radio_televebion_channel)
    
    radio_aio=radio.query("operator == 'آیو'")
    radio_aio_all_visit=radio_aio['تعداد بازدید'].sum()
    radio_aio_all_duration=radio_aio['مدت بازدید'].sum()
    radio_aio_all_duration=round(radio_aio_all_duration, 0)
    radio_aio_channel=radio_aio.copy()
    radio_aio_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    radio_aio_all_channel=len(radio_aio_channel)
    
    radio_sepehr=radio.query("operator == 'سپهر'")
    radio_sepehr_all_visit=radio_sepehr['تعداد بازدید'].sum()
    radio_sepehr_all_duration=radio_sepehr['مدت بازدید'].sum()
    radio_sepehr_all_duration=round(radio_sepehr_all_duration, 0)
    radio_sepehr_channel=radio_sepehr.copy()
    radio_sepehr_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    radio_sepehr_all_channel=len(radio_sepehr_channel)
    
    radio_site_channels=radio.query("operator == 'سایت شبکه ها'")
    radio_site_channels_all_visit=radio_site_channels['تعداد بازدید'].sum()
    radio_site_channels_all_duration=radio_site_channels['مدت بازدید'].sum()
    radio_site_channels_all_duration=round(radio_site_channels_all_duration, 0)
    radio_site_channels_channel=radio_site_channels.copy()
    radio_site_channels_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    radio_site_channels_all_channel=len(radio_site_channels_channel)
    
    print("ostani_operator summary")
    ostani_lenz=ostani.query("operator == 'لنز'")
    ostani_lenz_all_visit=ostani_lenz['تعداد بازدید'].sum()
    ostani_lenz_all_duration=ostani_lenz['مدت بازدید'].sum()
    ostani_lenz_all_duration=round(ostani_lenz_all_duration, 0)
    ostani_lenz_channel=ostani_lenz.copy()
    ostani_lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ostani_lenz_all_channel=len(ostani_lenz_channel)
    
    ostani_tva=ostani.query("operator == 'تیوا'")
    ostani_tva_all_visit=ostani_tva['تعداد بازدید'].sum()
    ostani_tva_all_duration=ostani_tva['مدت بازدید'].sum()
    ostani_tva_all_duration=round(ostani_tva_all_duration, 0)
    ostani_tva_channel=ostani_tva.copy()
    ostani_tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ostani_tva_all_channel=len(ostani_tva_channel)
    
    ostani_televebion=ostani.query("operator == 'تلوبیون'")
    ostani_televebion_all_visit=ostani_televebion['تعداد بازدید'].sum()
    ostani_televebion_all_duration=ostani_televebion['مدت بازدید'].sum()
    ostani_televebion_all_duration=round(ostani_televebion_all_duration, 0)
    ostani_televebion_channel=ostani_televebion.copy()
    ostani_televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ostani_televebion_all_channel=len(ostani_televebion_channel)
    
    ostani_aio=ostani.query("operator == 'آیو'")
    ostani_aio_all_visit=ostani_aio['تعداد بازدید'].sum()
    ostani_aio_all_duration=ostani_aio['مدت بازدید'].sum()
    ostani_aio_all_duration=round(ostani_aio_all_duration, 0)
    ostani_aio_channel=ostani_aio.copy()
    ostani_aio_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ostani_aio_all_channel=len(ostani_aio_channel)
    
    ostani_sepehr=ostani.query("operator == 'سپهر'")
    ostani_sepehr_all_visit=ostani_sepehr['تعداد بازدید'].sum()
    ostani_sepehr_all_duration=ostani_sepehr['مدت بازدید'].sum()
    ostani_sepehr_all_duration=round(ostani_sepehr_all_duration, 0)
    ostani_sepehr_channel=ostani_sepehr.copy()
    ostani_sepehr_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ostani_sepehr_all_channel=len(ostani_sepehr_channel)
    
    ostani_site_channels=ostani.query("operator == 'سایت شبکه ها'")
    ostani_site_channels_all_visit=ostani_site_channels['تعداد بازدید'].sum()
    ostani_site_channels_all_duration=ostani_site_channels['مدت بازدید'].sum()
    ostani_site_channels_all_duration=round(ostani_site_channels_all_duration, 0)
    ostani_site_channels_channel=ostani_site_channels.copy()
    ostani_site_channels_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ostani_site_channels_all_channel=len(ostani_site_channels_channel)
    
    print("ekhtesasi_operator summary")
    ekhtesasi_lenz=ekhtesasi.query("operator == 'لنز'")
    ekhtesasi_lenz_all_visit=ekhtesasi_lenz['تعداد بازدید'].sum()
    ekhtesasi_lenz_all_duration=ekhtesasi_lenz['مدت بازدید'].sum()
    ekhtesasi_lenz_all_duration=round(ekhtesasi_lenz_all_duration, 0)
    ekhtesasi_lenz_channel=ekhtesasi_lenz.copy()
    ekhtesasi_lenz_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ekhtesasi_lenz_all_channel=len(ekhtesasi_lenz_channel)
    
    ekhtesasi_tva=ekhtesasi.query("operator == 'تیوا'")
    ekhtesasi_tva_all_visit=ekhtesasi_tva['تعداد بازدید'].sum()
    ekhtesasi_tva_all_duration=ekhtesasi_tva['مدت بازدید'].sum()
    ekhtesasi_tva_all_duration=ekhtesasi_tva_all_duration
    ekhtesasi_tva_channel=ekhtesasi_tva.copy()
    ekhtesasi_tva_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_all_channel=len(ekhtesasi_tva_channel)
    
    ekhtesasi_televebion=ekhtesasi.query("operator == 'تلوبیون'")
    ekhtesasi_televebion_all_visit=ekhtesasi_televebion['تعداد بازدید'].sum()
    ekhtesasi_televebion_all_duration=ekhtesasi_televebion['مدت بازدید'].sum()
    ekhtesasi_televebion_all_duration=round(ekhtesasi_televebion_all_duration, 0)
    ekhtesasi_televebion_channel=ekhtesasi_televebion.copy()
    ekhtesasi_televebion_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ekhtesasi_televebion_all_channel=len(ekhtesasi_televebion_channel)
    
    ekhtesasi_aio=ekhtesasi.query("operator == 'آیو'")
    ekhtesasi_aio_all_visit=ekhtesasi_aio['تعداد بازدید'].sum()
    ekhtesasi_aio_all_duration=ekhtesasi_aio['مدت بازدید'].sum()
    ekhtesasi_aio_all_duration=round(ekhtesasi_aio_all_duration, 0)
    ekhtesasi_aio_channel=ekhtesasi_aio.copy()
    ekhtesasi_aio_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ekhtesasi_aio_all_channel=len(ekhtesasi_aio_channel)
    
    ekhtesasi_sepehr=ekhtesasi.query("operator == 'سپهر'")
    ekhtesasi_sepehr_all_visit=ekhtesasi_sepehr['تعداد بازدید'].sum()
    ekhtesasi_sepehr_all_duration=ekhtesasi_sepehr['مدت بازدید'].sum()
    ekhtesasi_sepehr_all_duration=round(ekhtesasi_sepehr_all_duration, 0)
    ekhtesasi_sepehr_channel=ekhtesasi_sepehr.copy()
    ekhtesasi_sepehr_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ekhtesasi_sepehr_all_channel=len(ekhtesasi_sepehr_channel)
    
    ekhtesasi_site_channels=ekhtesasi.query("operator == 'سایت شبکه ها'")
    ekhtesasi_site_channels_all_visit=ekhtesasi_site_channels['تعداد بازدید'].sum()
    ekhtesasi_site_channels_all_duration=ekhtesasi_site_channels['مدت بازدید'].sum()
    ekhtesasi_site_channels_all_duration=round(ekhtesasi_site_channels_all_duration, 0)
    ekhtesasi_site_channels_channel=ekhtesasi_site_channels.copy()
    ekhtesasi_site_channels_channel.drop_duplicates(subset =['channel'], keep = 'first', inplace = True)
    ekhtesasi_site_channels_all_channel=len(ekhtesasi_site_channels_channel)

    
    
    print("dataframe summary")
    data_summary_service={'field': ['تعداد محتوا', 'تعداد بازدید', 'مدت بازدید (به دقیقه)','تعداد شبکه','تعداد اپراتور'],
           'sima': [sima_all_content, sima_all_visit, sima_all_duration, sima_all_channel, sima_operators],
           'radio': [radio_all_content, radio_all_visit, radio_all_duration, radio_all_channel, radio_operators],
           'ostani': [ostani_all_content, ostani_all_visit, ostani_all_duration, ostani_all_channel, ostani_operators],
           'ekhtesasi': [ekhtesasi_all_content, ekhtesasi_all_visit, ekhtesasi_all_duration, ekhtesasi_all_channel, ekhtesasi_operators],}
    data_summary_service=pd.DataFrame(data_summary_service, columns=['field', 'sima', 'radio', 'ostani', 'ekhtesasi'])
    data_summary_service=data_summary_service.rename(columns={'field': 'حوزه', 'sima': 'سیما', 'radio': 'رادیویی' , 'ekhtesasi': 'اختصاصی', 'ostani': 'استانی'})
    
    data_summary_operator={'operator': ['تعداد شبکه', 'تعداد بازدید', 'مدت بازدید (به دقیقه)'],
           'lenz': [lenz_all_channel, lenz_all_visit, lenz_all_duration],
           'tva': [tva_all_channel, tva_all_visit, tva_all_duration],
           'televebion': [televebion_all_channel, televebion_all_visit, televebion_all_duration],
           'aio': [aio_all_channel, aio_all_visit, aio_all_duration],
           'sepehr': [sepehr_all_channel, sepehr_all_visit, sepehr_all_duration],
           'site_channels': [site_channels_all_channel, site_channels_all_visit, site_channels_all_duration],}
    data_summary_operator=pd.DataFrame(data_summary_operator, columns=['operator', 'lenz', 'tva', 'televebion', 'aio', 'sepehr', 'site_channels'])
    data_summary_operator=data_summary_operator.rename(columns={'operator': 'اپراتور', 'lenz': 'لنز', 'tva': 'تیوا', 'televebion': 'تلوبیون', 'aio': 'آیو', 'sepehr': 'سپهر', 'site_channels': 'سایت شبکه ها'})
    
    data_summary_service_operator={'parameters': ['تعداد شبکه ', 'تعداد بازدید', 'مدت بازدید (به دقیقه)'],
           'sima_lenz': [sima_lenz_all_channel, sima_lenz_all_visit, sima_lenz_all_duration],
           'sima_tva': [sima_tva_all_channel, sima_tva_all_visit, sima_tva_all_duration],
           'sima_televebion': [sima_televebion_all_channel, sima_televebion_all_visit, sima_televebion_all_duration],
           'sima_aio': [sima_aio_all_channel, sima_aio_all_visit, sima_aio_all_duration],
           'sima_sepehr': [sima_sepehr_all_channel, sima_sepehr_all_visit, sima_sepehr_all_duration],
           'sima_site_channels': [sima_site_channels_all_channel, sima_site_channels_all_visit, sima_site_channels_all_duration],
           'radio_lenz': [radio_lenz_all_channel, radio_lenz_all_visit, radio_lenz_all_duration],
           'radio_tva': [radio_tva_all_channel, radio_tva_all_visit, radio_tva_all_duration],
           'radio_televebion': [radio_televebion_all_channel, radio_televebion_all_visit, radio_televebion_all_duration],
           'radio_aio': [radio_aio_all_channel, radio_aio_all_visit, radio_aio_all_duration],
           'radio_sepehr': [radio_sepehr_all_channel, radio_sepehr_all_visit, radio_sepehr_all_duration],
           'radio_site_channels': [radio_site_channels_all_channel, radio_site_channels_all_visit, radio_site_channels_all_duration],
           'ostani_lenz': [ostani_lenz_all_channel, ostani_lenz_all_visit, ostani_lenz_all_duration],
           'ostani_tva': [ostani_tva_all_channel, ostani_tva_all_visit, ostani_tva_all_duration],
           'ostani_televebion': [ostani_televebion_all_channel, ostani_televebion_all_visit, ostani_televebion_all_duration],
           'ostani_aio': [ostani_aio_all_channel, ostani_aio_all_visit, ostani_aio_all_duration],
           'ostani_sepehr': [ostani_sepehr_all_channel, ostani_sepehr_all_visit, ostani_sepehr_all_duration],
           'ostani_site_channels': [ostani_site_channels_all_channel, ostani_site_channels_all_visit, ostani_site_channels_all_duration],
           'ekhtesasi_lenz': [ekhtesasi_lenz_all_channel, ekhtesasi_lenz_all_visit, ekhtesasi_lenz_all_duration],
           'ekhtesasi_tva': [ekhtesasi_tva_all_channel, ekhtesasi_tva_all_visit, ekhtesasi_tva_all_duration],
           'ekhtesasi_televebion': [ekhtesasi_televebion_all_channel, ekhtesasi_televebion_all_visit, ekhtesasi_televebion_all_duration],
           'ekhtesasi_aio': [ekhtesasi_aio_all_channel, ekhtesasi_aio_all_visit, ekhtesasi_aio_all_duration],
           'ekhtesasi_sepehr': [ekhtesasi_sepehr_all_channel, ekhtesasi_sepehr_all_visit, ekhtesasi_sepehr_all_duration],
           'ekhtesasi_site_channels': [ekhtesasi_site_channels_all_channel, ekhtesasi_site_channels_all_visit, ekhtesasi_site_channels_all_duration],}
    data_summary_service_operator=pd.DataFrame(data_summary_service_operator, columns=['parameters', 'sima_lenz', 'sima_tva','sima_televebion','sima_aio','sima_sepehr','sima_site_channels',
                                                                                       'radio_lenz', 'radio_tva','radio_televebion','radio_aio','radio_sepehr','radio_site_channels',
                                                                                       'ostani_lenz', 'ostani_tva','ostani_televebion','ostani_aio','ostani_sepehr','ostani_site_channels',
                                                                                       'ekhtesasi_lenz', 'ekhtesasi_tva','ekhtesasi_televebion','ekhtesasi_aio','ekhtesasi_sepehr','ekhtesasi_site_channels'])
    
    data_summary_service_operator=data_summary_service_operator.rename(columns={'parameters': 'پارامترها', 
                                                                                'sima_lenz': 'سیما-لنز', 'sima_tva': 'سیما-تیوا', 'sima_televebion': 'سیما-تلوبیون','sima_aio': 'سیما-آیو','sima_sepehr': 'سیما-سپهر','sima_site_channels': 'سیما-سایت شبکه ها',
                                                                                'radio_lenz': 'رادیو-لنز', 'radio_tva': 'رادیو-تیوا', 'radio_televebion': 'رادیو-تلوبیون','radio_aio': 'رادیو-آیو','radio_sepehr': 'رادیو-سپهر','radio_site_channels': 'رادیو-سایت شبکه ها',
                                                                                'ostani_lenz': 'استانی-لنز', 'ostani_tva': 'استانی-تیوا', 'ostani_televebion': 'استانی-تلوبیون','ostani_aio': 'استانی-آیو','ostani_sepehr': 'استانی-سپهر','ostani_site_channels': 'استانی-سایت شبکه ها',
                                                                                'ekhtesasi_lenz': 'اختصاصی-لنز', 'ekhtesasi_tva': 'اختصاصی-تیوا', 'ekhtesasi_televebion': 'اختصاصی-تلوبیون','ekhtesasi_aio': 'اختصاصی-آیو','ekhtesasi_sepehr': 'اختصاصی-سپهر','ekhtesasi_site_channels': 'اختصاصی-سایت شبکه ها'})
    
#    writer = pd.ExcelWriter('output/آمار ماه جاری/خلاصه آمار.xlsx', engine='xlsxwriter')
#    data_summary_service.to_excel(writer, 'آمار انواع سرویس های سازمان')
#    data_summary_operator.to_excel(writer, 'آمار اپراتورها')
#    data_summary_service_operator.to_excel(writer, 'اپراتورها و سرویسهای سازمان')
#    writer.save()
    
#    writer = pd.ExcelWriter('output/moh.rast/خلاصه آمار.xlsx', engine='xlsxwriter')
#    data_summary_service.to_excel(writer, 'آمار انواع سرویس های سازمان')
#    data_summary_operator.to_excel(writer, 'آمار اپراتورها')
#    data_summary_service_operator.to_excel(writer, 'اپراتورها و سرویسهای سازمان')
#    writer.save()
    
    writer = pd.ExcelWriter('output/zomorrodi/خلاصه آمار.xlsx', engine='xlsxwriter')
    data_summary_service.to_excel(writer, 'آمار انواع سرویس های سازمان', index = False)
    data_summary_operator.to_excel(writer, 'آمار اپراتورها', index = False)
    data_summary_service_operator.to_excel(writer, 'اپراتورها و سرویسهای سازمان', index = False)
    writer.save()
    
    writer = pd.ExcelWriter('output/pedram/خلاصه آمار.xlsx', engine='xlsxwriter')
    data_summary_service.to_excel(writer, 'آمار انواع سرویس های سازمان', index = False)
    data_summary_operator.to_excel(writer, 'آمار اپراتورها', index = False)
    data_summary_service_operator.to_excel(writer, 'اپراتورها و سرویسهای سازمان', index = False)
    writer.save()
    
    writer = pd.ExcelWriter('output/output.sending.hard/خلاصه آمار.xlsx', engine='xlsxwriter')
    data_summary_service.to_excel(writer, 'آمار انواع سرویس های سازمان', index = False)
    data_summary_operator.to_excel(writer, 'آمار اپراتورها', index = False)
    data_summary_service_operator.to_excel(writer, 'اپراتورها و سرویسهای سازمان', index = False)
    writer.save()
    
    print("End summary")
    
    
    return data_summary_service, data_summary_operator, data_summary_service_operator

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    