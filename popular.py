
def popular(sima, ekhtesasi, ostani, radio, boronmarzi):
    
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
    
    
    print("sima popular contents") 
    sima_popular_content=sima.copy()
    sima_popular_content=sima_popular_content.query("operator != 'سایت شبکه ها'")
    sima_popular_content=sima_popular_content.query("operator != 'سپهر'")
    sima_popular_content=sima_popular_content.groupby(['نام برنامه','channel']).sum().reset_index()
    sima_popular_content.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    sima_popular_content_visit=sima_popular_content.iloc[0:10 , [0, 2]]
    sima_popular_content_visit = sima_popular_content_visit.reset_index()
    del sima_popular_content_visit['index']
    
    sima_popular_content.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    sima_popular_content_duration=sima_popular_content.iloc[0:10 , [0, 4]]
    sima_popular_content_duration = sima_popular_content_duration.reset_index()
    del sima_popular_content_duration['index']
    
    sima_popular_content=pd.DataFrame()
    sima_popular_content=pd.concat([sima_popular_content_visit, sima_popular_content_duration], axis=1)
    
    print("ekhtesasi popular contents") 
    ekhtesasi_popular_content=ekhtesasi.copy()
    ekhtesasi_popular_content=ekhtesasi_popular_content.groupby(['نام برنامه','channel']).sum().reset_index()
    ekhtesasi_popular_content.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_popular_content_visit=ekhtesasi_popular_content.iloc[0:10 , [0, 2]]
    ekhtesasi_popular_content_visit = ekhtesasi_popular_content_visit.reset_index()
    del ekhtesasi_popular_content_visit['index']
    
    ekhtesasi_popular_content.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_popular_content_duration=ekhtesasi_popular_content.iloc[0:10 , [0, 4]]
    ekhtesasi_popular_content_duration = ekhtesasi_popular_content_duration.reset_index()
    del ekhtesasi_popular_content_duration['index']
    
    ekhtesasi_popular_content=pd.DataFrame()
    ekhtesasi_popular_content=pd.concat([ekhtesasi_popular_content_visit, ekhtesasi_popular_content_duration], axis=1)
    
    print("boronmarzi popular contents") 
    boronmarzi_popular_content=boronmarzi.copy()
    boronmarzi_popular_content=boronmarzi_popular_content.groupby(['نام برنامه','channel']).sum().reset_index()
    boronmarzi_popular_content.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_popular_content_visit=boronmarzi_popular_content.iloc[0:10 , [0, 2]]
    boronmarzi_popular_content_visit = boronmarzi_popular_content_visit.reset_index()
    del boronmarzi_popular_content_visit['index']
    
    boronmarzi_popular_content.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_popular_content_duration=boronmarzi_popular_content.iloc[0:10 , [0, 4]]
    boronmarzi_popular_content_duration = boronmarzi_popular_content_duration.reset_index()
    del boronmarzi_popular_content_duration['index']
    
    boronmarzi_popular_content=pd.DataFrame()
    boronmarzi_popular_content=pd.concat([boronmarzi_popular_content_visit, boronmarzi_popular_content_duration], axis=1)
    
    print("ostani popular contents") 
    ostani_popular_content=ostani.copy()
    ostani_popular_content=ostani_popular_content.groupby(['نام برنامه','channel']).sum().reset_index()
    ostani_popular_content111 = ostani_popular_content.copy()
    ostani_popular_content.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_popular_content_visit=ostani_popular_content.iloc[0:10 , [0, 2]]
    ostani_popular_content_visit = ostani_popular_content_visit.reset_index()
    del ostani_popular_content_visit['index']
    
    ostani_popular_content.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_popular_content_duration=ostani_popular_content.iloc[0:10 , [0, 4]]
    ostani_popular_content_duration = ostani_popular_content_duration.reset_index()
    del ostani_popular_content_duration['index']
    
    ostani_popular_content=pd.DataFrame()
    ostani_popular_content=pd.concat([ostani_popular_content_visit, ostani_popular_content_duration], axis=1)
    
    print("radio popular contents") 
    radio_popular_content=radio.copy()
    radio_popular_content=radio_popular_content.groupby(['نام برنامه','channel']).sum().reset_index()
    radio_popular_content.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_popular_content_visit=radio_popular_content.iloc[0:10 , [0, 2]]
    radio_popular_content_visit = radio_popular_content_visit.reset_index()
    del radio_popular_content_visit['index']
    
    radio_popular_content.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_popular_content_duration=radio_popular_content.iloc[0:10 , [0, 4]]
    radio_popular_content_duration = radio_popular_content_duration.reset_index()
    del radio_popular_content_duration['index']
    
    radio_popular_content=pd.DataFrame()
    radio_popular_content=pd.concat([radio_popular_content_visit, radio_popular_content_duration], axis=1)
    

    writer = pd.ExcelWriter('output/zomorrodi/محتواهای پربازدید.xlsx', engine='xlsxwriter')
    sima_popular_content.to_excel(writer, 'سیما', index=False)
#    ekhtesasi_popular_content.to_excel(writer, 'اختصاصی', index=False)
#    ostani_popular_content.to_excel(writer, 'استانی', index=False)
#    radio_popular_content.to_excel(writer, 'رادیویی', index=False)
    writer.save()
    

    writer = pd.ExcelWriter('output/output.sending.hard/محتواهای پربازدید.xlsx', engine='xlsxwriter')
    sima_popular_content.to_excel(writer, 'سیما', index=False)
    ekhtesasi_popular_content.to_excel(writer, 'اختصاصی', index=False)
    boronmarzi_popular_content.to_excel(writer, 'برونمرزی', index=False)
    ostani_popular_content.to_excel(writer, 'استانی', index=False)
    radio_popular_content.to_excel(writer, 'رادیویی', index=False)
    writer.save()
    
    return sima_popular_content, ekhtesasi_popular_content, ostani_popular_content, radio_popular_content, boronmarzi_popular_content

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    