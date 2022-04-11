    
def boronmarzi_data(boronmarzi):
            
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
        
        
    print("start boronmarzi")
    
    print("boronmarzi_alalam")
    boronmarzi_alalam=boronmarzi.query("channel == 'العالم'")
    boronmarzi_alalam_visit=boronmarzi_alalam['تعداد بازدید'].sum()
    boronmarzi_alalam_duration=boronmarzi_alalam['مدت بازدید'].sum()
    boronmarzi_alalam_duration=round(boronmarzi_alalam_duration, 0)
    boronmarzi_alalam_content=boronmarzi_alalam.copy()
    boronmarzi_alalam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_alalam_content=len(boronmarzi_alalam_content)
    boronmarzi_alalam_pivot=boronmarzi_alalam.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_alalam_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_alalam_popular_visit=boronmarzi_alalam_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_alalam_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_alalam_popular_duration=boronmarzi_alalam_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_alalam_popular_visit = boronmarzi_alalam_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه العالم', 'نام برنامه': 'محتواهای پربازدید شبکه العالم'})
    boronmarzi_alalam_popular_duration = boronmarzi_alalam_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه العالم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه العالم'})
    
    print("boronmarzi_alkosar")
    boronmarzi_alkosar=boronmarzi.query("channel == 'الکوثر'")
    boronmarzi_alkosar_visit=boronmarzi_alkosar['تعداد بازدید'].sum()
    boronmarzi_alkosar_duration=boronmarzi_alkosar['مدت بازدید'].sum()
    boronmarzi_alkosar_duration=round(boronmarzi_alkosar_duration, 0)
    boronmarzi_alkosar_content=boronmarzi_alkosar.copy()
    boronmarzi_alkosar_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_alkosar_content=len(boronmarzi_alkosar_content)
    boronmarzi_alkosar_pivot=boronmarzi_alkosar.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_alkosar_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_alkosar_popular_visit=boronmarzi_alkosar_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_alkosar_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_alkosar_popular_duration=boronmarzi_alkosar_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_alkosar_popular_visit = boronmarzi_alkosar_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه الکوثر', 'نام برنامه': 'محتواهای پربازدید شبکه الکوثر'})
    boronmarzi_alkosar_popular_duration = boronmarzi_alkosar_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه الکوثر (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه الکوثر'})
    
    print("boronmarzi_presstv")
    boronmarzi_presstv=boronmarzi.query("channel == 'پرس تی وی'")
    boronmarzi_presstv_visit=boronmarzi_presstv['تعداد بازدید'].sum()
    boronmarzi_presstv_duration=boronmarzi_presstv['مدت بازدید'].sum()
    boronmarzi_presstv_duration=round(boronmarzi_presstv_duration, 0)
    boronmarzi_presstv_content=boronmarzi_presstv.copy()
    boronmarzi_presstv_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_presstv_content=len(boronmarzi_presstv_content)
    boronmarzi_presstv_pivot=boronmarzi_presstv.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_presstv_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_presstv_popular_visit=boronmarzi_presstv_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_presstv_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_presstv_popular_duration=boronmarzi_presstv_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_presstv_popular_visit = boronmarzi_presstv_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه پرس تی وی', 'نام برنامه': 'محتواهای پربازدید شبکه پرس تی وی'})
    boronmarzi_presstv_popular_duration = boronmarzi_presstv_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه پرس تی وی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه پرس تی وی'})
    
    print("boronmarzi_ifilm")
    boronmarzi_ifilm=boronmarzi.query("channel == 'آی فیلم'")
    boronmarzi_ifilm_visit=boronmarzi_ifilm['تعداد بازدید'].sum()
    boronmarzi_ifilm_duration=boronmarzi_ifilm['مدت بازدید'].sum()
    boronmarzi_ifilm_duration=round(boronmarzi_ifilm_duration, 0)
    boronmarzi_ifilm_content=boronmarzi_ifilm.copy()
    boronmarzi_ifilm_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_ifilm_content=len(boronmarzi_ifilm_content)
    boronmarzi_ifilm_pivot=boronmarzi_ifilm.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_ifilm_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_ifilm_popular_visit=boronmarzi_ifilm_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_ifilm_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_ifilm_popular_duration=boronmarzi_ifilm_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_ifilm_popular_visit = boronmarzi_ifilm_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آی فیلم', 'نام برنامه': 'محتواهای پربازدید شبکه آی فیلم'})
    boronmarzi_ifilm_popular_duration = boronmarzi_ifilm_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه آی فیلم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه آی فیلم'})
    
    print("boronmarzi_sahar_kordi")
    boronmarzi_sahar_kordi=boronmarzi.query("channel == 'سحر کردی'")
    boronmarzi_sahar_kordi_visit=boronmarzi_sahar_kordi['تعداد بازدید'].sum()
    boronmarzi_sahar_kordi_duration=boronmarzi_sahar_kordi['مدت بازدید'].sum()
    boronmarzi_sahar_kordi_duration=round(boronmarzi_sahar_kordi_duration, 0)
    boronmarzi_sahar_kordi_content=boronmarzi_sahar_kordi.copy()
    boronmarzi_sahar_kordi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_sahar_kordi_content=len(boronmarzi_sahar_kordi_content)
    boronmarzi_sahar_kordi_pivot=boronmarzi_sahar_kordi.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_sahar_kordi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_sahar_kordi_popular_visit=boronmarzi_sahar_kordi_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_sahar_kordi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_sahar_kordi_popular_duration=boronmarzi_sahar_kordi_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_sahar_kordi_popular_visit = boronmarzi_sahar_kordi_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سحر کردی', 'نام برنامه': 'محتواهای پربازدید شبکه سحر کردی'})
    boronmarzi_sahar_kordi_popular_duration = boronmarzi_sahar_kordi_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه سحر کردی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه سحر کردی'})
    
    print("boronmarzi_sahar_balkan")
    boronmarzi_sahar_balkan=boronmarzi.query("channel == 'سحر بالکان'")
    boronmarzi_sahar_balkan_visit=boronmarzi_sahar_balkan['تعداد بازدید'].sum()
    boronmarzi_sahar_balkan_duration=boronmarzi_sahar_balkan['مدت بازدید'].sum()
    boronmarzi_sahar_balkan_duration=round(boronmarzi_sahar_balkan_duration, 0)
    boronmarzi_sahar_balkan_content=boronmarzi_sahar_balkan.copy()
    boronmarzi_sahar_balkan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_sahar_balkan_content=len(boronmarzi_sahar_balkan_content)
    boronmarzi_sahar_balkan_pivot=boronmarzi_sahar_balkan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_sahar_balkan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_sahar_balkan_popular_visit=boronmarzi_sahar_balkan_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_sahar_balkan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_sahar_balkan_popular_duration=boronmarzi_sahar_balkan_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_sahar_balkan_popular_visit = boronmarzi_sahar_balkan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سحر بالکان', 'نام برنامه': 'محتواهای پربازدید شبکه سحر بالکان'})
    boronmarzi_sahar_balkan_popular_duration = boronmarzi_sahar_balkan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه سحر بالکان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه سحر بالکان'})
    
    print("boronmarzi_jamejam1")
    boronmarzi_jamejam1=boronmarzi.query("channel == 'جام جم 1'")
    boronmarzi_jamejam1_visit=boronmarzi_jamejam1['تعداد بازدید'].sum()
    boronmarzi_jamejam1_duration=boronmarzi_jamejam1['مدت بازدید'].sum()
    boronmarzi_jamejam1_duration=round(boronmarzi_jamejam1_duration, 0)
    boronmarzi_jamejam1_content=boronmarzi_jamejam1.copy()
    boronmarzi_jamejam1_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_jamejam1_content=len(boronmarzi_jamejam1_content)
    boronmarzi_jamejam1_pivot=boronmarzi_jamejam1.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_jamejam1_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_jamejam1_popular_visit=boronmarzi_jamejam1_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_jamejam1_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_jamejam1_popular_duration=boronmarzi_jamejam1_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_jamejam1_popular_visit = boronmarzi_jamejam1_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه جام جم 1', 'نام برنامه': 'محتواهای پربازدید شبکه جام جم 1'})
    boronmarzi_jamejam1_popular_duration = boronmarzi_jamejam1_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه جام جم 1 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه جام جم 1'})
    
    print("boronmarzi_sahar_azari")
    boronmarzi_sahar_azari=boronmarzi.query("channel == 'سحر آذری'")
    boronmarzi_sahar_azari_visit=boronmarzi_sahar_azari['تعداد بازدید'].sum()
    boronmarzi_sahar_azari_duration=boronmarzi_sahar_azari['مدت بازدید'].sum()
    boronmarzi_sahar_azari_duration=round(boronmarzi_sahar_azari_duration, 0)
    boronmarzi_sahar_azari_content=boronmarzi_sahar_azari.copy()
    boronmarzi_sahar_azari_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_sahar_azari_content=len(boronmarzi_sahar_azari_content)
    boronmarzi_sahar_azari_pivot=boronmarzi_sahar_azari.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_sahar_azari_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_sahar_azari_popular_visit=boronmarzi_sahar_azari_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_sahar_azari_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_sahar_azari_popular_duration=boronmarzi_sahar_azari_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_sahar_azari_popular_visit = boronmarzi_sahar_azari_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سحر آذری', 'نام برنامه': 'محتواهای پربازدید شبکه سحر آذری'})
    boronmarzi_sahar_azari_popular_duration = boronmarzi_sahar_azari_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه سحر آذری (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه سحر آذری'})
    
    print("boronmarzi_sahar_ordu")
    boronmarzi_sahar_ordu=boronmarzi.query("channel == 'سحر اردو'")
    boronmarzi_sahar_ordu_visit=boronmarzi_sahar_ordu['تعداد بازدید'].sum()
    boronmarzi_sahar_ordu_duration=boronmarzi_sahar_ordu['مدت بازدید'].sum()
    boronmarzi_sahar_ordu_duration=round(boronmarzi_sahar_ordu_duration, 0)
    boronmarzi_sahar_ordu_content=boronmarzi_sahar_ordu.copy()
    boronmarzi_sahar_ordu_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    boronmarzi_sahar_ordu_content=len(boronmarzi_sahar_ordu_content)
    boronmarzi_sahar_ordu_pivot=boronmarzi_sahar_ordu.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    boronmarzi_sahar_ordu_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_sahar_ordu_popular_visit=boronmarzi_sahar_ordu_pivot.iloc[0:10 , [0, 3]]
    boronmarzi_sahar_ordu_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_sahar_ordu_popular_duration=boronmarzi_sahar_ordu_pivot.iloc[0:10 , [0, 5]]
    
    boronmarzi_sahar_ordu_popular_visit = boronmarzi_sahar_ordu_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سحر اردو', 'نام برنامه': 'محتواهای پربازدید شبکه سحر اردو'})
    boronmarzi_sahar_ordu_popular_duration = boronmarzi_sahar_ordu_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه سحر اردو (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه سحر اردو'})
    
   
    print("dataframe boronmarzi channels")
    boronmarzi_channels_statistics={'channel_name': ['برونمرزی العالم',
                                                    'برونمرزی الکوثر',
                                                    'برونمرزی پرس تی وی',
                                                    'برونمرزی آی فیلم',
                                         'برونمرزی سحر کردی',
                                         'برونمرزی سحر بالکان',
                                         'برونمرزی جام جم 1',
                                         'برونمرزی سحر آذری',
                                         'برونمرزی سحر اردو',],
           'channel_content': [boronmarzi_alalam_content, boronmarzi_alkosar_content, boronmarzi_presstv_content,
                               boronmarzi_ifilm_content, boronmarzi_sahar_kordi_content, boronmarzi_sahar_balkan_content,
                               boronmarzi_jamejam1_content, boronmarzi_sahar_azari_content, boronmarzi_sahar_ordu_content,],
           'channel_visit': [boronmarzi_alalam_visit, boronmarzi_alkosar_visit, boronmarzi_presstv_visit,
                               boronmarzi_ifilm_visit, boronmarzi_sahar_kordi_visit, boronmarzi_sahar_balkan_visit,
                               boronmarzi_jamejam1_visit, boronmarzi_sahar_azari_visit, boronmarzi_sahar_ordu_visit,],
           'channel_duration': [boronmarzi_alalam_duration, boronmarzi_alkosar_duration, boronmarzi_presstv_duration,
                               boronmarzi_ifilm_duration, boronmarzi_sahar_kordi_duration, boronmarzi_sahar_balkan_duration,
                               boronmarzi_jamejam1_duration, boronmarzi_sahar_azari_duration, boronmarzi_sahar_ordu_duration,],}
    boronmarzi_channels_statistics=pd.DataFrame(boronmarzi_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
    boronmarzi_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
    boronmarzi_channels_statistics=boronmarzi_channels_statistics.rename(columns={'channel_name': 'نام شبکه', 'channel_content': 'تعداد محتوا', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})
    
    boronmarzi_alalam_popular_visit = boronmarzi_alalam_popular_visit.reset_index()
    del boronmarzi_alalam_popular_visit['index']
    boronmarzi_alalam_popular_duration = boronmarzi_alalam_popular_duration.reset_index()
    del boronmarzi_alalam_popular_duration['index']
    
    boronmarzi_alkosar_popular_visit = boronmarzi_alkosar_popular_visit.reset_index()
    del boronmarzi_alkosar_popular_visit['index']
    boronmarzi_alkosar_popular_duration = boronmarzi_alkosar_popular_duration.reset_index()
    del boronmarzi_alkosar_popular_duration['index']
    
    boronmarzi_presstv_popular_visit = boronmarzi_presstv_popular_visit.reset_index()
    del boronmarzi_presstv_popular_visit['index']
    boronmarzi_presstv_popular_duration = boronmarzi_presstv_popular_duration.reset_index()
    del boronmarzi_presstv_popular_duration['index']
    
    boronmarzi_ifilm_popular_visit = boronmarzi_ifilm_popular_visit.reset_index()
    del boronmarzi_ifilm_popular_visit['index']
    boronmarzi_ifilm_popular_duration = boronmarzi_ifilm_popular_duration.reset_index()
    del boronmarzi_ifilm_popular_duration['index']
    
    boronmarzi_sahar_kordi_popular_visit = boronmarzi_sahar_kordi_popular_visit.reset_index()
    del boronmarzi_sahar_kordi_popular_visit['index']
    boronmarzi_sahar_kordi_popular_duration = boronmarzi_sahar_kordi_popular_duration.reset_index()
    del boronmarzi_sahar_kordi_popular_duration['index']
    
    boronmarzi_sahar_balkan_popular_visit = boronmarzi_sahar_balkan_popular_visit.reset_index()
    del boronmarzi_sahar_balkan_popular_visit['index']
    boronmarzi_sahar_balkan_popular_duration = boronmarzi_sahar_balkan_popular_duration.reset_index()
    del boronmarzi_sahar_balkan_popular_duration['index']
    
    boronmarzi_jamejam1_popular_visit = boronmarzi_jamejam1_popular_visit.reset_index()
    del boronmarzi_jamejam1_popular_visit['index']
    boronmarzi_jamejam1_popular_duration = boronmarzi_jamejam1_popular_duration.reset_index()
    del boronmarzi_jamejam1_popular_duration['index']
    
    boronmarzi_sahar_azari_popular_visit = boronmarzi_sahar_azari_popular_visit.reset_index()
    del boronmarzi_sahar_azari_popular_visit['index']
    boronmarzi_sahar_azari_popular_duration = boronmarzi_sahar_azari_popular_duration.reset_index()
    del boronmarzi_sahar_azari_popular_duration['index']
    
    boronmarzi_sahar_ordu_popular_visit = boronmarzi_sahar_ordu_popular_visit.reset_index()
    del boronmarzi_sahar_ordu_popular_visit['index']
    boronmarzi_sahar_ordu_popular_duration = boronmarzi_sahar_ordu_popular_duration.reset_index()
    del boronmarzi_sahar_ordu_popular_duration['index']
    
    
    boronmarzi_channels_popular_content=pd.DataFrame()
    boronmarzi_channels_popular_content=pd.concat([boronmarzi_alalam_popular_visit, boronmarzi_alalam_popular_duration,
                                                   boronmarzi_alkosar_popular_visit, boronmarzi_alkosar_popular_duration,
                                                   boronmarzi_presstv_popular_visit, boronmarzi_presstv_popular_duration,
                                                   boronmarzi_ifilm_popular_visit, boronmarzi_ifilm_popular_duration,
                                                   boronmarzi_sahar_kordi_popular_visit, boronmarzi_sahar_kordi_popular_duration,
                                                   boronmarzi_sahar_balkan_popular_visit, boronmarzi_sahar_balkan_popular_duration,
                                                   boronmarzi_jamejam1_popular_visit, boronmarzi_jamejam1_popular_duration,
                                                   boronmarzi_sahar_azari_popular_visit, boronmarzi_sahar_azari_popular_duration,
                                                   boronmarzi_sahar_ordu_popular_visit, boronmarzi_sahar_ordu_popular_duration,],axis=1)
    
    
    writer = pd.ExcelWriter('output/output.sending.hard/آمار برونمرزی.xlsx', engine='xlsxwriter')
    boronmarzi_channels_statistics.to_excel(writer, 'آمار شبکه های برونمرزی', index=False)
    boronmarzi_channels_popular_content.to_excel(writer, 'محتواهای پربازدید', index=False)
    writer.save()
    
    print("END boronmarzi")
    
    
    return boronmarzi_channels_statistics, boronmarzi_channels_popular_content
        
        
        
        
        
        
        
        
        
        
        
        
        
        