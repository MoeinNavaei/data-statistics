    
def sima_data(sima, all_data_Time):
        
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
        
    print("start sima")
    
    sima_all=sima.copy()
    sima_all=sima_all.query("operator != 'سایت شبکه ها'")
#    sima_all=sima_all.query("operator != 'سپهر'")
    sima_all_pivot=sima_all.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    sima_all_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    sima_all_popular_visit=sima_all_pivot.iloc[0:15 , [0, 3]]
    sima_all_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    sima_all_popular_duration=sima_all_pivot.iloc[0:15 , [0, 5]]
    
    print("shabake_1")
    shabake_1=sima_all.query("channel == 'شبکه 1'")
    shabake_1_visit=shabake_1['تعداد بازدید'].sum()
    shabake_1_duration=shabake_1['مدت بازدید'].sum()
    shabake_1_duration=round(shabake_1_duration, 0)
    shabake_1_content=shabake_1.copy()
    shabake_1_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_1_content=len(shabake_1_content)
    shabake_1_pivot=shabake_1.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_1_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_1_popular_visit=shabake_1_pivot.iloc[0:15 , [0, 3]]
    shabake_1_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_1_popular_duration=shabake_1_pivot.iloc[0:15 , [0, 5]]
    
    shabake_1_popular_visit = shabake_1_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 1', 'نام برنامه': 'محتواهای پربازدید شبکه 1'})
    shabake_1_popular_duration = shabake_1_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 1 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 1'})
    
    shabake_1_Time_visit=all_data_Time.query("channel == 'شبکه 1'")
    shabake_1_Time_visit=shabake_1_Time_visit.copy()
    shabake_1_Time_visit=shabake_1_Time_visit.groupby(['ساعت']).sum().reset_index()
    shabake_1_Time_visit.to_excel('aaa.xlsx')
#    del shabake_1_Time_visit['میانگین']
#    del shabake_1_Time_visit['تاریخ']
#    del shabake_1_Time_visit['ردیف']
#    del shabake_1_Time_visit['tag']
    shabake_1_Time_visit = shabake_1_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 1', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 1'})
    
    print("shabake_2")
    shabake_2=sima_all.query("channel == 'شبکه 2'")
    shabake_2_visit=shabake_2['تعداد بازدید'].sum()
    shabake_2_duration=shabake_2['مدت بازدید'].sum()
    shabake_2_duration=round(shabake_2_duration, 0)
    shabake_2_content=shabake_2.copy()
    shabake_2_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_2_content=len(shabake_2_content)
    shabake_2_pivot=shabake_2.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_2_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_2_popular_visit=shabake_2_pivot.iloc[0:15 , [0, 3]]
    shabake_2_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_2_popular_duration=shabake_2_pivot.iloc[0:15 , [0, 5]]
    
    shabake_2_popular_visit = shabake_2_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 2', 'نام برنامه': 'محتواهای پربازدید شبکه 2'})
    shabake_2_popular_duration = shabake_2_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 2 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 2'})
    
    shabake_2_Time_visit=all_data_Time.query("channel == 'شبکه 2'")
    shabake_2_Time_visit=shabake_2_Time_visit.copy()
    shabake_2_Time_visit=shabake_2_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_2_Time_visit['میانگین']
#    del shabake_2_Time_visit['تاریخ']
#    del shabake_2_Time_visit['ردیف']
#    del shabake_2_Time_visit['tag']
    shabake_2_Time_visit = shabake_2_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 2', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 2'})
    
    print("shabake_3")
    shabake_3=sima_all.query("channel == 'شبکه 3'")
    shabake_3_visit=shabake_3['تعداد بازدید'].sum()
    shabake_3_duration=shabake_3['مدت بازدید'].sum()
    shabake_3_duration=round(shabake_3_duration, 0)
    shabake_3_content=shabake_3.copy()
    shabake_3_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_3_content=len(shabake_3_content)
    shabake_3_pivot=shabake_3.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_3_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_3_popular_visit=shabake_3_pivot.iloc[0:15 , [0, 3]]
    shabake_3_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_3_popular_duration=shabake_3_pivot.iloc[0:15 , [0, 5]]
    
    shabake_3_popular_visit = shabake_3_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 3', 'نام برنامه': 'محتواهای پربازدید شبکه 3'})
    shabake_3_popular_duration = shabake_3_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 3 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 3'})
    
    shabake_3_Time_visit=all_data_Time.query("channel == 'شبکه 3'")
    shabake_3_Time_visit=shabake_3_Time_visit.copy()
    shabake_3_Time_visit=shabake_3_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_3_Time_visit['میانگین']
#    del shabake_3_Time_visit['تاریخ']
#    del shabake_3_Time_visit['ردیف']
#    del shabake_3_Time_visit['tag']
    shabake_3_Time_visit = shabake_3_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 3', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 3'})
    
    print("shabake_4")
    shabake_4=sima_all.query("channel == 'شبکه 4'")
    shabake_4_visit=shabake_4['تعداد بازدید'].sum()
    shabake_4_duration=shabake_4['مدت بازدید'].sum()
    shabake_4_duration=round(shabake_4_duration, 0)
    shabake_4_content=shabake_4.copy()
    shabake_4_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_4_content=len(shabake_4_content)
    shabake_4_pivot=shabake_4.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_4_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_4_popular_visit=shabake_4_pivot.iloc[0:15 , [0, 3]]
    shabake_4_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_4_popular_duration=shabake_4_pivot.iloc[0:15 , [0, 5]]
    
    shabake_4_popular_visit = shabake_4_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 4', 'نام برنامه': 'محتواهای پربازدید شبکه 4'})
    shabake_4_popular_duration = shabake_4_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 4 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 4'})
    
    shabake_4_Time_visit=all_data_Time.query("channel == 'شبکه 4'")
    shabake_4_Time_visit=shabake_4_Time_visit.copy()
    shabake_4_Time_visit=shabake_4_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_4_Time_visit['میانگین']
#    del shabake_4_Time_visit['تاریخ']
#    del shabake_4_Time_visit['ردیف']
#    del shabake_4_Time_visit['tag']
    shabake_4_Time_visit = shabake_4_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 4', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 4'})
    
    print("shabake_5")
    shabake_5=sima_all.query("channel == 'شبکه 5'")
    shabake_5_visit=shabake_5['تعداد بازدید'].sum()
    shabake_5_duration=shabake_5['مدت بازدید'].sum()
    shabake_5_duration=round(shabake_5_duration, 0)
    shabake_5_content=shabake_5.copy()
    shabake_5_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_5_content=len(shabake_5_content)
    shabake_5_pivot=shabake_5.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_5_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_5_popular_visit=shabake_5_pivot.iloc[0:15 , [0, 3]]
    shabake_5_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_5_popular_duration=shabake_5_pivot.iloc[0:15 , [0, 5]]
    
    shabake_5_popular_visit = shabake_5_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 5', 'نام برنامه': 'محتواهای پربازدید شبکه 5'})
    shabake_5_popular_duration = shabake_5_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه 5 (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه 5'})
    
    shabake_5_Time_visit=all_data_Time.query("channel == 'شبکه 5'")
    shabake_5_Time_visit=shabake_5_Time_visit.copy()
    shabake_5_Time_visit=shabake_5_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_5_Time_visit['میانگین']
#    del shabake_5_Time_visit['تاریخ']
#    del shabake_5_Time_visit['ردیف']
#    del shabake_5_Time_visit['tag']
    shabake_5_Time_visit = shabake_5_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه 5', 'مدت بازدید': 'مدت بازدید پربازدید شبکه 5'})
    
    print("shabake_khabar")
    shabake_khabar=sima_all.query("channel == 'خبر'")
    shabake_khabar_visit=shabake_khabar['تعداد بازدید'].sum()
    shabake_khabar_duration=shabake_khabar['مدت بازدید'].sum()
    shabake_khabar_duration=round(shabake_khabar_duration, 0)
    shabake_khabar_content=shabake_khabar.copy()
    shabake_khabar_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_khabar_content=len(shabake_khabar_content)
    shabake_khabar_pivot=shabake_khabar.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_khabar_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_khabar_popular_visit=shabake_khabar_pivot.iloc[0:15 , [0, 3]]
    shabake_khabar_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_khabar_popular_duration=shabake_khabar_pivot.iloc[0:15 , [0, 5]]
    
    shabake_khabar_popular_visit = shabake_khabar_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه خبر', 'نام برنامه': 'محتواهای پربازدید شبکه خبر'})
    shabake_khabar_popular_duration = shabake_khabar_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه خبر (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه خبر'})
    
    shabake_khabar_Time_visit=all_data_Time.query("channel == 'خبر'")
    shabake_khabar_Time_visit=shabake_khabar_Time_visit.copy()
    shabake_khabar_Time_visit=shabake_khabar_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_khabar_Time_visit['میانگین']
#    del shabake_khabar_Time_visit['تاریخ']
#    del shabake_khabar_Time_visit['ردیف']
#    del shabake_khabar_Time_visit['tag']
    shabake_khabar_Time_visit = shabake_khabar_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه خبر', 'مدت بازدید': 'مدت بازدید پربازدید شبکه خبر'})
    
    print("shabake_ofogh")
    shabake_ofogh=sima_all.query("channel == 'افق'")
    shabake_ofogh_visit=shabake_ofogh['تعداد بازدید'].sum()
    shabake_ofogh_duration=shabake_ofogh['مدت بازدید'].sum()
    shabake_ofogh_duration=round(shabake_ofogh_duration, 0)
    shabake_ofogh_content=shabake_ofogh.copy()
    shabake_ofogh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_ofogh_content=len(shabake_ofogh_content)
    shabake_ofogh_pivot=shabake_ofogh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_ofogh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_ofogh_popular_visit=shabake_ofogh_pivot.iloc[0:15 , [0, 3]]
    shabake_ofogh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_ofogh_popular_duration=shabake_ofogh_pivot.iloc[0:15 , [0, 5]]
    
    shabake_ofogh_popular_visit = shabake_ofogh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه افق', 'نام برنامه': 'محتواهای پربازدید شبکه افق'})
    shabake_ofogh_popular_duration = shabake_ofogh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه افق (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه افق'})
    
    shabake_ofogh_Time_visit=all_data_Time.query("channel == 'افق'")
    shabake_ofogh_Time_visit=shabake_ofogh_Time_visit.copy()
    shabake_ofogh_Time_visit=shabake_ofogh_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_ofogh_Time_visit['میانگین']
#    del shabake_ofogh_Time_visit['تاریخ']
#    del shabake_ofogh_Time_visit['ردیف']
#    del shabake_ofogh_Time_visit['tag']
    shabake_ofogh_Time_visit = shabake_ofogh_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه افق', 'مدت بازدید': 'مدت بازدید پربازدید شبکه افق'})
    
    print("shabake_pooya")
    shabake_pooya=sima_all.query("channel == 'پویا'")
    shabake_pooya_visit=shabake_pooya['تعداد بازدید'].sum()
    shabake_pooya_duration=shabake_pooya['مدت بازدید'].sum()
    shabake_pooya_duration=round(shabake_pooya_duration, 0)
    shabake_pooya_content=shabake_pooya.copy()
    shabake_pooya_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_pooya_content=len(shabake_pooya_content)
    shabake_pooya_pivot=shabake_pooya.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_pooya_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_pooya_popular_visit=shabake_pooya_pivot.iloc[0:15 , [0, 3]]
    shabake_pooya_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_pooya_popular_duration=shabake_pooya_pivot.iloc[0:15 , [0, 5]]
    
    shabake_pooya_popular_visit = shabake_pooya_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه پویا', 'نام برنامه': 'محتواهای پربازدید شبکه پویا'})
    shabake_pooya_popular_duration = shabake_pooya_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه پویا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه پویا'})
    
    shabake_pooya_Time_visit=all_data_Time.query("channel == 'پویا'")
    shabake_pooya_Time_visit=shabake_pooya_Time_visit.copy()
    shabake_pooya_Time_visit=shabake_pooya_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_pooya_Time_visit['میانگین']
#    del shabake_pooya_Time_visit['تاریخ']
#    del shabake_pooya_Time_visit['ردیف']
#    del shabake_pooya_Time_visit['tag']
    shabake_pooya_Time_visit = shabake_pooya_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه پویا', 'مدت بازدید': 'مدت بازدید پربازدید شبکه پویا'})
    
    print("shabake_omid")
    shabake_omid=sima_all.query("channel == 'امید'")
    shabake_omid_visit=shabake_omid['تعداد بازدید'].sum()
    shabake_omid_duration=shabake_omid['مدت بازدید'].sum()
    shabake_omid_duration=round(shabake_omid_duration, 0)
    shabake_omid_content=shabake_omid.copy()
    shabake_omid_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_omid_content=len(shabake_omid_content)
    shabake_omid_pivot=shabake_omid.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_omid_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_omid_popular_visit=shabake_omid_pivot.iloc[0:15 , [0, 3]]
    shabake_omid_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_omid_popular_duration=shabake_omid_pivot.iloc[0:15 , [0, 5]]
    
    shabake_omid_popular_visit = shabake_omid_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه امید', 'نام برنامه': 'محتواهای پربازدید شبکه امید'})
    shabake_omid_popular_duration = shabake_omid_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه امید (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه امید'})
    
    shabake_omid_Time_visit=all_data_Time.query("channel == 'امید'")
    shabake_omid_Time_visit=shabake_omid_Time_visit.copy()
    shabake_omid_Time_visit=shabake_omid_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_omid_Time_visit['میانگین']
#    del shabake_omid_Time_visit['تاریخ']
#    del shabake_omid_Time_visit['ردیف']
#    del shabake_omid_Time_visit['tag']
    shabake_omid_Time_visit = shabake_omid_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه امید', 'مدت بازدید': 'مدت بازدید پربازدید شبکه امید'})
    
    print("shabake_ifilm")
    shabake_ifilm=sima_all.query("channel == 'آی فیلم'")
    shabake_ifilm_visit=shabake_ifilm['تعداد بازدید'].sum()
    shabake_ifilm_duration=shabake_ifilm['مدت بازدید'].sum()
    shabake_ifilm_duration=round(shabake_ifilm_duration, 0)
    shabake_ifilm_content=shabake_ifilm.copy()
    shabake_ifilm_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_ifilm_content=len(shabake_ifilm_content)
    shabake_ifilm_pivot=shabake_ifilm.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_ifilm_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_ifilm_popular_visit=shabake_ifilm_pivot.iloc[0:15 , [0, 3]]
    shabake_ifilm_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_ifilm_popular_duration=shabake_ifilm_pivot.iloc[0:15 , [0, 5]]
    
    shabake_ifilm_popular_visit = shabake_ifilm_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آی فیلم', 'نام برنامه': 'محتواهای پربازدید شبکه آی فیلم'})
    shabake_ifilm_popular_duration = shabake_ifilm_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه آی فیلم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه آی فیلم'})
    
    shabake_ifilm_Time_visit=all_data_Time.query("channel == 'آی فیلم'")
    shabake_ifilm_Time_visit=shabake_ifilm_Time_visit.copy()
    shabake_ifilm_Time_visit=shabake_ifilm_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_ifilm_Time_visit['میانگین']
#    del shabake_ifilm_Time_visit['تاریخ']
#    del shabake_ifilm_Time_visit['ردیف']
#    del shabake_ifilm_Time_visit['tag']
    shabake_ifilm_Time_visit = shabake_ifilm_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آی فیلم', 'مدت بازدید': 'مدت بازدید پربازدید شبکه آی فیلم'})
    
    print("shabake_namayesh")
    shabake_namayesh=sima_all.query("channel == 'نمایش'")
    shabake_namayesh_visit=shabake_namayesh['تعداد بازدید'].sum()
    shabake_namayesh_duration=shabake_namayesh['مدت بازدید'].sum()
    shabake_namayesh_duration=round(shabake_namayesh_duration, 0)
    shabake_namayesh_content=shabake_namayesh.copy()
    shabake_namayesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_namayesh_content=len(shabake_namayesh_content)
    shabake_namayesh_pivot=shabake_namayesh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_namayesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_namayesh_popular_visit=shabake_namayesh_pivot.iloc[0:15 , [0, 3]]
    shabake_namayesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_namayesh_popular_duration=shabake_namayesh_pivot.iloc[0:15 , [0, 5]]
    
    shabake_namayesh_popular_visit = shabake_namayesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه نمایش', 'نام برنامه': 'محتواهای پربازدید شبکه نمایش'})
    shabake_namayesh_popular_duration = shabake_namayesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه نمایش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه نمایش'})
    
    shabake_namayesh_Time_visit=all_data_Time.query("channel == 'نمایش'")
    shabake_namayesh_Time_visit=shabake_namayesh_Time_visit.copy()
    shabake_namayesh_Time_visit=shabake_namayesh_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_namayesh_Time_visit['میانگین']
#    del shabake_namayesh_Time_visit['تاریخ']
#    del shabake_namayesh_Time_visit['ردیف']
#    del shabake_namayesh_Time_visit['tag']
    shabake_namayesh_Time_visit = shabake_namayesh_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه نمایش', 'مدت بازدید': 'مدت بازدید پربازدید شبکه نمایش'})
    
    print("shabake_tamasha")
    shabake_tamasha=sima_all.query("channel == 'تماشا'")
    shabake_tamasha_visit=shabake_tamasha['تعداد بازدید'].sum()
    shabake_tamasha_duration=shabake_tamasha['مدت بازدید'].sum()
    shabake_tamasha_duration=round(shabake_tamasha_duration, 0)
    shabake_tamasha_content=shabake_tamasha.copy()
    shabake_tamasha_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_tamasha_content=len(shabake_tamasha_content)
    shabake_tamasha_pivot=shabake_tamasha.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_tamasha_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_tamasha_popular_visit=shabake_tamasha_pivot.iloc[0:15 , [0, 3]]
    shabake_tamasha_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_tamasha_popular_duration=shabake_tamasha_pivot.iloc[0:15 , [0, 5]]
    
    shabake_tamasha_popular_visit = shabake_tamasha_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه تماشا', 'نام برنامه': 'محتواهای پربازدید شبکه تماشا'})
    shabake_tamasha_popular_duration = shabake_tamasha_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه تماشا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه تماشا'})
    
    shabake_tamasha_Time_visit=all_data_Time.query("channel == 'تماشا'")
    shabake_tamasha_Time_visit=shabake_tamasha_Time_visit.copy()
    shabake_tamasha_Time_visit=shabake_tamasha_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_tamasha_Time_visit['میانگین']
#    del shabake_tamasha_Time_visit['تاریخ']
#    del shabake_tamasha_Time_visit['ردیف']
#    del shabake_tamasha_Time_visit['tag']
    shabake_tamasha_Time_visit = shabake_tamasha_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه تماشا', 'مدت بازدید': 'مدت بازدید پربازدید شبکه تماشا'})
    
    print("shabake_mostanad")
    shabake_mostanad=sima_all.query("channel == 'مستند'")
    shabake_mostanad_visit=shabake_mostanad['تعداد بازدید'].sum()
    shabake_mostanad_duration=shabake_mostanad['مدت بازدید'].sum()
    shabake_mostanad_duration=round(shabake_mostanad_duration, 0)
    shabake_mostanad_content=shabake_mostanad.copy()
    shabake_mostanad_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_mostanad_content=len(shabake_mostanad_content)
    shabake_mostanad_pivot=shabake_mostanad.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_mostanad_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_mostanad_popular_visit=shabake_mostanad_pivot.iloc[0:15 , [0, 3]]
    shabake_mostanad_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_mostanad_popular_duration=shabake_mostanad_pivot.iloc[0:15 , [0, 5]]
    
    shabake_mostanad_popular_visit = shabake_mostanad_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه مستند', 'نام برنامه': 'محتواهای پربازدید شبکه مستند'})
    shabake_mostanad_popular_duration = shabake_mostanad_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه مستند (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه مستند'})
    
    shabake_mostanad_Time_visit=all_data_Time.query("channel == 'مستند'")
    shabake_mostanad_Time_visit=shabake_mostanad_Time_visit.copy()
    shabake_mostanad_Time_visit=shabake_mostanad_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_mostanad_Time_visit['میانگین']
#    del shabake_mostanad_Time_visit['تاریخ']
#    del shabake_mostanad_Time_visit['ردیف']
#    del shabake_mostanad_Time_visit['tag']
    shabake_mostanad_Time_visit = shabake_mostanad_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه مستند', 'مدت بازدید': 'مدت بازدید پربازدید شبکه مستند'})
    
    print("shabake_shoma")
    shabake_shoma=sima_all.query("channel == 'شما'")
    shabake_shoma_visit=shabake_shoma['تعداد بازدید'].sum()
    shabake_shoma_duration=shabake_shoma['مدت بازدید'].sum()
    shabake_shoma_duration=round(shabake_shoma_duration, 0)
    shabake_shoma_content=shabake_shoma.copy()
    shabake_shoma_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_shoma_content=len(shabake_shoma_content)
    shabake_shoma_pivot=shabake_shoma.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_shoma_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_shoma_popular_visit=shabake_shoma_pivot.iloc[0:15 , [0, 3]]
    shabake_shoma_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_shoma_popular_duration=shabake_shoma_pivot.iloc[0:15 , [0, 5]]
    
    shabake_shoma_popular_visit = shabake_shoma_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه شما', 'نام برنامه': 'محتواهای پربازدید شبکه شما'})
    shabake_shoma_popular_duration = shabake_shoma_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه شما (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه شما'})
    
    shabake_shoma_Time_visit=all_data_Time.query("channel == 'شما'")
    shabake_shoma_Time_visit=shabake_shoma_Time_visit.copy()
    shabake_shoma_Time_visit=shabake_shoma_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_shoma_Time_visit['میانگین']
#    del shabake_shoma_Time_visit['تاریخ']
#    del shabake_shoma_Time_visit['ردیف']
#    del shabake_shoma_Time_visit['tag']
    shabake_shoma_Time_visit = shabake_shoma_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه شما', 'مدت بازدید': 'مدت بازدید پربازدید شبکه شما'})
    
    print("shabake_amozesh")
    shabake_amozesh=sima_all.query("channel == 'آموزش'")
    shabake_amozesh_visit=shabake_amozesh['تعداد بازدید'].sum()
    shabake_amozesh_duration=shabake_amozesh['مدت بازدید'].sum()
    shabake_amozesh_duration=round(shabake_amozesh_duration, 0)
    shabake_amozesh_content=shabake_amozesh.copy()
    shabake_amozesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_amozesh_content=len(shabake_amozesh_content)
    shabake_amozesh_pivot=shabake_amozesh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_amozesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_amozesh_popular_visit=shabake_amozesh_pivot.iloc[0:15 , [0, 3]]
    shabake_amozesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_amozesh_popular_duration=shabake_amozesh_pivot.iloc[0:15 , [0, 5]]
    
    shabake_amozesh_popular_visit = shabake_amozesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آموزش', 'نام برنامه': 'محتواهای پربازدید شبکه آموزش'})
    shabake_amozesh_popular_duration = shabake_amozesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه آموزش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه آموزش'})
    
    shabake_amozesh_Time_visit=all_data_Time.query("channel == 'آموزش'")
    shabake_amozesh_Time_visit=shabake_amozesh_Time_visit.copy()
    shabake_amozesh_Time_visit=shabake_amozesh_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_amozesh_Time_visit['میانگین']
#    del shabake_amozesh_Time_visit['تاریخ']
#    del shabake_amozesh_Time_visit['ردیف']
#    del shabake_amozesh_Time_visit['tag']
    shabake_amozesh_Time_visit = shabake_amozesh_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه آموزش', 'مدت بازدید': 'مدت بازدید پربازدید شبکه آموزش'})
    
    print("shabake_varzesh")
    shabake_varzesh=sima_all.query("channel == 'ورزش'")
    shabake_varzesh_visit=shabake_varzesh['تعداد بازدید'].sum()
    shabake_varzesh_duration=shabake_varzesh['مدت بازدید'].sum()
    shabake_varzesh_duration=round(shabake_varzesh_duration, 0)
    shabake_varzesh_content=shabake_varzesh.copy()
    shabake_varzesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_varzesh_content=len(shabake_varzesh_content)
    shabake_varzesh_pivot=shabake_varzesh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_varzesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_varzesh_popular_visit=shabake_varzesh_pivot.iloc[0:15 , [0, 3]]
    shabake_varzesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_varzesh_popular_duration=shabake_varzesh_pivot.iloc[0:15 , [0, 5]]
    
    shabake_varzesh_popular_visit = shabake_varzesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه ورزش', 'نام برنامه': 'محتواهای پربازدید شبکه ورزش'})
    shabake_varzesh_popular_duration = shabake_varzesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه ورزش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه ورزش'})
    
    shabake_varzesh_Time_visit=all_data_Time.query("channel == 'ورزش'")
    shabake_varzesh_Time_visit=shabake_varzesh_Time_visit.copy()
    shabake_varzesh_Time_visit=shabake_varzesh_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_varzesh_Time_visit['میانگین']
#    del shabake_varzesh_Time_visit['تاریخ']
#    del shabake_varzesh_Time_visit['ردیف']
#    del shabake_varzesh_Time_visit['tag']
    shabake_varzesh_Time_visit = shabake_varzesh_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه ورزش', 'مدت بازدید': 'مدت بازدید پربازدید شبکه ورزش'})
    
    print("shabake_nasim")
    shabake_nasim=sima_all.query("channel == 'نسیم'")
    shabake_nasim_visit=shabake_nasim['تعداد بازدید'].sum()
    shabake_nasim_duration=shabake_nasim['مدت بازدید'].sum()
    shabake_nasim_duration=round(shabake_nasim_duration, 0)
    shabake_nasim_content=shabake_nasim.copy()
    shabake_nasim_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_nasim_content=len(shabake_nasim_content)
    shabake_nasim_pivot=shabake_nasim.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_nasim_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_nasim_popular_visit=shabake_nasim_pivot.iloc[0:15 , [0, 3]]
    shabake_nasim_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_nasim_popular_duration=shabake_nasim_pivot.iloc[0:15 , [0, 5]]
    
    shabake_nasim_popular_visit = shabake_nasim_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه نسیم', 'نام برنامه': 'محتواهای پربازدید شبکه نسیم'})
    shabake_nasim_popular_duration = shabake_nasim_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه نسیم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه نسیم'})
    
    shabake_nasim_Time_visit=all_data_Time.query("channel == 'نسیم'")
    shabake_nasim_Time_visit=shabake_nasim_Time_visit.copy()
    shabake_nasim_Time_visit=shabake_nasim_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_nasim_Time_visit['میانگین']
#    del shabake_nasim_Time_visit['تاریخ']
#    del shabake_nasim_Time_visit['ردیف']
#    del shabake_nasim_Time_visit['tag']
    shabake_nasim_Time_visit = shabake_nasim_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه نسیم', 'مدت بازدید': 'مدت بازدید پربازدید شبکه نسیم'})
    
    print("shabake_qoran")
    shabake_qoran=sima_all.query("channel == 'قرآن'")
    shabake_qoran_visit=shabake_qoran['تعداد بازدید'].sum()
    shabake_qoran_duration=shabake_qoran['مدت بازدید'].sum()
    shabake_qoran_duration=round(shabake_qoran_duration, 0)
    shabake_qoran_content=shabake_qoran.copy()
    shabake_qoran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_qoran_content=len(shabake_qoran_content)
    shabake_qoran_pivot=shabake_qoran.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_qoran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_qoran_popular_visit=shabake_qoran_pivot.iloc[0:15 , [0, 3]]
    shabake_qoran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_qoran_popular_duration=shabake_qoran_pivot.iloc[0:15 , [0, 5]]
    
    shabake_qoran_popular_visit = shabake_qoran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه قرآن', 'نام برنامه': 'محتواهای پربازدید شبکه قرآن'})
    shabake_qoran_popular_duration = shabake_qoran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه قرآن (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه قرآن'})
    
    shabake_qoran_Time_visit=all_data_Time.query("channel == 'قرآن'")
    shabake_qoran_Time_visit=shabake_qoran_Time_visit.copy()
    shabake_qoran_Time_visit=shabake_qoran_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_qoran_Time_visit['میانگین']
#    del shabake_qoran_Time_visit['تاریخ']
#    del shabake_qoran_Time_visit['ردیف']
#    del shabake_qoran_Time_visit['tag']
    shabake_qoran_Time_visit = shabake_qoran_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه قرآن', 'مدت بازدید': 'مدت بازدید پربازدید شبکه قرآن'})
    
    print("shabake_salamat")
    shabake_salamat=sima_all.query("channel == 'سلامت'")
    shabake_salamat_visit=shabake_salamat['تعداد بازدید'].sum()
    shabake_salamat_duration=shabake_salamat['مدت بازدید'].sum()
    shabake_salamat_duration=round(shabake_salamat_duration, 0)
    shabake_salamat_content=shabake_salamat.copy()
    shabake_salamat_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_salamat_content=len(shabake_salamat_content)
    shabake_salamat_pivot=shabake_salamat.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_salamat_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_salamat_popular_visit=shabake_salamat_pivot.iloc[0:15 , [0, 3]]
    shabake_salamat_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_salamat_popular_duration=shabake_salamat_pivot.iloc[0:15 , [0, 5]]
    
    shabake_salamat_popular_visit = shabake_salamat_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سلامت', 'نام برنامه': 'محتواهای پربازدید شبکه سلامت'})
    shabake_salamat_popular_duration = shabake_salamat_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه سلامت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه سلامت'})
    
    shabake_salamat_Time_visit=all_data_Time.query("channel == 'سلامت'")
    shabake_salamat_Time_visit=shabake_salamat_Time_visit.copy()
    shabake_salamat_Time_visit=shabake_salamat_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_salamat_Time_visit['میانگین']
#    del shabake_salamat_Time_visit['تاریخ']
#    del shabake_salamat_Time_visit['ردیف']
#    del shabake_salamat_Time_visit['tag']
    shabake_salamat_Time_visit = shabake_salamat_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سلامت', 'مدت بازدید': 'مدت بازدید پربازدید شبکه سلامت'})
    
    print("shabake_irankala")
    shabake_irankala=sima_all.query("channel == 'ایران کالا'")
    shabake_irankala_visit=shabake_irankala['تعداد بازدید'].sum()
    shabake_irankala_duration=shabake_irankala['مدت بازدید'].sum()
    shabake_irankala_duration=round(shabake_irankala_duration, 0)
    shabake_irankala_content=shabake_irankala.copy()
    shabake_irankala_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_irankala_content=len(shabake_irankala_content)
    shabake_irankala_pivot=shabake_irankala.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_irankala_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_irankala_popular_visit=shabake_irankala_pivot.iloc[0:15 , [0, 3]]
    shabake_irankala_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_irankala_popular_duration=shabake_irankala_pivot.iloc[0:15 , [0, 5]]
    
    shabake_irankala_popular_visit = shabake_irankala_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه ایران کالا', 'نام برنامه': 'محتواهای پربازدید شبکه ایران کالا'})
    shabake_irankala_popular_duration = shabake_irankala_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه ایران کالا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه ایران کالا'})
    
    shabake_irankala_Time_visit=all_data_Time.query("channel == 'ایران کالا'")
    shabake_irankala_Time_visit=shabake_irankala_Time_visit.copy()
    shabake_irankala_Time_visit=shabake_irankala_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_irankala_Time_visit['میانگین']
#    del shabake_irankala_Time_visit['تاریخ']
#    del shabake_irankala_Time_visit['ردیف']
#    del shabake_irankala_Time_visit['tag']
    shabake_irankala_Time_visit = shabake_irankala_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه ایران کالا', 'مدت بازدید': 'مدت بازدید پربازدید شبکه ایران کالا'})
    
    print("shabake_sepehr")
    shabake_sepehr=sima_all.query("channel == 'سپهر'")
    shabake_sepehr_visit=shabake_sepehr['تعداد بازدید'].sum()
    shabake_sepehr_duration=shabake_sepehr['مدت بازدید'].sum()
    shabake_sepehr_duration=round(shabake_sepehr_duration, 0)
    shabake_sepehr_content=shabake_sepehr.copy()
    shabake_sepehr_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    shabake_sepehr_content=len(shabake_sepehr_content)
    shabake_sepehr_pivot=shabake_sepehr.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    shabake_sepehr_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_sepehr_popular_visit=shabake_sepehr_pivot.iloc[0:15 , [0, 3]]
    shabake_sepehr_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    shabake_sepehr_popular_duration=shabake_sepehr_pivot.iloc[0:15 , [0, 5]]
    
    shabake_sepehr_popular_visit = shabake_sepehr_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سپهر', 'نام برنامه': 'محتواهای پربازدید شبکه سپهر'})
    shabake_sepehr_popular_duration = shabake_sepehr_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید شبکه سپهر (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید شبکه سپهر'})
    
    shabake_sepehr_Time_visit=all_data_Time.query("channel == 'سپهر'")
    shabake_sepehr_Time_visit=shabake_sepehr_Time_visit.copy()
    shabake_sepehr_Time_visit=shabake_sepehr_Time_visit.groupby(['ساعت']).sum().reset_index()
#    del shabake_sepehr_Time_visit['میانگین']
#    del shabake_sepehr_Time_visit['تاریخ']
#    del shabake_sepehr_Time_visit['ردیف']
#    del shabake_sepehr_Time_visit['tag']
    shabake_sepehr_Time_visit = shabake_sepehr_Time_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید شبکه سپهر', 'مدت بازدید': 'مدت بازدید پربازدید شبکه سپهر'})
    
    
    print("dataframe sima channels")
    sima_channels_statistics={'channel_name': ['شبکه 1', 'شبکه 2', 'شبکه 3', 'شبکه 4', 'شبکه 5',
                                         'شبکه خبر', 'شبکه افق', 'شبکه پویا', 'شبکه امید', 'شبکه آی فیلم',
                                         'شبکه نمایش', 'شبکه تماشا', 'شبکه مستند', 'شبکه شما', 'شبکه آموزش',
                                         'شبکه ورزش', 'شبکه نسیم', 'شبکه قرآن', 'شبکه سلامت', 'شبکه ایران کالا',
                                         'شبکه سپهر',],
           'channel_content': [shabake_1_content, shabake_2_content, shabake_3_content, shabake_5_content, shabake_5_content,
                               shabake_khabar_content, shabake_ofogh_content, shabake_pooya_content, shabake_omid_content, shabake_ifilm_content,
                               shabake_namayesh_content, shabake_tamasha_content, shabake_mostanad_content, shabake_shoma_content, shabake_amozesh_content,
                               shabake_varzesh_content, shabake_nasim_content, shabake_qoran_content, shabake_salamat_content, shabake_irankala_content,
                               shabake_sepehr_content,],
           'channel_visit': [shabake_1_visit, shabake_2_visit, shabake_3_visit, shabake_5_visit, shabake_5_visit,
                               shabake_khabar_visit, shabake_ofogh_visit, shabake_pooya_visit, shabake_omid_visit, shabake_ifilm_visit,
                               shabake_namayesh_visit, shabake_tamasha_visit, shabake_mostanad_visit, shabake_shoma_visit, shabake_amozesh_visit,
                               shabake_varzesh_visit, shabake_nasim_visit, shabake_qoran_visit, shabake_salamat_visit, shabake_irankala_visit,
                               shabake_sepehr_visit,],
           'channel_duration': [shabake_1_duration, shabake_2_duration, shabake_3_duration, shabake_5_duration, shabake_5_duration,
                               shabake_khabar_duration, shabake_ofogh_duration, shabake_pooya_duration, shabake_omid_duration, shabake_ifilm_duration,
                               shabake_namayesh_duration, shabake_tamasha_duration, shabake_mostanad_duration, shabake_shoma_duration, shabake_amozesh_duration,
                               shabake_varzesh_duration, shabake_nasim_duration, shabake_qoran_duration, shabake_salamat_duration, shabake_irankala_duration,
                               shabake_sepehr_duration,],}
    sima_channels_statistics=pd.DataFrame(sima_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
    sima_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
    sima_channels_statistics=sima_channels_statistics.rename(columns={'channel_name': 'نام شبکه', 'channel_content': 'تعداد محتوا', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})
    
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
    
    shabake_sepehr_popular_visit.to_excel('busy/shabake_sepehr_popular_visit.xlsx')
    shabake_sepehr_popular_duration.to_excel('busy/shabake_sepehr_popular_duration.xlsx')
    shabake_sepehr_popular_visit=pd.read_excel('busy/shabake_sepehr_popular_visit.xlsx')
    shabake_sepehr_popular_duration=pd.read_excel('busy/shabake_sepehr_popular_duration.xlsx')
    del shabake_sepehr_popular_visit['Unnamed: 0']
    del shabake_sepehr_popular_duration['Unnamed: 0']
    
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
                   shabake_sepehr_popular_visit, shabake_sepehr_popular_duration,],axis=1)
     
    sima_channels_Time_visit=pd.DataFrame()
    sima_channels_Time_visit=pd.concat([shabake_1_Time_visit, shabake_2_Time_visit, shabake_3_Time_visit,shabake_4_Time_visit, shabake_5_Time_visit, 
                   shabake_khabar_Time_visit, shabake_ofogh_Time_visit, shabake_pooya_Time_visit,shabake_omid_Time_visit, shabake_ifilm_Time_visit,
                   shabake_namayesh_Time_visit, shabake_tamasha_Time_visit, shabake_mostanad_Time_visit,shabake_shoma_Time_visit, shabake_amozesh_Time_visit,
                   shabake_varzesh_Time_visit, shabake_nasim_Time_visit, shabake_qoran_Time_visit,shabake_salamat_Time_visit, shabake_irankala_Time_visit,
                   shabake_sepehr_Time_visit,],axis=1)
    
#    writer = pd.ExcelWriter('output/آمار ماه جاری/آمار سیما.xlsx', engine='xlsxwriter')
#    sima_channels_statistics.to_excel(writer, 'آمار شبکه های سیما')
#    sima_channels_popular_content.to_excel(writer, 'محتواهای پربازدید')
#    sima_channels_Time_visit.to_excel(writer, 'آمار ساعتی')
#    writer.save()
    
#    writer = pd.ExcelWriter('output/moh.rast/آمار سیما.xlsx', engine='xlsxwriter')
#    sima_channels_statistics.to_excel(writer, 'آمار شبکه های سیما')
#    sima_channels_popular_content.to_excel(writer, 'محتواهای پربازدید')
#    sima_channels_Time_visit.to_excel(writer, 'آمار ساعتی')
#    writer.save()
    
    writer = pd.ExcelWriter('output/zomorrodi/آمار سیما.xlsx', engine='xlsxwriter')
    sima_channels_statistics.to_excel(writer, 'آمار شبکه های سیما', index=False)
    sima_channels_popular_content.to_excel(writer, 'محتواهای پربازدید', index=False)
    sima_channels_Time_visit.to_excel(writer, 'آمار ساعتی', index=False)
    writer.save()
    
    writer = pd.ExcelWriter('output/output.sending.hard/آمار سیما.xlsx', engine='xlsxwriter')
    sima_channels_statistics.to_excel(writer, 'آمار شبکه های سیما', index=False)
    sima_channels_popular_content.to_excel(writer, 'محتواهای پربازدید', index=False)
    sima_channels_Time_visit.to_excel(writer, 'آمار ساعتی', index=False)
    writer.save()
    
    print("End sima")
    
    
    return sima_channels_statistics, sima_channels_popular_content, sima_channels_Time_visit
        
       
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
