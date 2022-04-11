    
def ostani_data(ostani):
            
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
        
        
    print("start ostani")
    
    ostani_all=ostani.copy()
    ostani_all_pivot=ostani_all.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_all_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_all_popular_visit=ostani_all_pivot.iloc[0:10 , [0, 5]]
    ostani_all_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_all_popular_duration=ostani_all_pivot.iloc[0:10 , [0, 5]]
    
    
    print("ostani_abadan")
    ostani_abadan=ostani.query("channel == 'استانی آبادان'")
    ostani_abadan_visit=ostani_abadan['تعداد بازدید'].sum()
    ostani_abadan_duration=ostani_abadan['مدت بازدید'].sum()
    ostani_abadan_duration=round(ostani_abadan_duration, 0)
    ostani_abadan_content=ostani_abadan.copy()
    ostani_abadan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_abadan_content=len(ostani_abadan_content)
    ostani_abadan_pivot=ostani_abadan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_abadan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_abadan_popular_visit=ostani_abadan_pivot.iloc[0:10 , [0, 3]]
    ostani_abadan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_abadan_popular_duration=ostani_abadan_pivot.iloc[0:10 , [0, 5]]
    
    ostani_abadan_popular_visit = ostani_abadan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی آبادان', 'نام برنامه': 'محتواهای پربازدید استانی آبادان'})
    ostani_abadan_popular_duration = ostani_abadan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی آبادان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی آبادان'})
    
    print("ostani_azarbayjan_gharbi")
    ostani_azarbayjan_gharbi=ostani.query("channel == 'استانی آذربایجان غربی'")
    ostani_azarbayjan_gharbi_visit=ostani_azarbayjan_gharbi['تعداد بازدید'].sum()
    ostani_azarbayjan_gharbi_duration=ostani_azarbayjan_gharbi['مدت بازدید'].sum()
    ostani_azarbayjan_gharbi_duration=round(ostani_azarbayjan_gharbi_duration, 0)
    ostani_azarbayjan_gharbi_content=ostani_azarbayjan_gharbi.copy()
    ostani_azarbayjan_gharbi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_azarbayjan_gharbi_content=len(ostani_azarbayjan_gharbi_content)
    ostani_azarbayjan_gharbi_pivot=ostani_azarbayjan_gharbi.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_azarbayjan_gharbi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_azarbayjan_gharbi_popular_visit=ostani_azarbayjan_gharbi_pivot.iloc[0:10 , [0, 3]]
    ostani_azarbayjan_gharbi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_azarbayjan_gharbi_popular_duration=ostani_azarbayjan_gharbi_pivot.iloc[0:10 , [0, 5]]
    
    ostani_azarbayjan_gharbi_popular_visit = ostani_azarbayjan_gharbi_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی آذربایجان غربی', 'نام برنامه': 'محتواهای پربازدید استانی آذربایجان غربی'})
    ostani_azarbayjan_gharbi_popular_duration = ostani_azarbayjan_gharbi_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی آذربایجان غربی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی آذربایجان غربی'})
    
    print("ostani_esfahan")
    ostani_esfahan=ostani.query("channel == 'استانی اصفهان'")
    ostani_esfahan_visit=ostani_esfahan['تعداد بازدید'].sum()
    ostani_esfahan_duration=ostani_esfahan['مدت بازدید'].sum()
    ostani_esfahan_duration=round(ostani_esfahan_duration, 0)
    ostani_esfahan_content=ostani_esfahan.copy()
    ostani_esfahan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_esfahan_content=len(ostani_esfahan_content)
    ostani_esfahan_pivot=ostani_esfahan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_esfahan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_esfahan_popular_visit=ostani_esfahan_pivot.iloc[0:10 , [0, 3]]
    ostani_esfahan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_esfahan_popular_duration=ostani_esfahan_pivot.iloc[0:10 , [0, 5]]
    
    ostani_esfahan_popular_visit = ostani_esfahan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی اصفهان', 'نام برنامه': 'محتواهای پربازدید استانی اصفهان'})
    ostani_esfahan_popular_duration = ostani_esfahan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی اصفهان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی اصفهان'})
    
    print("ostani_aflak")
    ostani_aflak=ostani.query("channel == 'استانی افلاک'")
    ostani_aflak_visit=ostani_aflak['تعداد بازدید'].sum()
    ostani_aflak_duration=ostani_aflak['مدت بازدید'].sum()
    ostani_aflak_duration=round(ostani_aflak_duration, 0)
    ostani_aflak_content=ostani_aflak.copy()
    ostani_aflak_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_aflak_content=len(ostani_aflak_content)
    ostani_aflak_pivot=ostani_aflak.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_aflak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_aflak_popular_visit=ostani_aflak_pivot.iloc[0:10 , [0, 3]]
    ostani_aflak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_aflak_popular_duration=ostani_aflak_pivot.iloc[0:10 , [0, 5]]
    
    ostani_aflak_popular_visit = ostani_aflak_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی افلاک', 'نام برنامه': 'محتواهای پربازدید استانی افلاک'})
    ostani_aflak_popular_duration = ostani_aflak_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی افلاک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی افلاک'})
    
    print("ostani_alborz")
    ostani_alborz=ostani.query("channel == 'استانی البرز'")
    ostani_alborz_visit=ostani_alborz['تعداد بازدید'].sum()
    ostani_alborz_duration=ostani_alborz['مدت بازدید'].sum()
    ostani_alborz_duration=round(ostani_alborz_duration, 0)
    ostani_alborz_content=ostani_alborz.copy()
    ostani_alborz_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_alborz_content=len(ostani_alborz_content)
    ostani_alborz_pivot=ostani_alborz.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_alborz_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_alborz_popular_visit=ostani_alborz_pivot.iloc[0:10 , [0, 3]]
    ostani_alborz_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_alborz_popular_duration=ostani_alborz_pivot.iloc[0:10 , [0, 5]]
    
    ostani_alborz_popular_visit = ostani_alborz_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی البرز', 'نام برنامه': 'محتواهای پربازدید استانی البرز'})
    ostani_alborz_popular_duration = ostani_alborz_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی البرز (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی البرز'})
    
    print("ostani_ilam")
    ostani_ilam=ostani.query("channel == 'استانی ایلام'")
    ostani_ilam_visit=ostani_ilam['تعداد بازدید'].sum()
    ostani_ilam_duration=ostani_ilam['مدت بازدید'].sum()
    ostani_ilam_duration=round(ostani_ilam_duration, 0)
    ostani_ilam_content=ostani_ilam.copy()
    ostani_ilam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_ilam_content=len(ostani_ilam_content)
    ostani_ilam_pivot=ostani_ilam.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_ilam_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_ilam_popular_visit=ostani_ilam_pivot.iloc[0:10 , [0, 3]]
    ostani_ilam_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_ilam_popular_duration=ostani_ilam_pivot.iloc[0:10 , [0, 5]]
    
    ostani_ilam_popular_visit = ostani_ilam_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی ایلام', 'نام برنامه': 'محتواهای پربازدید استانی ایلام'})
    ostani_ilam_popular_duration = ostani_ilam_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی ایلام (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی ایلام'})
    
    print("ostani_baran")
    ostani_baran=ostani.query("channel == 'استانی باران'")
    ostani_baran_visit=ostani_baran['تعداد بازدید'].sum()
    ostani_baran_duration=ostani_baran['مدت بازدید'].sum()
    ostani_baran_duration=round(ostani_baran_duration, 0)
    ostani_baran_content=ostani_baran.copy()
    ostani_baran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_baran_content=len(ostani_baran_content)
    ostani_baran_pivot=ostani_baran.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_baran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_baran_popular_visit=ostani_baran_pivot.iloc[0:10 , [0, 3]]
    ostani_baran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_baran_popular_duration=ostani_baran_pivot.iloc[0:10 , [0, 5]]
    
    ostani_baran_popular_visit = ostani_baran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی باران', 'نام برنامه': 'محتواهای پربازدید استانی باران'})
    ostani_baran_popular_duration = ostani_baran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی باران (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی باران'})
    
    print("ostani_boshehr")
    ostani_boshehr=ostani.query("channel == 'استانی بوشهر'")
    ostani_boshehr_visit=ostani_boshehr['تعداد بازدید'].sum()
    ostani_boshehr_duration=ostani_boshehr['مدت بازدید'].sum()
    ostani_boshehr_duration=round(ostani_boshehr_duration, 0)
    ostani_boshehr_content=ostani_boshehr.copy()
    ostani_boshehr_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_boshehr_content=len(ostani_boshehr_content)
    ostani_boshehr_pivot=ostani_boshehr.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_boshehr_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_boshehr_popular_visit=ostani_boshehr_pivot.iloc[0:10 , [0, 3]]
    ostani_boshehr_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_boshehr_popular_duration=ostani_boshehr_pivot.iloc[0:10 , [0, 5]]
    
    ostani_boshehr_popular_visit = ostani_boshehr_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی بوشهر', 'نام برنامه': 'محتواهای پربازدید استانی بوشهر'})
    ostani_boshehr_popular_duration = ostani_boshehr_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی بوشهر (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی بوشهر'})
    
    print("ostani_taban")
    ostani_taban=ostani.query("channel == 'استانی تابان'")
    ostani_taban_visit=ostani_taban['تعداد بازدید'].sum()
    ostani_taban_duration=ostani_taban['مدت بازدید'].sum()
    ostani_taban_duration=round(ostani_taban_duration, 0)
    ostani_taban_content=ostani_taban.copy()
    ostani_taban_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_taban_content=len(ostani_taban_content)
    ostani_taban_pivot=ostani_taban.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_taban_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_taban_popular_visit=ostani_taban_pivot.iloc[0:10 , [0, 3]]
    ostani_taban_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_taban_popular_duration=ostani_taban_pivot.iloc[0:10 , [0, 5]]
    
    ostani_taban_popular_visit = ostani_taban_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی تابان', 'نام برنامه': 'محتواهای پربازدید استانی تابان'})
    ostani_taban_popular_duration = ostani_taban_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی تابان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی تابان'})
    
    print("ostani_khorasan_razavi")
    ostani_khorasan_razavi=ostani.query("channel == 'استانی خراسان رضوی'")
    ostani_khorasan_razavi_visit=ostani_khorasan_razavi['تعداد بازدید'].sum()
    ostani_khorasan_razavi_duration=ostani_khorasan_razavi['مدت بازدید'].sum()
    ostani_khorasan_razavi_duration=round(ostani_khorasan_razavi_duration, 0)
    ostani_khorasan_razavi_content=ostani_khorasan_razavi.copy()
    ostani_khorasan_razavi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_khorasan_razavi_content=len(ostani_khorasan_razavi_content)
    ostani_khorasan_razavi_pivot=ostani_khorasan_razavi.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_khorasan_razavi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_khorasan_razavi_popular_visit=ostani_khorasan_razavi_pivot.iloc[0:10 , [0, 3]]
    ostani_khorasan_razavi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_khorasan_razavi_popular_duration=ostani_khorasan_razavi_pivot.iloc[0:10 , [0, 5]]
    
    ostani_khorasan_razavi_popular_visit = ostani_khorasan_razavi_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی خراسان رضوی', 'نام برنامه': 'محتواهای پربازدید استانی خراسان رضوی'})
    ostani_khorasan_razavi_popular_duration = ostani_khorasan_razavi_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی خراسان رضوی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی خراسان رضوی'})
    
    print("ostani_khozestan")
    ostani_khozestan=ostani.query("channel == 'استانی خوزستان'")
    ostani_khozestan_visit=ostani_khozestan['تعداد بازدید'].sum()
    ostani_khozestan_duration=ostani_khozestan['مدت بازدید'].sum()
    ostani_khozestan_duration=round(ostani_khozestan_duration, 0)
    ostani_khozestan_content=ostani_khozestan.copy()
    ostani_khozestan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_khozestan_content=len(ostani_khozestan_content)
    ostani_khozestan_pivot=ostani_khozestan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_khozestan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_khozestan_popular_visit=ostani_khozestan_pivot.iloc[0:10 , [0, 3]]
    ostani_khozestan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_khozestan_popular_duration=ostani_khozestan_pivot.iloc[0:10 , [0, 5]]
    
    ostani_khozestan_popular_visit = ostani_khozestan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی خوزستان', 'نام برنامه': 'محتواهای پربازدید استانی خوزستان'})
    ostani_khozestan_popular_duration = ostani_khozestan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی خوزستان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی خوزستان'})
    
    print("ostani_dena")
    ostani_dena=ostani.query("channel == 'استانی دنا'")
    ostani_dena_visit=ostani_dena['تعداد بازدید'].sum()
    ostani_dena_duration=ostani_dena['مدت بازدید'].sum()
    ostani_dena_duration=round(ostani_dena_duration, 0)
    ostani_dena_content=ostani_dena.copy()
    ostani_dena_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_dena_content=len(ostani_dena_content)
    ostani_dena_pivot=ostani_dena.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_dena_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_dena_popular_visit=ostani_dena_pivot.iloc[0:10 , [0, 3]]
    ostani_dena_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_dena_popular_duration=ostani_dena_pivot.iloc[0:10 , [0, 5]]
    
    ostani_dena_popular_visit = ostani_dena_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی دنا', 'نام برنامه': 'محتواهای پربازدید استانی دنا'})
    ostani_dena_popular_duration = ostani_dena_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی دنا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی دنا'})
    
    print("ostani_sabalan")
    ostani_sabalan=ostani.query("channel == 'استانی سبلان'")
    ostani_sabalan_visit=ostani_sabalan['تعداد بازدید'].sum()
    ostani_sabalan_duration=ostani_sabalan['مدت بازدید'].sum()
    ostani_sabalan_duration=round(ostani_sabalan_duration, 0)
    ostani_sabalan_content=ostani_sabalan.copy()
    ostani_sabalan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_sabalan_content=len(ostani_sabalan_content)
    ostani_sabalan_pivot=ostani_sabalan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_sabalan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_sabalan_popular_visit=ostani_sabalan_pivot.iloc[0:10 , [0, 3]]
    ostani_sabalan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_sabalan_popular_duration=ostani_sabalan_pivot.iloc[0:10 , [0, 5]]
    
    ostani_sabalan_popular_visit = ostani_sabalan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی سبلان', 'نام برنامه': 'محتواهای پربازدید استانی سبلان'})
    ostani_sabalan_popular_duration = ostani_sabalan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی سبلان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی سبلان'})
    
    print("ostani_sahand")
    ostani_sahand=ostani.query("channel == 'استانی سهند'")
    ostani_sahand_visit=ostani_sahand['تعداد بازدید'].sum()
    ostani_sahand_duration=ostani_sahand['مدت بازدید'].sum()
    ostani_sahand_duration=round(ostani_sahand_duration, 0)
    ostani_sahand_content=ostani_sahand.copy()
    ostani_sahand_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_sahand_content=len(ostani_sahand_content)
    ostani_sahand_pivot=ostani_sahand.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_sahand_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_sahand_popular_visit=ostani_sahand_pivot.iloc[0:10 , [0, 3]]
    ostani_sahand_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_sahand_popular_duration=ostani_sahand_pivot.iloc[0:10 , [0, 5]]
    
    ostani_sahand_popular_visit = ostani_sahand_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی سهند', 'نام برنامه': 'محتواهای پربازدید استانی سهند'})
    ostani_sahand_popular_duration = ostani_sahand_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی سهند (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی سهند'})
    
    print("ostani_fars")
    ostani_fars=ostani.query("channel == 'استانی فارس'")
    ostani_fars_visit=ostani_fars['تعداد بازدید'].sum()
    ostani_fars_duration=ostani_fars['مدت بازدید'].sum()
    ostani_fars_duration=round(ostani_fars_duration, 0)
    ostani_fars_content=ostani_fars.copy()
    ostani_fars_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_fars_content=len(ostani_fars_content)
    ostani_fars_pivot=ostani_fars.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_fars_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_fars_popular_visit=ostani_fars_pivot.iloc[0:10 , [0, 3]]
    ostani_fars_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_fars_popular_duration=ostani_fars_pivot.iloc[0:10 , [0, 5]]
    
    ostani_fars_popular_visit = ostani_fars_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی فارس', 'نام برنامه': 'محتواهای پربازدید استانی فارس'})
    ostani_fars_popular_duration = ostani_fars_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی فارس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی فارس'})
    
    print("ostani_ghazvin")
    ostani_ghazvin=ostani.query("channel == 'استانی قزوین'")
    ostani_ghazvin_visit=ostani_ghazvin['تعداد بازدید'].sum()
    ostani_ghazvin_duration=ostani_ghazvin['مدت بازدید'].sum()
    ostani_ghazvin_duration=round(ostani_ghazvin_duration, 0)
    ostani_ghazvin_content=ostani_ghazvin.copy()
    ostani_ghazvin_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_ghazvin_content=len(ostani_ghazvin_content)
    ostani_ghazvin_pivot=ostani_ghazvin.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_ghazvin_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_ghazvin_popular_visit=ostani_ghazvin_pivot.iloc[0:10 , [0, 3]]
    ostani_ghazvin_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_ghazvin_popular_duration=ostani_ghazvin_pivot.iloc[0:10 , [0, 5]]
    
    ostani_ghazvin_popular_visit = ostani_ghazvin_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی قزوین', 'نام برنامه': 'محتواهای پربازدید استانی قزوین'})
    ostani_ghazvin_popular_duration = ostani_ghazvin_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی قزوین (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی قزوین'})
    
    print("ostani_kordestan")
    ostani_kordestan=ostani.query("channel == 'استانی کردستان'")
    ostani_kordestan_visit=ostani_kordestan['تعداد بازدید'].sum()
    ostani_kordestan_duration=ostani_kordestan['مدت بازدید'].sum()
    ostani_kordestan_duration=round(ostani_kordestan_duration, 0)
    ostani_kordestan_content=ostani_kordestan.copy()
    ostani_kordestan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_kordestan_content=len(ostani_kordestan_content)
    ostani_kordestan_pivot=ostani_kordestan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_kordestan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_kordestan_popular_visit=ostani_kordestan_pivot.iloc[0:10 , [0, 3]]
    ostani_kordestan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_kordestan_popular_duration=ostani_kordestan_pivot.iloc[0:10 , [0, 5]]
    
    ostani_kordestan_popular_visit = ostani_kordestan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی کردستان', 'نام برنامه': 'محتواهای پربازدید استانی کردستان'})
    ostani_kordestan_popular_duration = ostani_kordestan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی کردستان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی کردستان'})
    
    print("ostani_kermanshah")
    ostani_kermanshah=ostani.query("channel == 'استانی کرمانشاه'")
    ostani_kermanshah_visit=ostani_kermanshah['تعداد بازدید'].sum()
    ostani_kermanshah_duration=ostani_kermanshah['مدت بازدید'].sum()
    ostani_kermanshah_duration=round(ostani_kermanshah_duration, 0)
    ostani_kermanshah_content=ostani_kermanshah.copy()
    ostani_kermanshah_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_kermanshah_content=len(ostani_kermanshah_content)
    ostani_kermanshah_pivot=ostani_kermanshah.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_kermanshah_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_kermanshah_popular_visit=ostani_kermanshah_pivot.iloc[0:10 , [0, 3]]
    ostani_kermanshah_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_kermanshah_popular_duration=ostani_kermanshah_pivot.iloc[0:10 , [0, 5]]
    
    ostani_kermanshah_popular_visit = ostani_kermanshah_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی کرمانشاه', 'نام برنامه': 'محتواهای پربازدید استانی کرمانشاه'})
    ostani_kermanshah_popular_duration = ostani_kermanshah_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی کرمانشاه (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی کرمانشاه'})
    
    print("ostani_kish")
    ostani_kish=ostani.query("channel == 'استانی کیش'")
    ostani_kish_visit=ostani_kish['تعداد بازدید'].sum()
    ostani_kish_duration=ostani_kish['مدت بازدید'].sum()
    ostani_kish_duration=round(ostani_kish_duration, 0)
    ostani_kish_content=ostani_kish.copy()
    ostani_kish_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_kish_content=len(ostani_kish_content)
    ostani_kish_pivot=ostani_kish.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_kish_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_kish_popular_visit=ostani_kish_pivot.iloc[0:10 , [0, 3]]
    ostani_kish_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_kish_popular_duration=ostani_kish_pivot.iloc[0:10 , [0, 5]]
    
    ostani_kish_popular_visit = ostani_kish_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی کیش', 'نام برنامه': 'محتواهای پربازدید استانی کیش'})
    ostani_kish_popular_duration = ostani_kish_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی کیش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی کیش'})
    
    print("ostani_mazandaran")
    ostani_mazandaran=ostani.query("channel == 'استانی مازندران'")
    ostani_mazandaran_visit=ostani_mazandaran['تعداد بازدید'].sum()
    ostani_mazandaran_duration=ostani_mazandaran['مدت بازدید'].sum()
    ostani_mazandaran_duration=round(ostani_mazandaran_duration, 0)
    ostani_mazandaran_content=ostani_mazandaran.copy()
    ostani_mazandaran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_mazandaran_content=len(ostani_mazandaran_content)
    ostani_mazandaran_pivot=ostani_mazandaran.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_mazandaran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_mazandaran_popular_visit=ostani_mazandaran_pivot.iloc[0:10 , [0, 3]]
    ostani_mazandaran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_mazandaran_popular_duration=ostani_mazandaran_pivot.iloc[0:10 , [0, 5]]
    
    ostani_mazandaran_popular_visit = ostani_mazandaran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی مازندران', 'نام برنامه': 'محتواهای پربازدید استانی مازندران'})
    ostani_mazandaran_popular_duration = ostani_mazandaran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی مازندران (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی مازندران'})
    
    print("ostani_hamedan")
    ostani_hamedan=ostani.query("channel == 'استانی همدان'")
    ostani_hamedan_visit=ostani_hamedan['تعداد بازدید'].sum()
    ostani_hamedan_duration=ostani_hamedan['مدت بازدید'].sum()
    ostani_hamedan_duration=round(ostani_hamedan_duration, 0)
    ostani_hamedan_content=ostani_hamedan.copy()
    ostani_hamedan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_hamedan_content=len(ostani_hamedan_content)
    ostani_hamedan_pivot=ostani_hamedan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_hamedan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_hamedan_popular_visit=ostani_hamedan_pivot.iloc[0:10 , [0, 3]]
    ostani_hamedan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_hamedan_popular_duration=ostani_hamedan_pivot.iloc[0:10 , [0, 5]]
    
    ostani_hamedan_popular_visit = ostani_hamedan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی همدان', 'نام برنامه': 'محتواهای پربازدید استانی همدان'})
    ostani_hamedan_popular_duration = ostani_hamedan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی همدان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی همدان'})
    
    
    
    
    
    
    
    
    
    
    print("ostani_jahanbin")
    ostani_jahanbin=ostani.query("channel == 'استانی چهار محال بختیاری - جهان بین'")
    ostani_jahanbin_visit=ostani_jahanbin['تعداد بازدید'].sum()
    ostani_jahanbin_duration=ostani_jahanbin['مدت بازدید'].sum()
    ostani_jahanbin_duration=round(ostani_jahanbin_duration, 0)
    ostani_jahanbin_content=ostani_jahanbin.copy()
    ostani_jahanbin_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_jahanbin_content=len(ostani_jahanbin_content)
    ostani_jahanbin_pivot=ostani_jahanbin.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_jahanbin_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_jahanbin_popular_visit=ostani_jahanbin_pivot.iloc[0:10 , [0, 3]]
    ostani_jahanbin_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_jahanbin_popular_duration=ostani_jahanbin_pivot.iloc[0:10 , [0, 5]]
    
    ostani_jahanbin_popular_visit = ostani_jahanbin_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی چهار محال بختیاری - جهان بین', 'نام برنامه': 'محتواهای پربازدید استانی چهار محال بختیاری - جهان بین'})
    ostani_jahanbin_popular_duration = ostani_jahanbin_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی چهار محال بختیاری - جهان بین (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی چهار محال بختیاری - جهان بین'})
    
    print("ostani_khalij_fars")
    ostani_khalij_fars=ostani.query("channel == 'استانی خلیج فارس'")
    ostani_khalij_fars_visit=ostani_khalij_fars['تعداد بازدید'].sum()
    ostani_khalij_fars_duration=ostani_khalij_fars['مدت بازدید'].sum()
    ostani_khalij_fars_duration=round(ostani_khalij_fars_duration, 0)
    ostani_khalij_fars_content=ostani_khalij_fars.copy()
    ostani_khalij_fars_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_khalij_fars_content=len(ostani_khalij_fars_content)
    ostani_khalij_fars_pivot=ostani_khalij_fars.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_khalij_fars_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_khalij_fars_popular_visit=ostani_khalij_fars_pivot.iloc[0:10 , [0, 3]]
    ostani_khalij_fars_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_khalij_fars_popular_duration=ostani_khalij_fars_pivot.iloc[0:10 , [0, 5]]
    
    ostani_khalij_fars_popular_visit = ostani_khalij_fars_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی خلیج فارس', 'نام برنامه': 'محتواهای پربازدید استانی خلیج فارس'})
    ostani_khalij_fars_popular_duration = ostani_khalij_fars_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی خلیج فارس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی خلیج فارس'})
    
    print("ostani_dena")
    ostani_dena=ostani.query("channel == 'استانی کهگیلویه و بویر احمد - دنا'")
    ostani_dena_visit=ostani_dena['تعداد بازدید'].sum()
    ostani_dena_duration=ostani_dena['مدت بازدید'].sum()
    ostani_dena_duration=round(ostani_dena_duration, 0)
    ostani_dena_content=ostani_dena.copy()
    ostani_dena_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_dena_content=len(ostani_dena_content)
    ostani_dena_pivot=ostani_dena.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_dena_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_dena_popular_visit=ostani_dena_pivot.iloc[0:10 , [0, 3]]
    ostani_dena_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_dena_popular_duration=ostani_dena_pivot.iloc[0:10 , [0, 5]]
    
    ostani_dena_popular_visit = ostani_dena_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی کهگیلویه و بویر احمد - دنا', 'نام برنامه': 'محتواهای پربازدید استانی کهگیلویه و بویر احمد - دنا'})
    ostani_dena_popular_duration = ostani_dena_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی کهگیلویه و بویر احمد - دنا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی کهگیلویه و بویر احمد - دنا'})
    
    print("ostani_aftab")
    ostani_aftab=ostani.query("channel == 'استانی مرکزی-آفتاب'")
    ostani_aftab_visit=ostani_aftab['تعداد بازدید'].sum()
    ostani_aftab_duration=ostani_aftab['مدت بازدید'].sum()
    ostani_aftab_duration=round(ostani_aftab_duration, 0)
    ostani_aftab_content=ostani_aftab.copy()
    ostani_aftab_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_aftab_content=len(ostani_aftab_content)
    ostani_aftab_pivot=ostani_aftab.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_aftab_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_aftab_popular_visit=ostani_aftab_pivot.iloc[0:10 , [0, 3]]
    ostani_aftab_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_aftab_popular_duration=ostani_aftab_pivot.iloc[0:10 , [0, 5]]
    
    ostani_aftab_popular_visit = ostani_aftab_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی مرکزی-آفتاب', 'نام برنامه': 'محتواهای پربازدید استانی مرکزی-آفتاب'})
    ostani_aftab_popular_duration = ostani_aftab_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی مرکزی-آفتاب (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی مرکزی-آفتاب'})
    
    print("ostani_sabz")
    ostani_sabz=ostani.query("channel == 'استانی گلستان-سبز'")
    ostani_sabz_visit=ostani_sabz['تعداد بازدید'].sum()
    ostani_sabz_duration=ostani_sabz['مدت بازدید'].sum()
    ostani_sabz_duration=round(ostani_sabz_duration, 0)
    ostani_sabz_content=ostani_sabz.copy()
    ostani_sabz_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_sabz_content=len(ostani_sabz_content)
    ostani_sabz_pivot=ostani_sabz.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_sabz_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_sabz_popular_visit=ostani_sabz_pivot.iloc[0:10 , [0, 3]]
    ostani_sabz_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_sabz_popular_duration=ostani_sabz_pivot.iloc[0:10 , [0, 5]]
    
    ostani_sabz_popular_visit = ostani_sabz_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی گلستان-سبز', 'نام برنامه': 'محتواهای پربازدید استانی گلستان-سبز'})
    ostani_sabz_popular_duration = ostani_sabz_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی گلستان-سبز (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی گلستان-سبز'})
    
    print("ostani_semnan")
    ostani_semnan=ostani.query("channel == 'استانی سمنان'")
    ostani_semnan_visit=ostani_semnan['تعداد بازدید'].sum()
    ostani_semnan_duration=ostani_semnan['مدت بازدید'].sum()
    ostani_semnan_duration=round(ostani_semnan_duration, 0)
    ostani_semnan_content=ostani_semnan.copy()
    ostani_semnan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_semnan_content=len(ostani_semnan_content)
    ostani_semnan_pivot=ostani_semnan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_semnan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_semnan_popular_visit=ostani_semnan_pivot.iloc[0:10 , [0, 3]]
    ostani_semnan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_semnan_popular_duration=ostani_semnan_pivot.iloc[0:10 , [0, 5]]
    
    ostani_semnan_popular_visit = ostani_semnan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی سمنان', 'نام برنامه': 'محتواهای پربازدید استانی سمنان'})
    ostani_semnan_popular_duration = ostani_semnan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی سمنان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی سمنان'})
    
    print("ostani_noor")
    ostani_noor=ostani.query("channel == 'استانی قم-نور'")
    ostani_noor_visit=ostani_noor['تعداد بازدید'].sum()
    ostani_noor_duration=ostani_noor['مدت بازدید'].sum()
    ostani_noor_duration=round(ostani_noor_duration, 0)
    ostani_noor_content=ostani_noor.copy()
    ostani_noor_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_noor_content=len(ostani_noor_content)
    ostani_noor_pivot=ostani_noor.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_noor_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_noor_popular_visit=ostani_noor_pivot.iloc[0:10 , [0, 3]]
    ostani_noor_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_noor_popular_duration=ostani_noor_pivot.iloc[0:10 , [0, 5]]
    
    ostani_noor_popular_visit = ostani_noor_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی قم-نور', 'نام برنامه': 'محتواهای پربازدید استانی قم-نور'})
    ostani_noor_popular_duration = ostani_noor_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی قم-نور (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی قم-نور'})
    
    print("ostani_eshragh")
    ostani_eshragh=ostani.query("channel == 'استانی زنجان-اشراق'")
    ostani_eshragh_visit=ostani_eshragh['تعداد بازدید'].sum()
    ostani_eshragh_duration=ostani_eshragh['مدت بازدید'].sum()
    ostani_eshragh_duration=round(ostani_eshragh_duration, 0)
    ostani_eshragh_content=ostani_eshragh.copy()
    ostani_eshragh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_eshragh_content=len(ostani_eshragh_content)
    ostani_eshragh_pivot=ostani_eshragh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_eshragh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_eshragh_popular_visit=ostani_eshragh_pivot.iloc[0:10 , [0, 3]]
    ostani_eshragh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_eshragh_popular_duration=ostani_eshragh_pivot.iloc[0:10 , [0, 5]]
    
    ostani_eshragh_popular_visit = ostani_eshragh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی زنجان-اشراق', 'نام برنامه': 'محتواهای پربازدید استانی زنجان-اشراق'})
    ostani_eshragh_popular_duration = ostani_eshragh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی زنجان-اشراق (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی زنجان-اشراق'})
    
    print("ostani_khavaran")
    ostani_khavaran=ostani.query("channel == 'استانی خراسان جنوبی -خاوران'")
    ostani_khavaran_visit=ostani_khavaran['تعداد بازدید'].sum()
    ostani_khavaran_duration=ostani_khavaran['مدت بازدید'].sum()
    ostani_khavaran_duration=round(ostani_khavaran_duration, 0)
    ostani_khavaran_content=ostani_khavaran.copy()
    ostani_khavaran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_khavaran_content=len(ostani_khavaran_content)
    ostani_khavaran_pivot=ostani_khavaran.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_khavaran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_khavaran_popular_visit=ostani_khavaran_pivot.iloc[0:10 , [0, 3]]
    ostani_khavaran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_khavaran_popular_duration=ostani_khavaran_pivot.iloc[0:10 , [0, 5]]
    
    ostani_khavaran_popular_visit = ostani_khavaran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی خراسان جنوبی -خاوران', 'نام برنامه': 'محتواهای پربازدید استانی خراسان جنوبی -خاوران'})
    ostani_khavaran_popular_duration = ostani_khavaran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی خراسان جنوبی -خاوران (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی خراسان جنوبی -خاوران'})
    
    print("ostani_sahand")
    ostani_sahand=ostani.query("channel == 'استانی آذربایجان شرقی - سهند'")
    ostani_sahand_visit=ostani_sahand['تعداد بازدید'].sum()
    ostani_sahand_duration=ostani_sahand['مدت بازدید'].sum()
    ostani_sahand_duration=round(ostani_sahand_duration, 0)
    ostani_sahand_content=ostani_sahand.copy()
    ostani_sahand_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_sahand_content=len(ostani_sahand_content)
    ostani_sahand_pivot=ostani_sahand.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_sahand_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_sahand_popular_visit=ostani_sahand_pivot.iloc[0:10 , [0, 3]]
    ostani_sahand_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_sahand_popular_duration=ostani_sahand_pivot.iloc[0:10 , [0, 5]]
    
    ostani_sahand_popular_visit = ostani_sahand_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی آذربایجان شرقی - سهند', 'نام برنامه': 'محتواهای پربازدید استانی آذربایجان شرقی - سهند'})
    ostani_sahand_popular_duration = ostani_sahand_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی آذربایجان شرقی - سهند (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی آذربایجان شرقی - سهند'})
    
    print("ostani_kerman")
    ostani_kerman=ostani.query("channel == 'استانی کرمان'")
    ostani_kerman_visit=ostani_kerman['تعداد بازدید'].sum()
    ostani_kerman_duration=ostani_kerman['مدت بازدید'].sum()
    ostani_kerman_duration=round(ostani_kerman_duration, 0)
    ostani_kerman_content=ostani_kerman.copy()
    ostani_kerman_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_kerman_content=len(ostani_kerman_content)
    ostani_kerman_pivot=ostani_kerman.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_kerman_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_kerman_popular_visit=ostani_kerman_pivot.iloc[0:10 , [0, 3]]
    ostani_kerman_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_kerman_popular_duration=ostani_kerman_pivot.iloc[0:10 , [0, 5]]
    
    ostani_kerman_popular_visit = ostani_kerman_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی کرمان', 'نام برنامه': 'محتواهای پربازدید استانی کرمان'})
    ostani_kerman_popular_duration = ostani_kerman_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی کرمان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی کرمان'})
    
    print("ostani_mahabad")
    ostani_mahabad=ostani.query("channel == 'استانی مهاباد'")
    ostani_mahabad_visit=ostani_mahabad['تعداد بازدید'].sum()
    ostani_mahabad_duration=ostani_mahabad['مدت بازدید'].sum()
    ostani_mahabad_duration=round(ostani_mahabad_duration, 0)
    ostani_mahabad_content=ostani_mahabad.copy()
    ostani_mahabad_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_mahabad_content=len(ostani_mahabad_content)
    ostani_mahabad_pivot=ostani_mahabad.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_mahabad_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_mahabad_popular_visit=ostani_mahabad_pivot.iloc[0:10 , [0, 3]]
    ostani_mahabad_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_mahabad_popular_duration=ostani_mahabad_pivot.iloc[0:10 , [0, 5]]
    
    ostani_mahabad_popular_visit = ostani_mahabad_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی مهاباد', 'نام برنامه': 'محتواهای پربازدید استانی مهاباد'})
    ostani_mahabad_popular_duration = ostani_mahabad_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی مهاباد (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی مهاباد'})
    
    print("ostani_hamon")
    ostani_hamon=ostani.query("channel == 'استانی هامون'")
    ostani_hamon_visit=ostani_hamon['تعداد بازدید'].sum()
    ostani_hamon_duration=ostani_hamon['مدت بازدید'].sum()
    ostani_hamon_duration=round(ostani_hamon_duration, 0)
    ostani_hamon_content=ostani_hamon.copy()
    ostani_hamon_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_hamon_content=len(ostani_hamon_content)
    ostani_hamon_pivot=ostani_hamon.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_hamon_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_hamon_popular_visit=ostani_hamon_pivot.iloc[0:10 , [0, 3]]
    ostani_hamon_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_hamon_popular_duration=ostani_hamon_pivot.iloc[0:10 , [0, 5]]
    
    ostani_hamon_popular_visit = ostani_hamon_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی هامون', 'نام برنامه': 'محتواهای پربازدید استانی هامون'})
    ostani_hamon_popular_duration = ostani_hamon_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی هامون (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی هامون'})
    
    print("ostani_atrak")
    ostani_atrak=ostani.query("channel == 'استانی خراسان شمالی -اترک'")
    ostani_atrak_visit=ostani_atrak['تعداد بازدید'].sum()
    ostani_atrak_duration=ostani_atrak['مدت بازدید'].sum()
    ostani_atrak_duration=round(ostani_atrak_duration, 0)
    ostani_atrak_content=ostani_atrak.copy()
    ostani_atrak_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_atrak_content=len(ostani_atrak_content)
    ostani_atrak_pivot=ostani_atrak.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_atrak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_atrak_popular_visit=ostani_atrak_pivot.iloc[0:10 , [0, 3]]
    ostani_atrak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_atrak_popular_duration=ostani_atrak_pivot.iloc[0:10 , [0, 5]]
    
    ostani_atrak_popular_visit = ostani_atrak_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی خراسان شمالی -اترک', 'نام برنامه': 'محتواهای پربازدید استانی خراسان شمالی -اترک'})
    ostani_atrak_popular_duration = ostani_atrak_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی خراسان شمالی -اترک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی خراسان شمالی -اترک'})
    
    print("ostani_aflak")
    ostani_aflak=ostani.query("channel == 'استانی لرستان - افلاک'")
    ostani_aflak_visit=ostani_aflak['تعداد بازدید'].sum()
    ostani_aflak_duration=ostani_aflak['مدت بازدید'].sum()
    ostani_aflak_duration=round(ostani_aflak_duration, 0)
    ostani_aflak_content=ostani_aflak.copy()
    ostani_aflak_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ostani_aflak_content=len(ostani_aflak_content)
    ostani_aflak_pivot=ostani_aflak.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ostani_aflak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_aflak_popular_visit=ostani_aflak_pivot.iloc[0:10 , [0, 3]]
    ostani_aflak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_aflak_popular_duration=ostani_aflak_pivot.iloc[0:10 , [0, 5]]
    
    ostani_aflak_popular_visit = ostani_aflak_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید استانی لرستان - افلاک', 'نام برنامه': 'محتواهای پربازدید استانی لرستان - افلاک'})
    ostani_aflak_popular_duration = ostani_aflak_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید استانی لرستان - افلاک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید استانی لرستان - افلاک'})
    
    
    
    print("dataframe ostani channels")
    ostani_channels_statistics={'channel_name': ['استانی آبادان',
                                                 'استانی آذربایجان غربی',
                                                 'استانی اصفهان',
                                                 'استانی افلاک',
                                                 'استانی البرز',
                                         'استانی ایلام',
                                         'استانی باران',
                                         'استانی بوشهر',
                                         'استانی تابان',
                                         'استانی خراسان رضوی',
                                         'استانی خوزستان',
                                         'استانی دنا',
                                         'استانی سبلان',
                                         'استانی سهند',
                                         'استانی فارس',
                                         'استانی قزوین',
                                         'استانی کردستان',
                                         'استانی کرمانشاه',
                                         'استانی کیش',
                                         'استانی مازندران',
                                         'استانی همدان',
                                         'استانی چهار محال بختیاری - جهان بین',
                                         'استانی خلیج فارس',
                                         'استانی کهگیلویه و بویر احمد - دنا',
                                         'استانی مرکزی-آفتاب',
                                         'استانی گلستان-سبز',
                                         'استانی سمنان',
                                         'استانی قم-نور',
                                         'استانی زنجان-اشراق',
                                         'استانی خراسان جنوبی -خاوران',
                                         'استانی آذربایجان شرقی - سهند',
                                         'استانی کرمان',
                                         'استانی مهاباد',
                                         'استانی هامون',
                                         'استانی خراسان شمالی -اترک',
                                         'استانی لرستان - افلاک'],
           'channel_content': [ostani_abadan_content, ostani_azarbayjan_gharbi_content, ostani_esfahan_content, 
                               ostani_aflak_content, ostani_alborz_content, ostani_ilam_content,
                               ostani_baran_content, ostani_boshehr_content, ostani_taban_content, 
                               ostani_khorasan_razavi_content, ostani_khozestan_content, ostani_dena_content,
                               ostani_sabalan_content, ostani_sahand_content, ostani_fars_content,
                               ostani_ghazvin_content,ostani_kordestan_content, ostani_kermanshah_content, 
                               ostani_kish_content, ostani_mazandaran_content, ostani_hamedan_content,
                               ostani_jahanbin_content, ostani_khalij_fars_content, ostani_dena_content,
                               ostani_aftab_content, ostani_sabz_content, ostani_semnan_content,
                               ostani_noor_content, ostani_eshragh_content, ostani_khavaran_content,
                               ostani_sahand_content, ostani_kerman_content, ostani_mahabad_content,
                               ostani_hamon_content, ostani_atrak_content, ostani_aflak_content],
           'channel_visit': [ostani_abadan_visit, ostani_azarbayjan_gharbi_visit, ostani_esfahan_visit, 
                               ostani_aflak_visit, ostani_alborz_visit, ostani_ilam_visit,
                               ostani_baran_visit, ostani_boshehr_visit, ostani_taban_visit, 
                               ostani_khorasan_razavi_visit, ostani_khozestan_visit, ostani_dena_visit,
                               ostani_sabalan_visit, ostani_sahand_visit, ostani_fars_visit,
                               ostani_ghazvin_visit,ostani_kordestan_visit, ostani_kermanshah_visit, 
                               ostani_kish_visit, ostani_mazandaran_visit, ostani_hamedan_visit,
                               ostani_jahanbin_visit, ostani_khalij_fars_visit, ostani_dena_visit,
                               ostani_aftab_visit, ostani_sabz_visit, ostani_semnan_visit,
                               ostani_noor_visit, ostani_eshragh_visit, ostani_khavaran_visit,
                               ostani_sahand_visit, ostani_kerman_visit, ostani_mahabad_visit,
                               ostani_hamon_visit, ostani_atrak_visit, ostani_aflak_visit],
           'channel_duration': [ostani_abadan_duration, ostani_azarbayjan_gharbi_duration, ostani_esfahan_duration, 
                               ostani_aflak_duration, ostani_alborz_duration, ostani_ilam_duration,
                               ostani_baran_duration, ostani_boshehr_duration, ostani_taban_duration, 
                               ostani_khorasan_razavi_duration, ostani_khozestan_duration, ostani_dena_duration,
                               ostani_sabalan_duration, ostani_sahand_duration, ostani_fars_duration,
                               ostani_ghazvin_duration,ostani_kordestan_duration, ostani_kermanshah_duration, 
                               ostani_kish_duration, ostani_mazandaran_duration, ostani_hamedan_duration,
                               ostani_jahanbin_duration, ostani_khalij_fars_duration, ostani_dena_duration,
                               ostani_aftab_duration, ostani_sabz_duration, ostani_semnan_duration,
                               ostani_noor_duration, ostani_eshragh_duration, ostani_khavaran_duration,
                               ostani_sahand_duration, ostani_kerman_duration, ostani_mahabad_duration,
                               ostani_hamon_duration, ostani_atrak_duration, ostani_aflak_duration],}
    ostani_channels_statistics=pd.DataFrame(ostani_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
    ostani_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
    ostani_channels_statistics=ostani_channels_statistics.rename(columns={'channel_name': 'نام شبکه', 'channel_content': 'تعداد محتوا', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})
    
    ostani_abadan_popular_visit = ostani_abadan_popular_visit.reset_index()
    del ostani_abadan_popular_visit['index']
    ostani_abadan_popular_duration = ostani_abadan_popular_duration.reset_index()
    del ostani_abadan_popular_duration['index']
    
    ostani_azarbayjan_gharbi_popular_visit = ostani_azarbayjan_gharbi_popular_visit.reset_index()
    del ostani_azarbayjan_gharbi_popular_visit['index']
    ostani_azarbayjan_gharbi_popular_duration = ostani_azarbayjan_gharbi_popular_duration.reset_index()
    del ostani_azarbayjan_gharbi_popular_duration['index']
    
    ostani_esfahan_popular_visit = ostani_esfahan_popular_visit.reset_index()
    del ostani_esfahan_popular_visit['index']
    ostani_esfahan_popular_duration = ostani_esfahan_popular_duration.reset_index()
    del ostani_esfahan_popular_duration['index']
    
    ostani_ilam_popular_visit = ostani_ilam_popular_visit.reset_index()
    del ostani_ilam_popular_visit['index']
    ostani_ilam_popular_duration = ostani_ilam_popular_duration.reset_index()
    del ostani_ilam_popular_duration['index']
    
    ostani_alborz_popular_visit = ostani_alborz_popular_visit.reset_index()
    del ostani_alborz_popular_visit['index']
    ostani_alborz_popular_duration = ostani_alborz_popular_duration.reset_index()
    del ostani_alborz_popular_duration['index']
    
    ostani_aflak_popular_visit = ostani_aflak_popular_visit.reset_index()
    del ostani_aflak_popular_visit['index']
    ostani_aflak_popular_duration = ostani_aflak_popular_duration.reset_index()
    del ostani_aflak_popular_duration['index']
    
    ostani_baran_popular_visit = ostani_baran_popular_visit.reset_index()
    del ostani_baran_popular_visit['index']
    ostani_baran_popular_duration = ostani_baran_popular_duration.reset_index()
    del ostani_baran_popular_duration['index']
    
    ostani_boshehr_popular_visit = ostani_boshehr_popular_visit.reset_index()
    del ostani_boshehr_popular_visit['index']
    ostani_boshehr_popular_duration = ostani_boshehr_popular_duration.reset_index()
    del ostani_boshehr_popular_duration['index']
    
    ostani_taban_popular_visit = ostani_taban_popular_visit.reset_index()
    del ostani_taban_popular_visit['index']
    ostani_taban_popular_duration = ostani_taban_popular_duration.reset_index()
    del ostani_taban_popular_duration['index']
    
    ostani_khorasan_razavi_popular_visit = ostani_khorasan_razavi_popular_visit.reset_index()
    del ostani_khorasan_razavi_popular_visit['index']
    ostani_khorasan_razavi_popular_duration = ostani_khorasan_razavi_popular_duration.reset_index()
    del ostani_khorasan_razavi_popular_duration['index']
    
    ostani_khozestan_popular_visit = ostani_khozestan_popular_visit.reset_index()
    del ostani_khozestan_popular_visit['index']
    ostani_khozestan_popular_duration = ostani_khozestan_popular_duration.reset_index()
    del ostani_khozestan_popular_duration['index']
    
    ostani_dena_popular_visit = ostani_dena_popular_visit.reset_index()
    del ostani_dena_popular_visit['index']
    ostani_dena_popular_duration = ostani_dena_popular_duration.reset_index()
    del ostani_dena_popular_duration['index']
    
    ostani_sabalan_popular_visit = ostani_sabalan_popular_visit.reset_index()
    del ostani_sabalan_popular_visit['index']
    ostani_sabalan_popular_duration = ostani_sabalan_popular_duration.reset_index()
    del ostani_sabalan_popular_duration['index']
    
    ostani_sahand_popular_visit = ostani_sahand_popular_visit.reset_index()
    del ostani_sahand_popular_visit['index']
    ostani_sahand_popular_duration = ostani_sahand_popular_duration.reset_index()
    del ostani_sahand_popular_duration['index']
    
    ostani_fars_popular_visit = ostani_fars_popular_visit.reset_index()
    del ostani_fars_popular_visit['index']
    ostani_fars_popular_duration = ostani_fars_popular_duration.reset_index()
    del ostani_fars_popular_duration['index']
    
    ostani_ghazvin_popular_visit = ostani_ghazvin_popular_visit.reset_index()
    del ostani_ghazvin_popular_visit['index']
    ostani_ghazvin_popular_duration = ostani_ghazvin_popular_duration.reset_index()
    del ostani_ghazvin_popular_duration['index']
    
    ostani_kordestan_popular_visit = ostani_kordestan_popular_visit.reset_index()
    del ostani_kordestan_popular_visit['index']
    ostani_kordestan_popular_duration = ostani_kordestan_popular_duration.reset_index()
    del ostani_kordestan_popular_duration['index']
    
    ostani_kermanshah_popular_visit = ostani_kermanshah_popular_visit.reset_index()
    del ostani_kermanshah_popular_visit['index']
    ostani_kermanshah_popular_duration = ostani_kermanshah_popular_duration.reset_index()
    del ostani_kermanshah_popular_duration['index']
    
    ostani_kish_popular_visit = ostani_kish_popular_visit.reset_index()
    del ostani_kish_popular_visit['index']
    ostani_kish_popular_duration = ostani_kish_popular_duration.reset_index()
    del ostani_kish_popular_duration['index']
    
    ostani_mazandaran_popular_visit = ostani_mazandaran_popular_visit.reset_index()
    del ostani_mazandaran_popular_visit['index']
    ostani_mazandaran_popular_duration = ostani_mazandaran_popular_duration.reset_index()
    del ostani_mazandaran_popular_duration['index']
    
    ostani_hamedan_popular_visit = ostani_hamedan_popular_visit.reset_index()
    del ostani_hamedan_popular_visit['index']
    ostani_hamedan_popular_duration = ostani_hamedan_popular_duration.reset_index()
    del ostani_hamedan_popular_duration['index']
    
    
    
    
    
    
    
    ostani_jahanbin_popular_visit = ostani_jahanbin_popular_visit.reset_index()
    del ostani_jahanbin_popular_visit['index']
    ostani_jahanbin_popular_duration = ostani_jahanbin_popular_duration.reset_index()
    del ostani_jahanbin_popular_duration['index']
    
    ostani_khalij_fars_popular_visit = ostani_khalij_fars_popular_visit.reset_index()
    del ostani_khalij_fars_popular_visit['index']
    ostani_khalij_fars_popular_duration = ostani_khalij_fars_popular_duration.reset_index()
    del ostani_khalij_fars_popular_duration['index']
    
    ostani_dena_popular_visit = ostani_dena_popular_visit.reset_index()
    del ostani_dena_popular_visit['index']
    ostani_dena_popular_duration = ostani_dena_popular_duration.reset_index()
    del ostani_dena_popular_duration['index']
    
    ostani_aftab_popular_visit = ostani_aftab_popular_visit.reset_index()
    del ostani_aftab_popular_visit['index']
    ostani_aftab_popular_duration = ostani_aftab_popular_duration.reset_index()
    del ostani_aftab_popular_duration['index']
    
    ostani_sabz_popular_visit = ostani_sabz_popular_visit.reset_index()
    del ostani_sabz_popular_visit['index']
    ostani_sabz_popular_duration = ostani_sabz_popular_duration.reset_index()
    del ostani_sabz_popular_duration['index']
    
    ostani_semnan_popular_visit = ostani_semnan_popular_visit.reset_index()
    del ostani_semnan_popular_visit['index']
    ostani_semnan_popular_duration = ostani_semnan_popular_duration.reset_index()
    del ostani_semnan_popular_duration['index']
    
    ostani_noor_popular_visit = ostani_noor_popular_visit.reset_index()
    del ostani_noor_popular_visit['index']
    ostani_noor_popular_duration = ostani_noor_popular_duration.reset_index()
    del ostani_noor_popular_duration['index']
    
    ostani_eshragh_popular_visit = ostani_eshragh_popular_visit.reset_index()
    del ostani_eshragh_popular_visit['index']
    ostani_eshragh_popular_duration = ostani_eshragh_popular_duration.reset_index()
    del ostani_eshragh_popular_duration['index']
    
    ostani_khavaran_popular_visit = ostani_khavaran_popular_visit.reset_index()
    del ostani_khavaran_popular_visit['index']
    ostani_khavaran_popular_duration = ostani_khavaran_popular_duration.reset_index()
    del ostani_khavaran_popular_duration['index']
    
    ostani_sahand_popular_visit = ostani_sahand_popular_visit.reset_index()
    del ostani_sahand_popular_visit['index']
    ostani_sahand_popular_duration = ostani_sahand_popular_duration.reset_index()
    del ostani_sahand_popular_duration['index']
    
    ostani_kerman_popular_visit = ostani_kerman_popular_visit.reset_index()
    del ostani_kerman_popular_visit['index']
    ostani_kerman_popular_duration = ostani_kerman_popular_duration.reset_index()
    del ostani_kerman_popular_duration['index']
    
    ostani_mahabad_popular_visit = ostani_mahabad_popular_visit.reset_index()
    del ostani_mahabad_popular_visit['index']
    ostani_mahabad_popular_duration = ostani_mahabad_popular_duration.reset_index()
    del ostani_mahabad_popular_duration['index']
    
    ostani_hamon_popular_visit = ostani_hamon_popular_visit.reset_index()
    del ostani_hamon_popular_visit['index']
    ostani_hamon_popular_duration = ostani_hamon_popular_duration.reset_index()
    del ostani_hamon_popular_duration['index']
    
    ostani_atrak_popular_visit = ostani_atrak_popular_visit.reset_index()
    del ostani_atrak_popular_visit['index']
    ostani_atrak_popular_duration = ostani_atrak_popular_duration.reset_index()
    del ostani_atrak_popular_duration['index']
    
    ostani_aflak_popular_visit = ostani_aflak_popular_visit.reset_index()
    del ostani_aflak_popular_visit['index']
    ostani_aflak_popular_duration = ostani_aflak_popular_duration.reset_index()
    del ostani_aflak_popular_duration['index']
    
        
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
                                              ostani_hamedan_popular_visit, ostani_hamedan_popular_duration,
                                              ostani_jahanbin_popular_visit, ostani_jahanbin_popular_duration,
                                              ostani_khalij_fars_popular_visit, ostani_khalij_fars_popular_duration,
                                              ostani_dena_popular_visit, ostani_dena_popular_duration,
                                              ostani_aftab_popular_visit, ostani_aftab_popular_duration,
                                              ostani_sabz_popular_visit, ostani_sabz_popular_duration,
                                              ostani_semnan_popular_visit, ostani_semnan_popular_duration,
                                              ostani_noor_popular_visit, ostani_noor_popular_duration,
                                              ostani_eshragh_popular_visit, ostani_eshragh_popular_duration,
                                              ostani_khavaran_popular_visit, ostani_khavaran_popular_duration,
                                              ostani_sahand_popular_visit, ostani_sahand_popular_duration,
                                              ostani_kerman_popular_visit, ostani_kerman_popular_duration,
                                              ostani_mahabad_popular_visit, ostani_mahabad_popular_duration,
                                              ostani_hamon_popular_visit, ostani_hamon_popular_duration,
                                              ostani_atrak_popular_visit, ostani_atrak_popular_duration,
                                              ostani_aflak_popular_visit, ostani_aflak_popular_duration,],axis=1)
    

    writer = pd.ExcelWriter('output/output.sending.hard/آمار استانی.xlsx', engine='xlsxwriter')
    ostani_channels_statistics.to_excel(writer, 'آمار شبکه های استانی', index=False)
    ostani_channels_popular_content.to_excel(writer, 'محتواهای پربازدید', index=False)
    writer.save()
    
    print("End ostani")
    
    
    return ostani_channels_statistics, ostani_channels_popular_content
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        