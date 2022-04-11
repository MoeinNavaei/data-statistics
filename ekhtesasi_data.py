    
def ekhtesasi_data(ekhtesasi):
            
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
        
        
    print("start EKHTESASI")
    
    print("ekhtesasi_perspolis")
    ekhtesasi_perspolis=ekhtesasi.query("channel == 'پرسپولیس'")
    ekhtesasi_perspolis_visit=ekhtesasi_perspolis['تعداد بازدید'].sum()
    ekhtesasi_perspolis_duration=ekhtesasi_perspolis['مدت بازدید'].sum()
    ekhtesasi_perspolis_duration=round(ekhtesasi_perspolis_duration, 0)
    ekhtesasi_perspolis_content=ekhtesasi_perspolis.copy()
    ekhtesasi_perspolis_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_perspolis_content=len(ekhtesasi_perspolis_content)
    ekhtesasi_perspolis_pivot=ekhtesasi_perspolis.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_perspolis_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_perspolis_popular_visit=ekhtesasi_perspolis_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_perspolis_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_perspolis_popular_duration=ekhtesasi_perspolis_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_perspolis_popular_visit = ekhtesasi_perspolis_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی پرسپولیس', 'نام برنامه': 'محتواهای پربازدید اختصاصی پرسپولیس'})
    ekhtesasi_perspolis_popular_duration = ekhtesasi_perspolis_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی پرسپولیس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی پرسپولیس'})
    
    print("ekhtesasi_esteghlal")
    ekhtesasi_esteghlal=ekhtesasi.query("channel == 'استقلال'")
    ekhtesasi_esteghlal_visit=ekhtesasi_esteghlal['تعداد بازدید'].sum()
    ekhtesasi_esteghlal_duration=ekhtesasi_esteghlal['مدت بازدید'].sum()
    ekhtesasi_esteghlal_duration=round(ekhtesasi_esteghlal_duration, 0)
    ekhtesasi_esteghlal_content=ekhtesasi_esteghlal.copy()
    ekhtesasi_esteghlal_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_esteghlal_content=len(ekhtesasi_esteghlal_content)
    ekhtesasi_esteghlal_pivot=ekhtesasi_esteghlal.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_esteghlal_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_esteghlal_popular_visit=ekhtesasi_esteghlal_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_esteghlal_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_esteghlal_popular_duration=ekhtesasi_esteghlal_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_esteghlal_popular_visit = ekhtesasi_esteghlal_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی استقلال', 'نام برنامه': 'محتواهای پربازدید اختصاصی استقلال'})
    ekhtesasi_esteghlal_popular_duration = ekhtesasi_esteghlal_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی استقلال (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی استقلال'})
    
    print("ekhtesasi_shaparak")
    ekhtesasi_shaparak=ekhtesasi.query("channel == 'شاپرک'")
    ekhtesasi_shaparak_visit=ekhtesasi_shaparak['تعداد بازدید'].sum()
    ekhtesasi_shaparak_duration=ekhtesasi_shaparak['مدت بازدید'].sum()
    ekhtesasi_shaparak_duration=round(ekhtesasi_shaparak_duration, 0)
    ekhtesasi_shaparak_content=ekhtesasi_shaparak.copy()
    ekhtesasi_shaparak_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_shaparak_content=len(ekhtesasi_shaparak_content)
    ekhtesasi_shaparak_pivot=ekhtesasi_shaparak.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_shaparak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_shaparak_popular_visit=ekhtesasi_shaparak_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_shaparak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_shaparak_popular_duration=ekhtesasi_shaparak_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_shaparak_popular_visit = ekhtesasi_shaparak_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی شاپرک', 'نام برنامه': 'محتواهای پربازدید اختصاصی شاپرک'})
    ekhtesasi_shaparak_popular_duration = ekhtesasi_shaparak_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی شاپرک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی شاپرک'})
    
    print("ekhtesasi_shetab")
    ekhtesasi_shetab=ekhtesasi.query("channel == 'شتاب'")
    ekhtesasi_shetab_visit=ekhtesasi_shetab['تعداد بازدید'].sum()
    ekhtesasi_shetab_duration=ekhtesasi_shetab['مدت بازدید'].sum()
    ekhtesasi_shetab_duration=round(ekhtesasi_shetab_duration, 0)
    ekhtesasi_shetab_content=ekhtesasi_shetab.copy()
    ekhtesasi_shetab_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_shetab_content=len(ekhtesasi_shetab_content)
    ekhtesasi_shetab_pivot=ekhtesasi_shetab.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_shetab_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_shetab_popular_visit=ekhtesasi_shetab_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_shetab_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_shetab_popular_duration=ekhtesasi_shetab_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_shetab_popular_visit = ekhtesasi_shetab_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی شتاب', 'نام برنامه': 'محتواهای پربازدید اختصاصی شتاب'})
    ekhtesasi_shetab_popular_duration = ekhtesasi_shetab_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی شتاب (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی شتاب'})
    
    print("ekhtesasi_kodak_digiton")
    ekhtesasi_kodak_digiton=ekhtesasi.query("channel == 'کودک دیجیتون'")
    ekhtesasi_kodak_digiton_visit=ekhtesasi_kodak_digiton['تعداد بازدید'].sum()
    ekhtesasi_kodak_digiton_duration=ekhtesasi_kodak_digiton['مدت بازدید'].sum()
    ekhtesasi_kodak_digiton_duration=round(ekhtesasi_kodak_digiton_duration, 0)
    ekhtesasi_kodak_digiton_content=ekhtesasi_kodak_digiton.copy()
    ekhtesasi_kodak_digiton_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_kodak_digiton_content=len(ekhtesasi_kodak_digiton_content)
    ekhtesasi_kodak_digiton_pivot=ekhtesasi_kodak_digiton.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_kodak_digiton_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_kodak_digiton_popular_visit=ekhtesasi_kodak_digiton_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_kodak_digiton_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_kodak_digiton_popular_duration=ekhtesasi_kodak_digiton_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_kodak_digiton_popular_visit = ekhtesasi_kodak_digiton_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی کودک دیجیتون', 'نام برنامه': 'محتواهای پربازدید اختصاصی کودک دیجیتون'})
    ekhtesasi_kodak_digiton_popular_duration = ekhtesasi_kodak_digiton_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی کودک دیجیتون (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی کودک دیجیتون'})
    
    print("ekhtesasi_lenz_sport")
    ekhtesasi_lenz_sport=ekhtesasi.query("channel == 'لنزاسپورت'")
    ekhtesasi_lenz_sport_visit=ekhtesasi_lenz_sport['تعداد بازدید'].sum()
    ekhtesasi_lenz_sport_duration=ekhtesasi_lenz_sport['مدت بازدید'].sum()
    ekhtesasi_lenz_sport_duration=round(ekhtesasi_lenz_sport_duration, 0)
    ekhtesasi_lenz_sport_content=ekhtesasi_lenz_sport.copy()
    ekhtesasi_lenz_sport_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_lenz_sport_content=len(ekhtesasi_lenz_sport_content)
    ekhtesasi_lenz_sport_pivot=ekhtesasi_lenz_sport.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_lenz_sport_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_lenz_sport_popular_visit=ekhtesasi_lenz_sport_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_lenz_sport_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_lenz_sport_popular_duration=ekhtesasi_lenz_sport_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_lenz_sport_popular_visit = ekhtesasi_lenz_sport_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی لنز اسپرت', 'نام برنامه': 'محتواهای پربازدید اختصاصی لنز اسپرت'})
    ekhtesasi_lenz_sport_popular_duration = ekhtesasi_lenz_sport_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی لنز اسپرت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی لنز اسپرت'})
    
    print("ekhtesasi_lenz_sport_plus")
    ekhtesasi_lenz_sport_plus=ekhtesasi.query("channel == 'لنز اسپورت پلاس'")
    ekhtesasi_lenz_sport_plus_visit=ekhtesasi_lenz_sport_plus['تعداد بازدید'].sum()
    ekhtesasi_lenz_sport_plus_duration=ekhtesasi_lenz_sport_plus['مدت بازدید'].sum()
    ekhtesasi_lenz_sport_plus_duration=round(ekhtesasi_lenz_sport_plus_duration, 0)
    ekhtesasi_lenz_sport_plus_content=ekhtesasi_lenz_sport_plus.copy()
    ekhtesasi_lenz_sport_plus_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_lenz_sport_plus_content=len(ekhtesasi_lenz_sport_plus_content)
    ekhtesasi_lenz_sport_plus_pivot=ekhtesasi_lenz_sport_plus.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_lenz_sport_plus_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_lenz_sport_plus_popular_visit=ekhtesasi_lenz_sport_plus_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_lenz_sport_plus_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_lenz_sport_plus_popular_duration=ekhtesasi_lenz_sport_plus_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_lenz_sport_plus_popular_visit = ekhtesasi_lenz_sport_plus_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی لنز اسپرت پلاس', 'نام برنامه': 'محتواهای پربازدید اختصاصی لنز اسپرت پلاس'})
    ekhtesasi_lenz_sport_plus_popular_duration = ekhtesasi_lenz_sport_plus_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی لنز اسپرت پلاس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی لنز اسپرت پلاس'})
    
    print("ekhtesasi_tva_sport")
    ekhtesasi_tva_sport=ekhtesasi.query("channel == 'تیوا اسپورت'")
    ekhtesasi_tva_sport_visit=ekhtesasi_tva_sport['تعداد بازدید'].sum()
    ekhtesasi_tva_sport_duration=ekhtesasi_tva_sport['مدت بازدید'].sum()
    ekhtesasi_tva_sport_duration=round(ekhtesasi_tva_sport_duration, 0)
    ekhtesasi_tva_sport_content=ekhtesasi_tva_sport.copy()
    ekhtesasi_tva_sport_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_sport_content=len(ekhtesasi_tva_sport_content)
    ekhtesasi_tva_sport_pivot=ekhtesasi_tva_sport.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_sport_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_sport_popular_visit=ekhtesasi_tva_sport_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_sport_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_sport_popular_duration=ekhtesasi_tva_sport_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_sport_popular_visit = ekhtesasi_tva_sport_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا اسپرت', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا اسپرت'})
    ekhtesasi_tva_sport_popular_duration = ekhtesasi_tva_sport_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا اسپرت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا اسپرت'})
    
    print("ekhtesasi_tva_sport_two")
    ekhtesasi_tva_sport_two=ekhtesasi.query("channel == 'تیوا اسپورت دو'")
    ekhtesasi_tva_sport_two_visit=ekhtesasi_tva_sport_two['تعداد بازدید'].sum()
    ekhtesasi_tva_sport_two_duration=ekhtesasi_tva_sport_two['مدت بازدید'].sum()
    ekhtesasi_tva_sport_two_duration=round(ekhtesasi_tva_sport_two_duration, 0)
    ekhtesasi_tva_sport_two_content=ekhtesasi_tva_sport_two.copy()
    ekhtesasi_tva_sport_two_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_sport_two_content=len(ekhtesasi_tva_sport_two_content)
    ekhtesasi_tva_sport_two_pivot=ekhtesasi_tva_sport_two.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_sport_two_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_sport_two_popular_visit=ekhtesasi_tva_sport_two_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_sport_two_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_sport_two_popular_duration=ekhtesasi_tva_sport_two_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_sport_two_popular_visit = ekhtesasi_tva_sport_two_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا اسپرت دو', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا اسپرت دو'})
    ekhtesasi_tva_sport_two_popular_duration = ekhtesasi_tva_sport_two_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا اسپرت دو (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا اسپرت دو'})
    
    print("ekhtesasi_tva_boors")
    ekhtesasi_tva_boors=ekhtesasi.query("channel == 'تیوا بورس'")
    ekhtesasi_tva_boors_visit=ekhtesasi_tva_boors['تعداد بازدید'].sum()
    ekhtesasi_tva_boors_duration=ekhtesasi_tva_boors['مدت بازدید'].sum()
    ekhtesasi_tva_boors_duration=round(ekhtesasi_tva_boors_duration, 0)
    ekhtesasi_tva_boors_content=ekhtesasi_tva_boors.copy()
    ekhtesasi_tva_boors_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_boors_content=len(ekhtesasi_tva_boors_content)
    ekhtesasi_tva_boors_pivot=ekhtesasi_tva_boors.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_boors_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_boors_popular_visit=ekhtesasi_tva_boors_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_boors_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_boors_popular_duration=ekhtesasi_tva_boors_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_boors_popular_visit = ekhtesasi_tva_boors_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا بورس', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا بورس'})
    ekhtesasi_tva_boors_popular_duration = ekhtesasi_tva_boors_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا بورس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا بورس'})
    
    print("ekhtesasi_tva_two")
    ekhtesasi_tva_two=ekhtesasi.query("channel == 'تیوا دو'")
    ekhtesasi_tva_two_visit=ekhtesasi_tva_two['تعداد بازدید'].sum()
    ekhtesasi_tva_two_duration=ekhtesasi_tva_two['مدت بازدید'].sum()
    ekhtesasi_tva_two_duration=round(ekhtesasi_tva_two_duration, 0)
    ekhtesasi_tva_two_content=ekhtesasi_tva_two.copy()
    ekhtesasi_tva_two_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_two_content=len(ekhtesasi_tva_two_content)
    ekhtesasi_tva_two_pivot=ekhtesasi_tva_two.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_two_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_two_popular_visit=ekhtesasi_tva_two_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_two_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_two_popular_duration=ekhtesasi_tva_two_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_two_popular_visit = ekhtesasi_tva_two_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا دو', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا دو'})
    ekhtesasi_tva_two_popular_duration = ekhtesasi_tva_two_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا دو (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا دو'})
    
    print("ekhtesasi_tva_film")
    ekhtesasi_tva_film=ekhtesasi.query("channel == 'تیوا فیلم'")
    ekhtesasi_tva_film_visit=ekhtesasi_tva_film['تعداد بازدید'].sum()
    ekhtesasi_tva_film_duration=ekhtesasi_tva_film['مدت بازدید'].sum()
    ekhtesasi_tva_film_duration=round(ekhtesasi_tva_film_duration, 0)
    ekhtesasi_tva_film_content=ekhtesasi_tva_film.copy()
    ekhtesasi_tva_film_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_film_content=len(ekhtesasi_tva_film_content)
    ekhtesasi_tva_film_pivot=ekhtesasi_tva_film.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_film_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_film_popular_visit=ekhtesasi_tva_film_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_film_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_film_popular_duration=ekhtesasi_tva_film_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_film_popular_visit = ekhtesasi_tva_film_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا فیلم', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا فیلم'})
    ekhtesasi_tva_film_popular_duration = ekhtesasi_tva_film_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا فیلم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا فیلم'})
    
    print("ekhtesasi_tva_kodak")
    ekhtesasi_tva_kodak=ekhtesasi.query("channel == 'تیوا کودک'")
    ekhtesasi_tva_kodak_visit=ekhtesasi_tva_kodak['تعداد بازدید'].sum()
    ekhtesasi_tva_kodak_duration=ekhtesasi_tva_kodak['مدت بازدید'].sum()
    ekhtesasi_tva_kodak_duration=round(ekhtesasi_tva_kodak_duration, 0)
    ekhtesasi_tva_kodak_content=ekhtesasi_tva_kodak.copy()
    ekhtesasi_tva_kodak_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_kodak_content=len(ekhtesasi_tva_kodak_content)
    ekhtesasi_tva_kodak_pivot=ekhtesasi_tva_kodak.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_kodak_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_kodak_popular_visit=ekhtesasi_tva_kodak_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_kodak_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_kodak_popular_duration=ekhtesasi_tva_kodak_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_kodak_popular_visit = ekhtesasi_tva_kodak_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا کودک', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا کودک'})
    ekhtesasi_tva_kodak_popular_duration = ekhtesasi_tva_kodak_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا کودک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا کودک'})
    
    print("ekhtesasi_tva_nava")
    ekhtesasi_tva_nava=ekhtesasi.query("channel == 'تیوا نوا'")
    ekhtesasi_tva_nava_visit=ekhtesasi_tva_nava['تعداد بازدید'].sum()
    ekhtesasi_tva_nava_duration=ekhtesasi_tva_nava['مدت بازدید'].sum()
    ekhtesasi_tva_nava_duration=round(ekhtesasi_tva_nava_duration, 0)
    ekhtesasi_tva_nava_content=ekhtesasi_tva_nava.copy()
    ekhtesasi_tva_nava_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_nava_content=len(ekhtesasi_tva_nava_content)
    ekhtesasi_tva_nava_pivot=ekhtesasi_tva_nava.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_nava_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_nava_popular_visit=ekhtesasi_tva_nava_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_nava_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_nava_popular_duration=ekhtesasi_tva_nava_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_nava_popular_visit = ekhtesasi_tva_nava_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا نوا', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا نوا'})
    ekhtesasi_tva_nava_popular_duration = ekhtesasi_tva_nava_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا نوا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا نوا'})
    
    print("ekhtesasi_tva_one")
    ekhtesasi_tva_one=ekhtesasi.query("channel == 'تیوا یک'")
    ekhtesasi_tva_one_visit=ekhtesasi_tva_one['تعداد بازدید'].sum()
    ekhtesasi_tva_one_duration=ekhtesasi_tva_one['مدت بازدید'].sum()
    ekhtesasi_tva_one_duration=round(ekhtesasi_tva_one_duration, 0)
    ekhtesasi_tva_one_content=ekhtesasi_tva_one.copy()
    ekhtesasi_tva_one_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_one_content=len(ekhtesasi_tva_one_content)
    ekhtesasi_tva_one_pivot=ekhtesasi_tva_one.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_one_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_one_popular_visit=ekhtesasi_tva_one_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_one_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_one_popular_duration=ekhtesasi_tva_one_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_one_popular_visit = ekhtesasi_tva_one_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا یک', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا یک'})
    ekhtesasi_tva_one_popular_duration = ekhtesasi_tva_one_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا یک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا یک'})
    
    print("ekhtesasi_mahfel")
    ekhtesasi_mahfel=ekhtesasi.query("channel == 'محفل'")
    ekhtesasi_mahfel_visit=ekhtesasi_mahfel['تعداد بازدید'].sum()
    ekhtesasi_mahfel_duration=ekhtesasi_mahfel['مدت بازدید'].sum()
    ekhtesasi_mahfel_duration=round(ekhtesasi_mahfel_duration, 0)
    ekhtesasi_mahfel_content=ekhtesasi_mahfel.copy()
    ekhtesasi_mahfel_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_mahfel_content=len(ekhtesasi_mahfel_content)
    ekhtesasi_mahfel_pivot=ekhtesasi_mahfel.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_mahfel_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_mahfel_popular_visit=ekhtesasi_mahfel_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_mahfel_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_mahfel_popular_duration=ekhtesasi_mahfel_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_mahfel_popular_visit = ekhtesasi_mahfel_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی محفل', 'نام برنامه': 'محتواهای پربازدید اختصاصی محفل'})
    ekhtesasi_mahfel_popular_duration = ekhtesasi_mahfel_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی محفل (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی محفل'})
    
    print("ekhtesasi_tva_avand")
    ekhtesasi_tva_avand=ekhtesasi.query("channel == 'تیوا آوند'")
    ekhtesasi_tva_avand_visit=ekhtesasi_tva_avand['تعداد بازدید'].sum()
    ekhtesasi_tva_avand_duration=ekhtesasi_tva_avand['مدت بازدید'].sum()
    ekhtesasi_tva_avand_duration=round(ekhtesasi_tva_avand_duration, 0)
    ekhtesasi_tva_avand_content=ekhtesasi_tva_avand.copy()
    ekhtesasi_tva_avand_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_avand_content=len(ekhtesasi_tva_avand_content)
    ekhtesasi_tva_avand_pivot=ekhtesasi_tva_avand.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_avand_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_avand_popular_visit=ekhtesasi_tva_avand_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_avand_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_avand_popular_duration=ekhtesasi_tva_avand_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_avand_popular_visit = ekhtesasi_tva_avand_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا آوند', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا آوند'})
    ekhtesasi_tva_avand_popular_duration = ekhtesasi_tva_avand_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا آوند (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا آوند'})
    
    
    
    
    
    
    
    
    
    
    
    print("ekhtesasi_tva_sport_one")
    ekhtesasi_tva_sport_one=ekhtesasi.query("channel == 'تیوا اسپورت یک'")
    ekhtesasi_tva_sport_one_visit=ekhtesasi_tva_sport_one['تعداد بازدید'].sum()
    ekhtesasi_tva_sport_one_duration=ekhtesasi_tva_sport_one['مدت بازدید'].sum()
    ekhtesasi_tva_sport_one_duration=round(ekhtesasi_tva_sport_one_duration, 0)
    ekhtesasi_tva_sport_one_content=ekhtesasi_tva_sport_one.copy()
    ekhtesasi_tva_sport_one_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_tva_sport_one_content=len(ekhtesasi_tva_sport_one_content)
    ekhtesasi_tva_sport_one_pivot=ekhtesasi_tva_sport_one.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_tva_sport_one_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_sport_one_popular_visit=ekhtesasi_tva_sport_one_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_tva_sport_one_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_tva_sport_one_popular_duration=ekhtesasi_tva_sport_one_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_tva_sport_one_popular_visit = ekhtesasi_tva_sport_one_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی تیوا اسپورت یک', 'نام برنامه': 'محتواهای پربازدید اختصاصی تیوا اسپورت یک'})
    ekhtesasi_tva_sport_one_popular_duration = ekhtesasi_tva_sport_one_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی تیوا اسپورت یک (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی تیوا اسپورت یک'})
    
    print("ekhtesasi_mes_rafsanjan")
    ekhtesasi_mes_rafsanjan=ekhtesasi.query("channel == 'مس رفسنجان'")
    ekhtesasi_mes_rafsanjan_visit=ekhtesasi_mes_rafsanjan['تعداد بازدید'].sum()
    ekhtesasi_mes_rafsanjan_duration=ekhtesasi_mes_rafsanjan['مدت بازدید'].sum()
    ekhtesasi_mes_rafsanjan_duration=round(ekhtesasi_mes_rafsanjan_duration, 0)
    ekhtesasi_mes_rafsanjan_content=ekhtesasi_mes_rafsanjan.copy()
    ekhtesasi_mes_rafsanjan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_mes_rafsanjan_content=len(ekhtesasi_mes_rafsanjan_content)
    ekhtesasi_mes_rafsanjan_pivot=ekhtesasi_mes_rafsanjan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_mes_rafsanjan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_mes_rafsanjan_popular_visit=ekhtesasi_mes_rafsanjan_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_mes_rafsanjan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_mes_rafsanjan_popular_duration=ekhtesasi_mes_rafsanjan_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_mes_rafsanjan_popular_visit = ekhtesasi_mes_rafsanjan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی مس رفسنجان', 'نام برنامه': 'محتواهای پربازدید اختصاصی مس رفسنجان'})
    ekhtesasi_mes_rafsanjan_popular_duration = ekhtesasi_mes_rafsanjan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی مس رفسنجان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی مس رفسنجان'})
    
    print("ekhtesasi_emroz")
    ekhtesasi_emroz=ekhtesasi.query("channel == 'امروز'")
    ekhtesasi_emroz_visit=ekhtesasi_emroz['تعداد بازدید'].sum()
    ekhtesasi_emroz_duration=ekhtesasi_emroz['مدت بازدید'].sum()
    ekhtesasi_emroz_duration=round(ekhtesasi_emroz_duration, 0)
    ekhtesasi_emroz_content=ekhtesasi_emroz.copy()
    ekhtesasi_emroz_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_emroz_content=len(ekhtesasi_emroz_content)
    ekhtesasi_emroz_pivot=ekhtesasi_emroz.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_emroz_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_emroz_popular_visit=ekhtesasi_emroz_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_emroz_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_emroz_popular_duration=ekhtesasi_emroz_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_emroz_popular_visit = ekhtesasi_emroz_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی امروز', 'نام برنامه': 'محتواهای پربازدید اختصاصی امروز'})
    ekhtesasi_emroz_popular_duration = ekhtesasi_emroz_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی امروز (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی امروز'})
    
    print("ekhtesasi_keypad")
    ekhtesasi_keypad=ekhtesasi.query("channel == 'کیپاد'")
    ekhtesasi_keypad_visit=ekhtesasi_keypad['تعداد بازدید'].sum()
    ekhtesasi_keypad_duration=ekhtesasi_keypad['مدت بازدید'].sum()
    ekhtesasi_keypad_duration=round(ekhtesasi_keypad_duration, 0)
    ekhtesasi_keypad_content=ekhtesasi_keypad.copy()
    ekhtesasi_keypad_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_keypad_content=len(ekhtesasi_keypad_content)
    ekhtesasi_keypad_pivot=ekhtesasi_keypad.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_keypad_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_keypad_popular_visit=ekhtesasi_keypad_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_keypad_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_keypad_popular_duration=ekhtesasi_keypad_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_keypad_popular_visit = ekhtesasi_keypad_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی کیپاد', 'نام برنامه': 'محتواهای پربازدید اختصاصی کیپاد'})
    ekhtesasi_keypad_popular_duration = ekhtesasi_keypad_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی کیپاد (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی کیپاد'})
    
    print("ekhtesasi_dorfa")
    ekhtesasi_dorfa=ekhtesasi.query("channel == 'دُرفا'")
    ekhtesasi_dorfa_visit=ekhtesasi_dorfa['تعداد بازدید'].sum()
    ekhtesasi_dorfa_duration=ekhtesasi_dorfa['مدت بازدید'].sum()
    ekhtesasi_dorfa_duration=round(ekhtesasi_dorfa_duration, 0)
    ekhtesasi_dorfa_content=ekhtesasi_dorfa.copy()
    ekhtesasi_dorfa_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_dorfa_content=len(ekhtesasi_dorfa_content)
    ekhtesasi_dorfa_pivot=ekhtesasi_dorfa.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_dorfa_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_dorfa_popular_visit=ekhtesasi_dorfa_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_dorfa_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_dorfa_popular_duration=ekhtesasi_dorfa_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_dorfa_popular_visit = ekhtesasi_dorfa_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی دُرفا', 'نام برنامه': 'محتواهای پربازدید اختصاصی دُرفا'})
    ekhtesasi_dorfa_popular_duration = ekhtesasi_dorfa_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی دُرفا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی دُرفا'})
    
    print("ekhtesasi_lenz_film")
    ekhtesasi_lenz_film=ekhtesasi.query("channel == 'لنز فیلم'")
    ekhtesasi_lenz_film_visit=ekhtesasi_lenz_film['تعداد بازدید'].sum()
    ekhtesasi_lenz_film_duration=ekhtesasi_lenz_film['مدت بازدید'].sum()
    ekhtesasi_lenz_film_duration=round(ekhtesasi_lenz_film_duration, 0)
    ekhtesasi_lenz_film_content=ekhtesasi_lenz_film.copy()
    ekhtesasi_lenz_film_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_lenz_film_content=len(ekhtesasi_lenz_film_content)
    ekhtesasi_lenz_film_pivot=ekhtesasi_lenz_film.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_lenz_film_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_lenz_film_popular_visit=ekhtesasi_lenz_film_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_lenz_film_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_lenz_film_popular_duration=ekhtesasi_lenz_film_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_lenz_film_popular_visit = ekhtesasi_lenz_film_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی لنز فیلم', 'نام برنامه': 'محتواهای پربازدید اختصاصی لنز فیلم'})
    ekhtesasi_lenz_film_popular_duration = ekhtesasi_lenz_film_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی لنز فیلم (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی لنز فیلم'})
    
    print("ekhtesasi_eco_pars")
    ekhtesasi_eco_pars=ekhtesasi.query("channel == 'اکو پارس'")
    ekhtesasi_eco_pars_visit=ekhtesasi_eco_pars['تعداد بازدید'].sum()
    ekhtesasi_eco_pars_duration=ekhtesasi_eco_pars['مدت بازدید'].sum()
    ekhtesasi_eco_pars_duration=round(ekhtesasi_eco_pars_duration, 0)
    ekhtesasi_eco_pars_content=ekhtesasi_eco_pars.copy()
    ekhtesasi_eco_pars_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_eco_pars_content=len(ekhtesasi_eco_pars_content)
    ekhtesasi_eco_pars_pivot=ekhtesasi_eco_pars.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_eco_pars_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_eco_pars_popular_visit=ekhtesasi_eco_pars_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_eco_pars_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_eco_pars_popular_duration=ekhtesasi_eco_pars_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_eco_pars_popular_visit = ekhtesasi_eco_pars_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی اکو پارس', 'نام برنامه': 'محتواهای پربازدید اختصاصی اکو پارس'})
    ekhtesasi_eco_pars_popular_duration = ekhtesasi_eco_pars_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی اکو پارس (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی اکو پارس'})
    
    print("ekhtesasi_ara")
    ekhtesasi_ara=ekhtesasi.query("channel == 'آرا'")
    ekhtesasi_ara_visit=ekhtesasi_ara['تعداد بازدید'].sum()
    ekhtesasi_ara_duration=ekhtesasi_ara['مدت بازدید'].sum()
    ekhtesasi_ara_duration=round(ekhtesasi_ara_duration, 0)
    ekhtesasi_ara_content=ekhtesasi_ara.copy()
    ekhtesasi_ara_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_ara_content=len(ekhtesasi_ara_content)
    ekhtesasi_ara_pivot=ekhtesasi_ara.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_ara_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_ara_popular_visit=ekhtesasi_ara_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_ara_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_ara_popular_duration=ekhtesasi_ara_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_ara_popular_visit = ekhtesasi_ara_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی آرا', 'نام برنامه': 'محتواهای پربازدید اختصاصی آرا'})
    ekhtesasi_ara_popular_duration = ekhtesasi_ara_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی آرا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی آرا'})
    
    print("ekhtesasi_eshareh")
    ekhtesasi_eshareh=ekhtesasi.query("channel == 'اشاره'")
    ekhtesasi_eshareh_visit=ekhtesasi_eshareh['تعداد بازدید'].sum()
    ekhtesasi_eshareh_duration=ekhtesasi_eshareh['مدت بازدید'].sum()
    ekhtesasi_eshareh_duration=round(ekhtesasi_eshareh_duration, 0)
    ekhtesasi_eshareh_content=ekhtesasi_eshareh.copy()
    ekhtesasi_eshareh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_eshareh_content=len(ekhtesasi_eshareh_content)
    ekhtesasi_eshareh_pivot=ekhtesasi_eshareh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_eshareh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_eshareh_popular_visit=ekhtesasi_eshareh_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_eshareh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_eshareh_popular_duration=ekhtesasi_eshareh_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_eshareh_popular_visit = ekhtesasi_eshareh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی اشاره', 'نام برنامه': 'محتواهای پربازدید اختصاصی اشاره'})
    ekhtesasi_eshareh_popular_duration = ekhtesasi_eshareh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی اشاره (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی اشاره'})
    
    print("ekhtesasi_astan_ghods_razavi")
    ekhtesasi_astan_ghods_razavi=ekhtesasi.query("channel == 'آستان قدس رضوی'")
    ekhtesasi_astan_ghods_razavi_visit=ekhtesasi_astan_ghods_razavi['تعداد بازدید'].sum()
    ekhtesasi_astan_ghods_razavi_duration=ekhtesasi_astan_ghods_razavi['مدت بازدید'].sum()
    ekhtesasi_astan_ghods_razavi_duration=round(ekhtesasi_astan_ghods_razavi_duration, 0)
    ekhtesasi_astan_ghods_razavi_content=ekhtesasi_astan_ghods_razavi.copy()
    ekhtesasi_astan_ghods_razavi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_astan_ghods_razavi_content=len(ekhtesasi_astan_ghods_razavi_content)
    ekhtesasi_astan_ghods_razavi_pivot=ekhtesasi_astan_ghods_razavi.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_astan_ghods_razavi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_astan_ghods_razavi_popular_visit=ekhtesasi_astan_ghods_razavi_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_astan_ghods_razavi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_astan_ghods_razavi_popular_duration=ekhtesasi_astan_ghods_razavi_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_astan_ghods_razavi_popular_visit = ekhtesasi_astan_ghods_razavi_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی آستان قدس رضوی', 'نام برنامه': 'محتواهای پربازدید اختصاصی آستان قدس رضوی'})
    ekhtesasi_astan_ghods_razavi_popular_duration = ekhtesasi_astan_ghods_razavi_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی آستان قدس رضوی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی آستان قدس رضوی'})
    
    print("ekhtesasi_aio_jahanbin")
    ekhtesasi_aio_jahanbin=ekhtesasi.query("channel == 'آیو جهان‌بین'")
    ekhtesasi_aio_jahanbin_visit=ekhtesasi_aio_jahanbin['تعداد بازدید'].sum()
    ekhtesasi_aio_jahanbin_duration=ekhtesasi_aio_jahanbin['مدت بازدید'].sum()
    ekhtesasi_aio_jahanbin_duration=round(ekhtesasi_aio_jahanbin_duration, 0)
    ekhtesasi_aio_jahanbin_content=ekhtesasi_aio_jahanbin.copy()
    ekhtesasi_aio_jahanbin_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_aio_jahanbin_content=len(ekhtesasi_aio_jahanbin_content)
    ekhtesasi_aio_jahanbin_pivot=ekhtesasi_aio_jahanbin.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_aio_jahanbin_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_aio_jahanbin_popular_visit=ekhtesasi_aio_jahanbin_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_aio_jahanbin_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_aio_jahanbin_popular_duration=ekhtesasi_aio_jahanbin_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_aio_jahanbin_popular_visit = ekhtesasi_aio_jahanbin_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی آیو جهان‌بین', 'نام برنامه': 'محتواهای پربازدید اختصاصی آیو جهان‌بین'})
    ekhtesasi_aio_jahanbin_popular_duration = ekhtesasi_aio_jahanbin_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی آیو جهان‌بین (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی آیو جهان‌بین'})
    
    print("ekhtesasi_sepahan_tv")
    ekhtesasi_sepahan_tv=ekhtesasi.query("channel == 'سپاهان‌ تی‌وی'")
    ekhtesasi_sepahan_tv_visit=ekhtesasi_sepahan_tv['تعداد بازدید'].sum()
    ekhtesasi_sepahan_tv_duration=ekhtesasi_sepahan_tv['مدت بازدید'].sum()
    ekhtesasi_sepahan_tv_duration=round(ekhtesasi_sepahan_tv_duration, 0)
    ekhtesasi_sepahan_tv_content=ekhtesasi_sepahan_tv.copy()
    ekhtesasi_sepahan_tv_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_sepahan_tv_content=len(ekhtesasi_sepahan_tv_content)
    ekhtesasi_sepahan_tv_pivot=ekhtesasi_sepahan_tv.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_sepahan_tv_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_sepahan_tv_popular_visit=ekhtesasi_sepahan_tv_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_sepahan_tv_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_sepahan_tv_popular_duration=ekhtesasi_sepahan_tv_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_sepahan_tv_popular_visit = ekhtesasi_sepahan_tv_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی سپاهان‌ تی‌وی', 'نام برنامه': 'محتواهای پربازدید اختصاصی سپاهان‌ تی‌وی'})
    ekhtesasi_sepahan_tv_popular_duration = ekhtesasi_sepahan_tv_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی سپاهان‌ تی‌وی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی سپاهان‌ تی‌وی'})
    
    print("ekhtesasi_borsan")
    ekhtesasi_borsan=ekhtesasi.query("channel == 'بورسان'")
    ekhtesasi_borsan_visit=ekhtesasi_borsan['تعداد بازدید'].sum()
    ekhtesasi_borsan_duration=ekhtesasi_borsan['مدت بازدید'].sum()
    ekhtesasi_borsan_duration=round(ekhtesasi_borsan_duration, 0)
    ekhtesasi_borsan_content=ekhtesasi_borsan.copy()
    ekhtesasi_borsan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_borsan_content=len(ekhtesasi_borsan_content)
    ekhtesasi_borsan_pivot=ekhtesasi_borsan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_borsan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_borsan_popular_visit=ekhtesasi_borsan_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_borsan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_borsan_popular_duration=ekhtesasi_borsan_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_borsan_popular_visit = ekhtesasi_borsan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی بورسان', 'نام برنامه': 'محتواهای پربازدید اختصاصی بورسان'})
    ekhtesasi_borsan_popular_duration = ekhtesasi_borsan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی بورسان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی بورسان'})
    
    print("ekhtesasi_haram_razavi")
    ekhtesasi_haram_razavi=ekhtesasi.query("channel == 'حرم رضوی'")
    ekhtesasi_haram_razavi_visit=ekhtesasi_haram_razavi['تعداد بازدید'].sum()
    ekhtesasi_haram_razavi_duration=ekhtesasi_haram_razavi['مدت بازدید'].sum()
    ekhtesasi_haram_razavi_duration=round(ekhtesasi_haram_razavi_duration, 0)
    ekhtesasi_haram_razavi_content=ekhtesasi_haram_razavi.copy()
    ekhtesasi_haram_razavi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_haram_razavi_content=len(ekhtesasi_haram_razavi_content)
    ekhtesasi_haram_razavi_pivot=ekhtesasi_haram_razavi.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_haram_razavi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_haram_razavi_popular_visit=ekhtesasi_haram_razavi_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_haram_razavi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_haram_razavi_popular_duration=ekhtesasi_haram_razavi_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_haram_razavi_popular_visit = ekhtesasi_haram_razavi_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی حرم رضوی', 'نام برنامه': 'محتواهای پربازدید اختصاصی حرم رضوی'})
    ekhtesasi_haram_razavi_popular_duration = ekhtesasi_haram_razavi_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی حرم رضوی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی حرم رضوی'})
    
    print("ekhtesasi_javaneh")
    ekhtesasi_javaneh=ekhtesasi.query("channel == 'جوانه'")
    ekhtesasi_javaneh_visit=ekhtesasi_javaneh['تعداد بازدید'].sum()
    ekhtesasi_javaneh_duration=ekhtesasi_javaneh['مدت بازدید'].sum()
    ekhtesasi_javaneh_duration=round(ekhtesasi_javaneh_duration, 0)
    ekhtesasi_javaneh_content=ekhtesasi_javaneh.copy()
    ekhtesasi_javaneh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_javaneh_content=len(ekhtesasi_javaneh_content)
    ekhtesasi_javaneh_pivot=ekhtesasi_javaneh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_javaneh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_javaneh_popular_visit=ekhtesasi_javaneh_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_javaneh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_javaneh_popular_duration=ekhtesasi_javaneh_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_javaneh_popular_visit = ekhtesasi_javaneh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی جوانه', 'نام برنامه': 'محتواهای پربازدید اختصاصی جوانه'})
    ekhtesasi_javaneh_popular_duration = ekhtesasi_javaneh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی جوانه (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی جوانه'})
    
    print("ekhtesasi_jam")
    ekhtesasi_jam=ekhtesasi.query("channel == 'جام'")
    ekhtesasi_jam_visit=ekhtesasi_jam['تعداد بازدید'].sum()
    ekhtesasi_jam_duration=ekhtesasi_jam['مدت بازدید'].sum()
    ekhtesasi_jam_duration=round(ekhtesasi_jam_duration, 0)
    ekhtesasi_jam_content=ekhtesasi_jam.copy()
    ekhtesasi_jam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_jam_content=len(ekhtesasi_jam_content)
    ekhtesasi_jam_pivot=ekhtesasi_jam.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_jam_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_jam_popular_visit=ekhtesasi_jam_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_jam_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_jam_popular_duration=ekhtesasi_jam_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_jam_popular_visit = ekhtesasi_jam_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی جام', 'نام برنامه': 'محتواهای پربازدید اختصاصی جام'})
    ekhtesasi_jam_popular_duration = ekhtesasi_jam_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی جام (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی جام'})
    
    print("ekhtesasi_habib")
    ekhtesasi_habib=ekhtesasi.query("channel == 'حبیب'")
    ekhtesasi_habib_visit=ekhtesasi_habib['تعداد بازدید'].sum()
    ekhtesasi_habib_duration=ekhtesasi_habib['مدت بازدید'].sum()
    ekhtesasi_habib_duration=round(ekhtesasi_habib_duration, 0)
    ekhtesasi_habib_content=ekhtesasi_habib.copy()
    ekhtesasi_habib_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_habib_content=len(ekhtesasi_habib_content)
    ekhtesasi_habib_pivot=ekhtesasi_habib.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_habib_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_habib_popular_visit=ekhtesasi_habib_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_habib_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_habib_popular_duration=ekhtesasi_habib_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_habib_popular_visit = ekhtesasi_habib_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی حبیب', 'نام برنامه': 'محتواهای پربازدید اختصاصی حبیب'})
    ekhtesasi_habib_popular_duration = ekhtesasi_habib_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی حبیب (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی حبیب'})
    
    print("ekhtesasi_rahro")
    ekhtesasi_rahro=ekhtesasi.query("channel == 'رهرو'")
    ekhtesasi_rahro_visit=ekhtesasi_rahro['تعداد بازدید'].sum()
    ekhtesasi_rahro_duration=ekhtesasi_rahro['مدت بازدید'].sum()
    ekhtesasi_rahro_duration=round(ekhtesasi_rahro_duration, 0)
    ekhtesasi_rahro_content=ekhtesasi_rahro.copy()
    ekhtesasi_rahro_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_rahro_content=len(ekhtesasi_rahro_content)
    ekhtesasi_rahro_pivot=ekhtesasi_rahro.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_rahro_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_rahro_popular_visit=ekhtesasi_rahro_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_rahro_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_rahro_popular_duration=ekhtesasi_rahro_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_rahro_popular_visit = ekhtesasi_rahro_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی رهرو', 'نام برنامه': 'محتواهای پربازدید اختصاصی رهرو'})
    ekhtesasi_rahro_popular_duration = ekhtesasi_rahro_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی رهرو (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی رهرو'})
    
    print("ekhtesasi_iran_economy")
    ekhtesasi_iran_economy=ekhtesasi.query("channel == 'ایران اکونومی'")
    ekhtesasi_iran_economy_visit=ekhtesasi_iran_economy['تعداد بازدید'].sum()
    ekhtesasi_iran_economy_duration=ekhtesasi_iran_economy['مدت بازدید'].sum()
    ekhtesasi_iran_economy_duration=round(ekhtesasi_iran_economy_duration, 0)
    ekhtesasi_iran_economy_content=ekhtesasi_iran_economy.copy()
    ekhtesasi_iran_economy_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_iran_economy_content=len(ekhtesasi_iran_economy_content)
    ekhtesasi_iran_economy_pivot=ekhtesasi_iran_economy.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_iran_economy_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_iran_economy_popular_visit=ekhtesasi_iran_economy_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_iran_economy_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_iran_economy_popular_duration=ekhtesasi_iran_economy_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_iran_economy_popular_visit = ekhtesasi_iran_economy_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی ایران اکونومی', 'نام برنامه': 'محتواهای پربازدید اختصاصی ایران اکونومی'})
    ekhtesasi_iran_economy_popular_duration = ekhtesasi_iran_economy_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی ایران اکونومی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی ایران اکونومی'})
    
    print("ekhtesasi_nama")
    ekhtesasi_nama=ekhtesasi.query("channel == 'نما'")
    ekhtesasi_nama_visit=ekhtesasi_nama['تعداد بازدید'].sum()
    ekhtesasi_nama_duration=ekhtesasi_nama['مدت بازدید'].sum()
    ekhtesasi_nama_duration=round(ekhtesasi_nama_duration, 0)
    ekhtesasi_nama_content=ekhtesasi_nama.copy()
    ekhtesasi_nama_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_nama_content=len(ekhtesasi_nama_content)
    ekhtesasi_nama_pivot=ekhtesasi_nama.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_nama_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_nama_popular_visit=ekhtesasi_nama_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_nama_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_nama_popular_duration=ekhtesasi_nama_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_nama_popular_visit = ekhtesasi_nama_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی نما', 'نام برنامه': 'محتواهای پربازدید اختصاصی نما'})
    ekhtesasi_nama_popular_duration = ekhtesasi_nama_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی نما (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی نما'})
    
    print("ekhtesasi_aio_sport")
    ekhtesasi_aio_sport=ekhtesasi.query("channel == 'آیواسپرت'")
    ekhtesasi_aio_sport_visit=ekhtesasi_aio_sport['تعداد بازدید'].sum()
    ekhtesasi_aio_sport_duration=ekhtesasi_aio_sport['مدت بازدید'].sum()
    ekhtesasi_aio_sport_duration=round(ekhtesasi_aio_sport_duration, 0)
    ekhtesasi_aio_sport_content=ekhtesasi_aio_sport.copy()
    ekhtesasi_aio_sport_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    ekhtesasi_aio_sport_content=len(ekhtesasi_aio_sport_content)
    ekhtesasi_aio_sport_pivot=ekhtesasi_aio_sport.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    ekhtesasi_aio_sport_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_aio_sport_popular_visit=ekhtesasi_aio_sport_pivot.iloc[0:10 , [0, 3]]
    ekhtesasi_aio_sport_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_aio_sport_popular_duration=ekhtesasi_aio_sport_pivot.iloc[0:10 , [0, 5]]
    
    ekhtesasi_aio_sport_popular_visit = ekhtesasi_aio_sport_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید اختصاصی آیواسپرت', 'نام برنامه': 'محتواهای پربازدید اختصاصی آیواسپرت'})
    ekhtesasi_aio_sport_popular_duration = ekhtesasi_aio_sport_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید اختصاصی آیواسپرت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید اختصاصی آیواسپرت'})
    
#    print("ekhtesasi_baghimande")
#    total_ekhtesasi_malom=pd.DataFrame()
#    total_ekhtesasi_majhol=pd.DataFrame()
#           
#    total_ekhtesasi_malom=pd.concat([ekhtesasi_mahfel,
#                                     ekhtesasi_tva_one,
#                                     ekhtesasi_tva_nava,
#                                     ekhtesasi_tva_kodak,
#                                     ekhtesasi_tva_film,
#                                     ekhtesasi_tva_two,
#                                     ekhtesasi_tva_boors,
#                                     ekhtesasi_tva_sport_two,
#                                     ekhtesasi_tva_sport,
#                                     ekhtesasi_lenz_sport_plus,
#                                     ekhtesasi_lenz_sport,
#                                     ekhtesasi_kodak_digiton,
#                                     ekhtesasi_shetab,
#                                     ekhtesasi_shaparak,
#                                     ekhtesasi_esteghlal,
#                                     ekhtesasi_perspolis,
#                                     ekhtesasi_tva_avand,], axis=0)
#    
#    total_ekhtesasi_majhol=pd.concat([ekhtesasi, total_ekhtesasi_malom]).drop_duplicates(keep=False)
    
    print("dataframe ekhtesasi channels")
    ekhtesasi_channels_statistics={'channel_name': ['اختصاصی محفل',
                                                    'اختصاصی تیوا یک',
                                                    'اختصاصی تیوا نوا',
                                                    'اختصاصی تیوا کودک',
                                         'اختصاصی تیوا فیلم',
                                         'اختصاصی تیوا دو',
                                         'اختصاصی تیوا بورس',
                                         'اختصاصی تیوا اسپرت دو',
                                         'اختصاصی تیوا اسپرت',
                                         'اختصاصی لنز اسپرت پلاس',
                                         'اختصاصی لنز اسپرت',
                                         'اختصاصی کودک دیجیتون',
                                         'اختصاصی شتاب',
                                         'اختصاصی شاپرک',
                                         'اختصاصی استقلال',
                                         'اختصاصی پرسپولیس',
                                         'اختصاصی تیوا آوند',
                                         'اختصاصی تیوا اسپرت یک',
                                         'اختصاصی مس رفسنجان',
                                         'اختصاصی امروز',
                                         'اختصاصی کیپاد',
                                         'اختصاصی دُرفا',
                                         'اختصاصی لنز فیلم',
                                         'اختصاصی اکو فارس',
                                         'اختصاصی آرا',
                                          'اختصاصی اشاره',
                                           'اختصاصی آستان قدس رضوی',
                                            'اختصاصی آیو جهان بین',
                                             'اختصاصی سپاهان تی وی',
                                              'اختصاصی بورسان',
                                               'اختصاصی حرم رضوی',
                                                'اختصاصی جوانه',
                                                 'اختصاصی جام',
                                                  'اختصاصی حبیب',
                                                   'اختصاصی رهرو',
                                                    'اختصاصی ایران اکونومی',
                                                     'اختصاصی نما',
                                                      'اختصاصی آیو اسپرت',],
           'channel_content': [ekhtesasi_mahfel_content, ekhtesasi_tva_one_content, ekhtesasi_tva_nava_content, ekhtesasi_tva_kodak_content, 
                               ekhtesasi_tva_film_content, ekhtesasi_tva_two_content, ekhtesasi_tva_boors_content, ekhtesasi_tva_sport_two_content, 
                               ekhtesasi_tva_sport_content, ekhtesasi_lenz_sport_plus_content, ekhtesasi_lenz_sport_content, ekhtesasi_kodak_digiton_content,
                               ekhtesasi_shetab_content, ekhtesasi_shaparak_content, ekhtesasi_esteghlal_content, ekhtesasi_perspolis_content,
                               ekhtesasi_tva_avand_content,
                               ekhtesasi_tva_sport_one_content,ekhtesasi_mes_rafsanjan_content,ekhtesasi_emroz_content,ekhtesasi_keypad_content,
                               ekhtesasi_dorfa_content,ekhtesasi_lenz_film_content,ekhtesasi_eco_pars_content,ekhtesasi_ara_content,
                               ekhtesasi_eshareh_content,ekhtesasi_astan_ghods_razavi_content,ekhtesasi_aio_jahanbin_content,ekhtesasi_sepahan_tv_content,
                               ekhtesasi_borsan_content,ekhtesasi_haram_razavi_content,ekhtesasi_javaneh_content,ekhtesasi_jam_content,
                               ekhtesasi_habib_content,ekhtesasi_rahro_content,ekhtesasi_iran_economy_content,ekhtesasi_nama_content,
                               ekhtesasi_aio_sport_content,],
           'channel_visit': [ekhtesasi_mahfel_visit, ekhtesasi_tva_one_visit, ekhtesasi_tva_nava_visit, ekhtesasi_tva_kodak_visit, 
                               ekhtesasi_tva_film_visit, ekhtesasi_tva_two_visit, ekhtesasi_tva_boors_visit, ekhtesasi_tva_sport_two_visit, 
                               ekhtesasi_tva_sport_visit, ekhtesasi_lenz_sport_plus_visit, ekhtesasi_lenz_sport_visit, ekhtesasi_kodak_digiton_visit,
                               ekhtesasi_shetab_visit, ekhtesasi_shaparak_visit, ekhtesasi_esteghlal_visit, ekhtesasi_perspolis_visit,
                               ekhtesasi_tva_avand_visit,
                               ekhtesasi_tva_sport_one_visit,ekhtesasi_mes_rafsanjan_visit,ekhtesasi_emroz_visit,ekhtesasi_keypad_visit,
                               ekhtesasi_dorfa_visit,ekhtesasi_lenz_film_visit,ekhtesasi_eco_pars_visit,ekhtesasi_ara_visit,
                               ekhtesasi_eshareh_visit,ekhtesasi_astan_ghods_razavi_visit,ekhtesasi_aio_jahanbin_visit,ekhtesasi_sepahan_tv_visit,
                               ekhtesasi_borsan_visit,ekhtesasi_haram_razavi_visit,ekhtesasi_javaneh_visit,ekhtesasi_jam_visit,
                               ekhtesasi_habib_visit,ekhtesasi_rahro_visit,ekhtesasi_iran_economy_visit,ekhtesasi_nama_visit,
                               ekhtesasi_aio_sport_visit,],
           'channel_duration': [ekhtesasi_mahfel_duration, ekhtesasi_tva_one_duration, ekhtesasi_tva_nava_duration, ekhtesasi_tva_kodak_duration, 
                               ekhtesasi_tva_film_duration, ekhtesasi_tva_two_duration, ekhtesasi_tva_boors_duration, ekhtesasi_tva_sport_two_duration, 
                               ekhtesasi_tva_sport_duration, ekhtesasi_lenz_sport_plus_duration, ekhtesasi_lenz_sport_duration, ekhtesasi_kodak_digiton_duration,
                               ekhtesasi_shetab_duration, ekhtesasi_shaparak_duration, ekhtesasi_esteghlal_duration, ekhtesasi_perspolis_duration,
                               ekhtesasi_tva_avand_duration,
                               ekhtesasi_tva_sport_one_duration,ekhtesasi_mes_rafsanjan_duration,ekhtesasi_emroz_duration,ekhtesasi_keypad_duration,
                               ekhtesasi_dorfa_duration,ekhtesasi_lenz_film_duration,ekhtesasi_eco_pars_duration,ekhtesasi_ara_duration,
                               ekhtesasi_eshareh_duration,ekhtesasi_astan_ghods_razavi_duration,ekhtesasi_aio_jahanbin_duration,ekhtesasi_sepahan_tv_duration,
                               ekhtesasi_borsan_duration,ekhtesasi_haram_razavi_duration,ekhtesasi_javaneh_duration,ekhtesasi_jam_duration,
                               ekhtesasi_habib_duration,ekhtesasi_rahro_duration,ekhtesasi_iran_economy_duration,ekhtesasi_nama_duration,
                               ekhtesasi_aio_sport_duration,],}
    ekhtesasi_channels_statistics=pd.DataFrame(ekhtesasi_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
    ekhtesasi_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
    ekhtesasi_channels_statistics=ekhtesasi_channels_statistics.rename(columns={'channel_name': 'نام شبکه', 'channel_content': 'تعداد محتوا', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})
    
    ekhtesasi_mahfel_popular_visit = ekhtesasi_mahfel_popular_visit.reset_index()
    del ekhtesasi_mahfel_popular_visit['index']
    ekhtesasi_mahfel_popular_duration = ekhtesasi_mahfel_popular_duration.reset_index()
    del ekhtesasi_mahfel_popular_duration['index']
    
    ekhtesasi_tva_one_popular_visit = ekhtesasi_tva_one_popular_visit.reset_index()
    del ekhtesasi_tva_one_popular_visit['index']
    ekhtesasi_tva_one_popular_duration = ekhtesasi_tva_one_popular_duration.reset_index()
    del ekhtesasi_tva_one_popular_duration['index']
    
    ekhtesasi_tva_nava_popular_visit = ekhtesasi_tva_nava_popular_visit.reset_index()
    del ekhtesasi_tva_nava_popular_visit['index']
    ekhtesasi_tva_nava_popular_duration = ekhtesasi_tva_nava_popular_duration.reset_index()
    del ekhtesasi_tva_nava_popular_duration['index']
    
    ekhtesasi_tva_kodak_popular_visit = ekhtesasi_tva_kodak_popular_visit.reset_index()
    del ekhtesasi_tva_kodak_popular_visit['index']
    ekhtesasi_tva_kodak_popular_duration = ekhtesasi_tva_kodak_popular_duration.reset_index()
    del ekhtesasi_tva_kodak_popular_duration['index']
    
    ekhtesasi_tva_film_popular_visit = ekhtesasi_tva_film_popular_visit.reset_index()
    del ekhtesasi_tva_film_popular_visit['index']
    ekhtesasi_tva_film_popular_duration = ekhtesasi_tva_film_popular_duration.reset_index()
    del ekhtesasi_tva_film_popular_duration['index']
    
    ekhtesasi_tva_two_popular_visit = ekhtesasi_tva_two_popular_visit.reset_index()
    del ekhtesasi_tva_two_popular_visit['index']
    ekhtesasi_tva_two_popular_duration = ekhtesasi_tva_two_popular_duration.reset_index()
    del ekhtesasi_tva_two_popular_duration['index']
    
    ekhtesasi_tva_boors_popular_visit = ekhtesasi_tva_boors_popular_visit.reset_index()
    del ekhtesasi_tva_boors_popular_visit['index']
    ekhtesasi_tva_boors_popular_duration = ekhtesasi_tva_boors_popular_duration.reset_index()
    del ekhtesasi_tva_boors_popular_duration['index']
    
    ekhtesasi_tva_sport_two_popular_visit = ekhtesasi_tva_sport_two_popular_visit.reset_index()
    del ekhtesasi_tva_sport_two_popular_visit['index']
    ekhtesasi_tva_sport_two_popular_duration = ekhtesasi_tva_sport_two_popular_duration.reset_index()
    del ekhtesasi_tva_sport_two_popular_duration['index']
    
    ekhtesasi_tva_sport_popular_visit = ekhtesasi_tva_sport_popular_visit.reset_index()
    del ekhtesasi_tva_sport_popular_visit['index']
    ekhtesasi_tva_sport_popular_duration = ekhtesasi_tva_sport_popular_duration.reset_index()
    del ekhtesasi_tva_sport_popular_duration['index']
    
    ekhtesasi_lenz_sport_plus_popular_visit = ekhtesasi_lenz_sport_plus_popular_visit.reset_index()
    del ekhtesasi_lenz_sport_plus_popular_visit['index']
    ekhtesasi_lenz_sport_plus_popular_duration = ekhtesasi_lenz_sport_plus_popular_duration.reset_index()
    del ekhtesasi_lenz_sport_plus_popular_duration['index']
    
    ekhtesasi_lenz_sport_popular_visit = ekhtesasi_lenz_sport_popular_visit.reset_index()
    del ekhtesasi_lenz_sport_popular_visit['index']
    ekhtesasi_lenz_sport_popular_duration = ekhtesasi_lenz_sport_popular_duration.reset_index()
    del ekhtesasi_lenz_sport_popular_duration['index']
    
    ekhtesasi_kodak_digiton_popular_visit = ekhtesasi_kodak_digiton_popular_visit.reset_index()
    del ekhtesasi_kodak_digiton_popular_visit['index']
    ekhtesasi_kodak_digiton_popular_duration = ekhtesasi_kodak_digiton_popular_duration.reset_index()
    del ekhtesasi_kodak_digiton_popular_duration['index']
    
    ekhtesasi_shetab_popular_visit = ekhtesasi_shetab_popular_visit.reset_index()
    del ekhtesasi_shetab_popular_visit['index']
    ekhtesasi_shetab_popular_duration = ekhtesasi_shetab_popular_duration.reset_index()
    del ekhtesasi_shetab_popular_duration['index']
    
    ekhtesasi_shaparak_popular_visit = ekhtesasi_shaparak_popular_visit.reset_index()
    del ekhtesasi_shaparak_popular_visit['index']
    ekhtesasi_shaparak_popular_duration = ekhtesasi_shaparak_popular_duration.reset_index()
    del ekhtesasi_shaparak_popular_duration['index']
    
    ekhtesasi_esteghlal_popular_visit = ekhtesasi_esteghlal_popular_visit.reset_index()
    del ekhtesasi_esteghlal_popular_visit['index']
    ekhtesasi_esteghlal_popular_duration = ekhtesasi_esteghlal_popular_duration.reset_index()
    del ekhtesasi_esteghlal_popular_duration['index']
    
    ekhtesasi_perspolis_popular_visit = ekhtesasi_perspolis_popular_visit.reset_index()
    del ekhtesasi_perspolis_popular_visit['index']
    ekhtesasi_perspolis_popular_duration = ekhtesasi_perspolis_popular_duration.reset_index()
    del ekhtesasi_perspolis_popular_duration['index']
    
    ekhtesasi_tva_avand_popular_visit = ekhtesasi_tva_avand_popular_visit.reset_index()
    del ekhtesasi_tva_avand_popular_visit['index']
    ekhtesasi_tva_avand_popular_duration = ekhtesasi_tva_avand_popular_duration.reset_index()
    del ekhtesasi_tva_avand_popular_duration['index']
    
    ekhtesasi_tva_sport_one_popular_visit = ekhtesasi_tva_sport_one_popular_visit.reset_index()
    del ekhtesasi_tva_sport_one_popular_visit['index']
    ekhtesasi_tva_sport_one_popular_duration = ekhtesasi_tva_sport_one_popular_duration.reset_index()
    del ekhtesasi_tva_sport_one_popular_duration['index']
    
    ekhtesasi_mes_rafsanjan_popular_visit = ekhtesasi_mes_rafsanjan_popular_visit.reset_index()
    del ekhtesasi_mes_rafsanjan_popular_visit['index']
    ekhtesasi_mes_rafsanjan_popular_duration = ekhtesasi_mes_rafsanjan_popular_duration.reset_index()
    del ekhtesasi_mes_rafsanjan_popular_duration['index']
    
    ekhtesasi_emroz_popular_visit = ekhtesasi_emroz_popular_visit.reset_index()
    del ekhtesasi_emroz_popular_visit['index']
    ekhtesasi_emroz_popular_duration = ekhtesasi_emroz_popular_duration.reset_index()
    del ekhtesasi_emroz_popular_duration['index']
    
    ekhtesasi_keypad_popular_visit = ekhtesasi_keypad_popular_visit.reset_index()
    del ekhtesasi_keypad_popular_visit['index']
    ekhtesasi_keypad_popular_duration = ekhtesasi_keypad_popular_duration.reset_index()
    del ekhtesasi_keypad_popular_duration['index']
    
    ekhtesasi_dorfa_popular_visit = ekhtesasi_dorfa_popular_visit.reset_index()
    del ekhtesasi_dorfa_popular_visit['index']
    ekhtesasi_dorfa_popular_duration = ekhtesasi_dorfa_popular_duration.reset_index()
    del ekhtesasi_dorfa_popular_duration['index']
    
    ekhtesasi_lenz_film_popular_visit = ekhtesasi_lenz_film_popular_visit.reset_index()
    del ekhtesasi_lenz_film_popular_visit['index']
    ekhtesasi_lenz_film_popular_duration = ekhtesasi_lenz_film_popular_duration.reset_index()
    del ekhtesasi_lenz_film_popular_duration['index']
    
    ekhtesasi_eco_pars_popular_visit = ekhtesasi_eco_pars_popular_visit.reset_index()
    del ekhtesasi_eco_pars_popular_visit['index']
    ekhtesasi_eco_pars_popular_duration = ekhtesasi_eco_pars_popular_duration.reset_index()
    del ekhtesasi_eco_pars_popular_duration['index']
    
    ekhtesasi_ara_popular_visit = ekhtesasi_ara_popular_visit.reset_index()
    del ekhtesasi_ara_popular_visit['index']
    ekhtesasi_ara_popular_duration = ekhtesasi_ara_popular_duration.reset_index()
    del ekhtesasi_ara_popular_duration['index']
    
    ekhtesasi_eshareh_popular_visit = ekhtesasi_eshareh_popular_visit.reset_index()
    del ekhtesasi_eshareh_popular_visit['index']
    ekhtesasi_eshareh_popular_duration = ekhtesasi_eshareh_popular_duration.reset_index()
    del ekhtesasi_eshareh_popular_duration['index']
    
    ekhtesasi_astan_ghods_razavi_popular_visit = ekhtesasi_astan_ghods_razavi_popular_visit.reset_index()
    del ekhtesasi_astan_ghods_razavi_popular_visit['index']
    ekhtesasi_astan_ghods_razavi_popular_duration = ekhtesasi_astan_ghods_razavi_popular_duration.reset_index()
    del ekhtesasi_astan_ghods_razavi_popular_duration['index']
    
    ekhtesasi_aio_jahanbin_popular_visit = ekhtesasi_aio_jahanbin_popular_visit.reset_index()
    del ekhtesasi_aio_jahanbin_popular_visit['index']
    ekhtesasi_aio_jahanbin_popular_duration = ekhtesasi_aio_jahanbin_popular_duration.reset_index()
    del ekhtesasi_aio_jahanbin_popular_duration['index']
    
    ekhtesasi_sepahan_tv_popular_visit = ekhtesasi_sepahan_tv_popular_visit.reset_index()
    del ekhtesasi_sepahan_tv_popular_visit['index']
    ekhtesasi_sepahan_tv_popular_duration = ekhtesasi_sepahan_tv_popular_duration.reset_index()
    del ekhtesasi_sepahan_tv_popular_duration['index']
    
    ekhtesasi_borsan_popular_visit = ekhtesasi_borsan_popular_visit.reset_index()
    del ekhtesasi_borsan_popular_visit['index']
    ekhtesasi_borsan_popular_duration = ekhtesasi_borsan_popular_duration.reset_index()
    del ekhtesasi_borsan_popular_duration['index']
    
    ekhtesasi_haram_razavi_popular_visit = ekhtesasi_haram_razavi_popular_visit.reset_index()
    del ekhtesasi_haram_razavi_popular_visit['index']
    ekhtesasi_haram_razavi_popular_duration = ekhtesasi_haram_razavi_popular_duration.reset_index()
    del ekhtesasi_haram_razavi_popular_duration['index']
    
    ekhtesasi_javaneh_popular_visit = ekhtesasi_javaneh_popular_visit.reset_index()
    del ekhtesasi_javaneh_popular_visit['index']
    ekhtesasi_javaneh_popular_duration = ekhtesasi_javaneh_popular_duration.reset_index()
    del ekhtesasi_javaneh_popular_duration['index']
    
    ekhtesasi_jam_popular_visit = ekhtesasi_jam_popular_visit.reset_index()
    del ekhtesasi_jam_popular_visit['index']
    ekhtesasi_jam_popular_duration = ekhtesasi_jam_popular_duration.reset_index()
    del ekhtesasi_jam_popular_duration['index']
    
    ekhtesasi_habib_popular_visit = ekhtesasi_habib_popular_visit.reset_index()
    del ekhtesasi_habib_popular_visit['index']
    ekhtesasi_habib_popular_duration = ekhtesasi_habib_popular_duration.reset_index()
    del ekhtesasi_habib_popular_duration['index']
    
    ekhtesasi_rahro_popular_visit = ekhtesasi_rahro_popular_visit.reset_index()
    del ekhtesasi_rahro_popular_visit['index']
    ekhtesasi_rahro_popular_duration = ekhtesasi_rahro_popular_duration.reset_index()
    del ekhtesasi_rahro_popular_duration['index']
    
    ekhtesasi_iran_economy_popular_visit = ekhtesasi_iran_economy_popular_visit.reset_index()
    del ekhtesasi_iran_economy_popular_visit['index']
    ekhtesasi_iran_economy_popular_duration = ekhtesasi_iran_economy_popular_duration.reset_index()
    del ekhtesasi_iran_economy_popular_duration['index']
    
    ekhtesasi_nama_popular_visit = ekhtesasi_nama_popular_visit.reset_index()
    del ekhtesasi_nama_popular_visit['index']
    ekhtesasi_nama_popular_duration = ekhtesasi_nama_popular_duration.reset_index()
    del ekhtesasi_nama_popular_duration['index']
    
    ekhtesasi_aio_sport_popular_visit = ekhtesasi_aio_sport_popular_visit.reset_index()
    del ekhtesasi_aio_sport_popular_visit['index']
    ekhtesasi_aio_sport_popular_duration = ekhtesasi_aio_sport_popular_duration.reset_index()
    del ekhtesasi_aio_sport_popular_duration['index']
    
    
    ekhtesasi_channels_popular_content=pd.DataFrame()
    ekhtesasi_channels_popular_content=pd.concat([ekhtesasi_tva_kodak_popular_visit, ekhtesasi_tva_kodak_popular_duration,
                                               ekhtesasi_tva_nava_popular_visit, ekhtesasi_tva_kodak_popular_duration,
                                               ekhtesasi_tva_one_popular_visit, ekhtesasi_tva_one_popular_duration,
                                               ekhtesasi_mahfel_popular_visit, ekhtesasi_mahfel_popular_duration,
                                               ekhtesasi_tva_film_popular_visit, ekhtesasi_tva_film_popular_duration,
                                               ekhtesasi_tva_two_popular_visit, ekhtesasi_tva_two_popular_duration,
                                               ekhtesasi_tva_boors_popular_visit, ekhtesasi_tva_boors_popular_duration,
                                               ekhtesasi_tva_sport_two_popular_visit, ekhtesasi_tva_sport_two_popular_duration,
                                               ekhtesasi_tva_sport_popular_visit, ekhtesasi_tva_sport_popular_duration,
                                               ekhtesasi_lenz_sport_plus_popular_visit, ekhtesasi_lenz_sport_plus_popular_duration,
                                               ekhtesasi_lenz_sport_popular_visit, ekhtesasi_lenz_sport_popular_duration,
                                               ekhtesasi_kodak_digiton_popular_visit, ekhtesasi_kodak_digiton_popular_duration,
                                               ekhtesasi_shetab_popular_visit, ekhtesasi_shetab_popular_duration,
                                               ekhtesasi_shaparak_popular_visit, ekhtesasi_shaparak_popular_duration,
                                               ekhtesasi_esteghlal_popular_visit, ekhtesasi_esteghlal_popular_duration,
                                               ekhtesasi_perspolis_popular_visit, ekhtesasi_perspolis_popular_duration,
                                               ekhtesasi_tva_avand_popular_visit, ekhtesasi_tva_avand_popular_duration,
                                               ekhtesasi_tva_sport_one_popular_visit, ekhtesasi_tva_sport_one_popular_duration,
                                               ekhtesasi_mes_rafsanjan_popular_visit, ekhtesasi_mes_rafsanjan_popular_duration,
                                               ekhtesasi_emroz_popular_visit, ekhtesasi_emroz_popular_duration,
                                               ekhtesasi_keypad_popular_visit, ekhtesasi_keypad_popular_duration,
                                               ekhtesasi_dorfa_popular_visit, ekhtesasi_dorfa_popular_duration,
                                               ekhtesasi_lenz_film_popular_visit, ekhtesasi_lenz_film_popular_duration,
                                               ekhtesasi_eco_pars_popular_visit, ekhtesasi_eco_pars_popular_duration,
                                               ekhtesasi_ara_popular_visit, ekhtesasi_ara_popular_duration,
                                               ekhtesasi_eshareh_popular_visit, ekhtesasi_eshareh_popular_duration,
                                               ekhtesasi_astan_ghods_razavi_popular_visit, ekhtesasi_astan_ghods_razavi_popular_duration,
                                               ekhtesasi_aio_jahanbin_popular_visit, ekhtesasi_aio_jahanbin_popular_duration,
                                               ekhtesasi_sepahan_tv_popular_visit, ekhtesasi_sepahan_tv_popular_duration,
                                               ekhtesasi_borsan_popular_visit, ekhtesasi_borsan_popular_duration,
                                               ekhtesasi_haram_razavi_popular_visit, ekhtesasi_haram_razavi_popular_duration,
                                               ekhtesasi_javaneh_popular_visit, ekhtesasi_javaneh_popular_duration,
                                               ekhtesasi_jam_popular_visit, ekhtesasi_jam_popular_duration,
                                               ekhtesasi_habib_popular_visit, ekhtesasi_habib_popular_duration,
                                               ekhtesasi_rahro_popular_visit, ekhtesasi_rahro_popular_duration,
                                               ekhtesasi_iran_economy_popular_visit, ekhtesasi_iran_economy_popular_duration,
                                               ekhtesasi_nama_popular_visit, ekhtesasi_nama_popular_duration,
                                               ekhtesasi_aio_sport_popular_visit, ekhtesasi_aio_sport_popular_duration,],axis=1)
    
#    writer = pd.ExcelWriter('output/آمار ماه جاری/آمار اختصاصی.xlsx', engine='xlsxwriter')
#    ekhtesasi_channels_statistics.to_excel(writer, 'آمار شبکه های اختصاصی')
#    ekhtesasi_channels_popular_content.to_excel(writer, 'محتواهای پربازدید')
#    writer.save()
    
#    writer = pd.ExcelWriter('output/moh.rast/آمار اختصاصی.xlsx', engine='xlsxwriter')
#    ekhtesasi_channels_statistics.to_excel(writer, 'آمار شبکه های اختصاصی')
#    ekhtesasi_channels_popular_content.to_excel(writer, 'محتواهای پربازدید')
#    writer.save()
    
    writer = pd.ExcelWriter('output/output.sending.hard/آمار اختصاصی.xlsx', engine='xlsxwriter')
    ekhtesasi_channels_statistics.to_excel(writer, 'آمار شبکه های اختصاصی', index=False)
    ekhtesasi_channels_popular_content.to_excel(writer, 'محتواهای پربازدید', index=False)
    writer.save()
    
    print("END EKHTESASI")
    
    
    return ekhtesasi_channels_statistics, ekhtesasi_channels_popular_content
        
        
        
        
        
        
        
        
        
        
        
        
        
        