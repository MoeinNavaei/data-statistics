    
def radio_data(radio):
            
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
        
    print("start radio")
    
    radio_all=radio.copy()
    radio_all_pivot=radio_all.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_all_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_all_popular_visit=radio_all_pivot.iloc[0:10 , [0, 5]]
    radio_all_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_all_popular_duration=radio_all_pivot.iloc[0:10 , [0, 5]] 
    
    print("radio_eghtesad")
    radio_eghtesad=radio.query("channel == 'رادیو اقتصاد'")
    radio_eghtesad_visit=radio_eghtesad['تعداد بازدید'].sum()
    radio_eghtesad_duration=radio_eghtesad['مدت بازدید'].sum()
    radio_eghtesad_duration=round(radio_eghtesad_duration, 0)
    radio_eghtesad_content=radio_eghtesad.copy()
    radio_eghtesad_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_eghtesad_content=len(radio_eghtesad_content)
    radio_eghtesad_pivot=radio_eghtesad.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_eghtesad_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_eghtesad_popular_visit=radio_eghtesad_pivot.iloc[0:10 , [0, 3]]
    radio_eghtesad_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_eghtesad_popular_duration=radio_eghtesad_pivot.iloc[0:10 , [0, 5]]
    
    radio_eghtesad_popular_visit = radio_eghtesad_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو اقتصاد', 'نام برنامه': 'محتواهای پربازدید رادیو اقتصاد'})
    radio_eghtesad_popular_duration = radio_eghtesad_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو اقتصاد (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو اقتصاد'})
    
    print("radio_ava")
    radio_ava=radio.query("channel == 'رادیو آوا'")
    radio_ava_visit=radio_ava['تعداد بازدید'].sum()
    radio_ava_duration=radio_ava['مدت بازدید'].sum()
    radio_ava_duration=round(radio_ava_duration, 0)
    radio_ava_content=radio_ava.copy()
    radio_ava_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_ava_content=len(radio_ava_content)
    radio_ava_pivot=radio_ava.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_ava_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_ava_popular_visit=radio_ava_pivot.iloc[0:10 , [0, 3]]
    radio_ava_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_ava_popular_duration=radio_ava_pivot.iloc[0:10 , [0, 5]]
    
    radio_ava_popular_visit = radio_ava_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو آوا', 'نام برنامه': 'محتواهای پربازدید رادیو آوا'})
    radio_ava_popular_duration = radio_ava_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو آوا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو آوا'})
    
    print("radio_iran")
    radio_iran=radio.query("channel == 'رادیو ایران'")
    radio_iran_visit=radio_iran['تعداد بازدید'].sum()
    radio_iran_duration=radio_iran['مدت بازدید'].sum()
    radio_iran_duration=round(radio_iran_duration, 0)
    radio_iran_content=radio_iran.copy()
    radio_iran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_iran_content=len(radio_iran_content)
    radio_iran_pivot=radio_iran.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_iran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_iran_popular_visit=radio_iran_pivot.iloc[0:10 , [0, 3]]
    radio_iran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_iran_popular_duration=radio_iran_pivot.iloc[0:10 , [0, 5]]
    
    radio_iran_popular_visit = radio_iran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو ایران', 'نام برنامه': 'محتواهای پربازدید رادیو ایران'})
    radio_iran_popular_duration = radio_iran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو ایران (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو ایران'})
    
    print("radio_payam")
    radio_payam=radio.query("channel == 'رادیو پیام'")
    radio_payam_visit=radio_payam['تعداد بازدید'].sum()
    radio_payam_duration=radio_payam['مدت بازدید'].sum()
    radio_payam_duration=round(radio_payam_duration, 0)
    radio_payam_content=radio_payam.copy()
    radio_payam_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_payam_content=len(radio_payam_content)
    radio_payam_pivot=radio_payam.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_payam_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_payam_popular_visit=radio_payam_pivot.iloc[0:10 , [0, 3]]
    radio_payam_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_payam_popular_duration=radio_payam_pivot.iloc[0:10 , [0, 5]]
    
    radio_payam_popular_visit = radio_payam_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو پیام', 'نام برنامه': 'محتواهای پربازدید رادیو پیام'})
    radio_payam_popular_duration = radio_payam_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو پیام (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو پیام'})
    
    print("radio_javan")
    radio_javan=radio.query("channel == 'رادیو جوان'")
    radio_javan_visit=radio_javan['تعداد بازدید'].sum()
    radio_javan_duration=radio_javan['مدت بازدید'].sum()
    radio_javan_duration=round(radio_javan_duration, 0)
    radio_javan_content=radio_javan.copy()
    radio_javan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_javan_content=len(radio_javan_content)
    radio_javan_pivot=radio_javan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_javan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_javan_popular_visit=radio_javan_pivot.iloc[0:10 , [0, 3]]
    radio_javan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_javan_popular_duration=radio_javan_pivot.iloc[0:10 , [0, 5]]
    
    radio_javan_popular_visit = radio_javan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو جوان', 'نام برنامه': 'محتواهای پربازدید رادیو جوان'})
    radio_javan_popular_duration = radio_javan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو جوان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو جوان'})
    
    print("radio_salamat")
    radio_salamat=radio.query("channel == 'رادیو سلامت'")
    radio_salamat_visit=radio_salamat['تعداد بازدید'].sum()
    radio_salamat_duration=radio_salamat['مدت بازدید'].sum()
    radio_salamat_duration=round(radio_salamat_duration, 0)
    radio_salamat_content=radio_salamat.copy()
    radio_salamat_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_salamat_content=len(radio_salamat_content)
    radio_salamat_pivot=radio_salamat.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_salamat_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_salamat_popular_visit=radio_salamat_pivot.iloc[0:10 , [0, 3]]
    radio_salamat_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_salamat_popular_duration=radio_salamat_pivot.iloc[0:10 , [0, 5]]
    
    radio_salamat_popular_visit = radio_salamat_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو سلامت', 'نام برنامه': 'محتواهای پربازدید رادیو سلامت'})
    radio_salamat_popular_duration = radio_salamat_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو سلامت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو سلامت'})
    
    print("radio_saba")
    radio_saba=radio.query("channel == 'رادیو صبا'")
    radio_saba_visit=radio_saba['تعداد بازدید'].sum()
    radio_saba_duration=radio_saba['مدت بازدید'].sum()
    radio_saba_duration=round(radio_saba_duration, 0)
    radio_saba_content=radio_saba.copy()
    radio_saba_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_saba_content=len(radio_saba_content)
    radio_saba_pivot=radio_saba.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_saba_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_saba_popular_visit=radio_saba_pivot.iloc[0:10 , [0, 3]]
    radio_saba_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_saba_popular_duration=radio_saba_pivot.iloc[0:10 , [0, 5]]
    
    radio_saba_popular_visit = radio_saba_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو سبا', 'نام برنامه': 'محتواهای پربازدید رادیو سبا'})
    radio_saba_popular_duration = radio_saba_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو سبا (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو سبا'})
    
    print("radio_farhang")
    radio_farhang=radio.query("channel == 'رادیو فرهنگ'")
    radio_farhang_visit=radio_farhang['تعداد بازدید'].sum()
    radio_farhang_duration=radio_farhang['مدت بازدید'].sum()
    radio_farhang_duration=round(radio_farhang_duration, 0)
    radio_farhang_content=radio_farhang.copy()
    radio_farhang_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_farhang_content=len(radio_farhang_content)
    radio_farhang_pivot=radio_farhang.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_farhang_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_farhang_popular_visit=radio_farhang_pivot.iloc[0:10 , [0, 3]]
    radio_farhang_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_farhang_popular_duration=radio_farhang_pivot.iloc[0:10 , [0, 5]]
    
    radio_farhang_popular_visit = radio_farhang_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو فرهنگ', 'نام برنامه': 'محتواهای پربازدید رادیو فرهنگ'})
    radio_farhang_popular_duration = radio_farhang_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو فرهنگ (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو فرهنگ'})
    
    print("radio_qoran")
    radio_qoran=radio.query("channel == 'رادیو قرآن'")
    radio_qoran_visit=radio_qoran['تعداد بازدید'].sum()
    radio_qoran_duration=radio_qoran['مدت بازدید'].sum()
    radio_qoran_duration=round(radio_qoran_duration, 0)
    radio_qoran_content=radio_qoran.copy()
    radio_qoran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_qoran_content=len(radio_qoran_content)
    radio_qoran_pivot=radio_qoran.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_qoran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_qoran_popular_visit=radio_qoran_pivot.iloc[0:10 , [0, 3]]
    radio_qoran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_qoran_popular_duration=radio_qoran_pivot.iloc[0:10 , [0, 5]]
    
    radio_qoran_popular_visit = radio_qoran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو قرآن', 'نام برنامه': 'محتواهای پربازدید رادیو قرآن'})
    radio_qoran_popular_duration = radio_qoran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو قرآن (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو قرآن'})
    
    print("radio_goftego")
    radio_goftego=radio.query("channel == 'رادیو گفتگو'")
    radio_goftego_visit=radio_goftego['تعداد بازدید'].sum()
    radio_goftego_duration=radio_goftego['مدت بازدید'].sum()
    radio_goftego_duration=round(radio_goftego_duration, 0)
    radio_goftego_content=radio_goftego.copy()
    radio_goftego_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_goftego_content=len(radio_goftego_content)
    radio_goftego_pivot=radio_goftego.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_goftego_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_goftego_popular_visit=radio_goftego_pivot.iloc[0:10 , [0, 3]]
    radio_goftego_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_goftego_popular_duration=radio_goftego_pivot.iloc[0:10 , [0, 5]]
    
    radio_goftego_popular_visit = radio_goftego_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو گفتگو', 'نام برنامه': 'محتواهای پربازدید رادیو گفتگو'})
    radio_goftego_popular_duration = radio_goftego_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو گفتگو (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو گفتگو'})
    
    print("radio_maaref")
    radio_maaref=radio.query("channel == 'رادیو معارف'")
    radio_maaref_visit=radio_maaref['تعداد بازدید'].sum()
    radio_maaref_duration=radio_maaref['مدت بازدید'].sum()
    radio_maaref_duration=round(radio_maaref_duration, 0)
    radio_maaref_content=radio_maaref.copy()
    radio_maaref_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_maaref_content=len(radio_maaref_content)
    radio_maaref_pivot=radio_maaref.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_maaref_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_maaref_popular_visit=radio_maaref_pivot.iloc[0:10 , [0, 3]]
    radio_maaref_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_maaref_popular_duration=radio_maaref_pivot.iloc[0:10 , [0, 5]]
    
    radio_maaref_popular_visit = radio_maaref_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو معارف', 'نام برنامه': 'محتواهای پربازدید رادیو معارف'})
    radio_maaref_popular_duration = radio_maaref_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو معارف (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو معارف'})
    
    print("radio_namayesh")
    radio_namayesh=radio.query("channel == 'رادیو نمایش'")
    radio_namayesh_visit=radio_namayesh['تعداد بازدید'].sum()
    radio_namayesh_duration=radio_namayesh['مدت بازدید'].sum()
    radio_namayesh_duration=round(radio_namayesh_duration, 0)
    radio_namayesh_content=radio_namayesh.copy()
    radio_namayesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_namayesh_content=len(radio_namayesh_content)
    radio_namayesh_pivot=radio_namayesh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_namayesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_namayesh_popular_visit=radio_namayesh_pivot.iloc[0:10 , [0, 3]]
    radio_namayesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_namayesh_popular_duration=radio_namayesh_pivot.iloc[0:10 , [0, 5]]
    
    radio_namayesh_popular_visit = radio_namayesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو نمایش', 'نام برنامه': 'محتواهای پربازدید رادیو نمایش'})
    radio_namayesh_popular_duration = radio_namayesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو نمایش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو نمایش'})
    
    print("radio_varzesh")
    radio_varzesh=radio.query("channel == 'رادیو ورزش'")
    radio_varzesh_visit=radio_varzesh['تعداد بازدید'].sum()
    radio_varzesh_duration=radio_varzesh['مدت بازدید'].sum()
    radio_varzesh_duration=round(radio_varzesh_duration, 0)
    radio_varzesh_content=radio_varzesh.copy()
    radio_varzesh_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_varzesh_content=len(radio_varzesh_content)
    radio_varzesh_pivot=radio_varzesh.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_varzesh_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_varzesh_popular_visit=radio_varzesh_pivot.iloc[0:10 , [0, 3]]
    radio_varzesh_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_varzesh_popular_duration=radio_varzesh_pivot.iloc[0:10 , [0, 5]]
    
    radio_varzesh_popular_visit = radio_varzesh_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو ورزش', 'نام برنامه': 'محتواهای پربازدید رادیو ورزش'})
    radio_varzesh_popular_duration = radio_varzesh_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو ورزش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو ورزش'})
    
    print("radio_ardebil")
    radio_ardebil=radio.query("channel == 'رادیو اردبیل'")
    radio_ardebil_visit=radio_ardebil['تعداد بازدید'].sum()
    radio_ardebil_duration=radio_ardebil['مدت بازدید'].sum()
    radio_ardebil_duration=round(radio_ardebil_duration, 0)
    radio_ardebil_content=radio_ardebil.copy()
    radio_ardebil_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_ardebil_content=len(radio_ardebil_content)
    radio_ardebil_pivot=radio_ardebil.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_ardebil_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_ardebil_popular_visit=radio_ardebil_pivot.iloc[0:10 , [0, 3]]
    radio_ardebil_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_ardebil_popular_duration=radio_ardebil_pivot.iloc[0:10 , [0, 5]]
    
    radio_ardebil_popular_visit = radio_ardebil_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو اردبیل', 'نام برنامه': 'محتواهای پربازدید رادیو اردبیل'})
    radio_ardebil_popular_duration = radio_ardebil_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو اردبیل (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو اردبیل'})
    
    print("radio_yazd")
    radio_yazd=radio.query("channel == 'رادیو یزد'")
    radio_yazd_visit=radio_yazd['تعداد بازدید'].sum()
    radio_yazd_duration=radio_yazd['مدت بازدید'].sum()
    radio_yazd_duration=round(radio_yazd_duration, 0)
    radio_yazd_content=radio_yazd.copy()
    radio_yazd_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_yazd_content=len(radio_yazd_content)
    radio_yazd_pivot=radio_yazd.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_yazd_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_yazd_popular_visit=radio_yazd_pivot.iloc[0:10 , [0, 3]]
    radio_yazd_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_yazd_popular_duration=radio_yazd_pivot.iloc[0:10 , [0, 5]]
    
    radio_yazd_popular_visit = radio_yazd_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو یزد', 'نام برنامه': 'محتواهای پربازدید رادیو یزد'})
    radio_yazd_popular_duration = radio_yazd_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو یزد (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو یزد'})
    
    print("radio_hamedan")
    radio_hamedan=radio.query("channel == 'رادیو همدان'")
    radio_hamedan_visit=radio_hamedan['تعداد بازدید'].sum()
    radio_hamedan_duration=radio_hamedan['مدت بازدید'].sum()
    radio_hamedan_duration=round(radio_hamedan_duration, 0)
    radio_hamedan_content=radio_hamedan.copy()
    radio_hamedan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_hamedan_content=len(radio_hamedan_content)
    radio_hamedan_pivot=radio_hamedan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_hamedan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_hamedan_popular_visit=radio_hamedan_pivot.iloc[0:10 , [0, 3]]
    radio_hamedan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_hamedan_popular_duration=radio_hamedan_pivot.iloc[0:10 , [0, 5]]
    
    radio_hamedan_popular_visit = radio_hamedan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو همدان', 'نام برنامه': 'محتواهای پربازدید رادیو همدان'})
    radio_hamedan_popular_duration = radio_hamedan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو همدان (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو همدان'})
    
    print("radio_markazi")
    radio_markazi=radio.query("channel == 'رادیو مرکزی'")
    radio_markazi_visit=radio_markazi['تعداد بازدید'].sum()
    radio_markazi_duration=radio_markazi['مدت بازدید'].sum()
    radio_markazi_duration=round(radio_markazi_duration, 0)
    radio_markazi_content=radio_markazi.copy()
    radio_markazi_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_markazi_content=len(radio_markazi_content)
    radio_markazi_pivot=radio_markazi.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_markazi_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_markazi_popular_visit=radio_markazi_pivot.iloc[0:10 , [0, 3]]
    radio_markazi_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_markazi_popular_duration=radio_markazi_pivot.iloc[0:10 , [0, 5]]
    
    radio_markazi_popular_visit = radio_markazi_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو مرکزی', 'نام برنامه': 'محتواهای پربازدید رادیو مرکزی'})
    radio_markazi_popular_duration = radio_markazi_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو مرکزی (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو مرکزی'})
    
    print("radio_telavat")
    radio_telavat=radio.query("channel == 'رادیو تلاوت'")
    radio_telavat_visit=radio_telavat['تعداد بازدید'].sum()
    radio_telavat_duration=radio_telavat['مدت بازدید'].sum()
    radio_telavat_duration=round(radio_telavat_duration, 0)
    radio_telavat_content=radio_telavat.copy()
    radio_telavat_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_telavat_content=len(radio_telavat_content)
    radio_telavat_pivot=radio_telavat.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_telavat_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_telavat_popular_visit=radio_telavat_pivot.iloc[0:10 , [0, 3]]
    radio_telavat_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_telavat_popular_duration=radio_telavat_pivot.iloc[0:10 , [0, 5]]
    
    radio_telavat_popular_visit = radio_telavat_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو تلاوت', 'نام برنامه': 'محتواهای پربازدید رادیو تلاوت'})
    radio_telavat_popular_duration = radio_telavat_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو تلاوت (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو تلاوت'})
    
    print("radio_tehran")
    radio_tehran=radio.query("channel == 'رادیو تهران'")
    radio_tehran_visit=radio_tehran['تعداد بازدید'].sum()
    radio_tehran_duration=radio_tehran['مدت بازدید'].sum()
    radio_tehran_duration=round(radio_tehran_duration, 0)
    radio_tehran_content=radio_tehran.copy()
    radio_tehran_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_tehran_content=len(radio_tehran_content)
    radio_tehran_pivot=radio_tehran.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_tehran_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_tehran_popular_visit=radio_tehran_pivot.iloc[0:10 , [0, 3]]
    radio_tehran_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_tehran_popular_duration=radio_tehran_pivot.iloc[0:10 , [0, 5]]
    
    radio_tehran_popular_visit = radio_tehran_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو تهران', 'نام برنامه': 'محتواهای پربازدید رادیو تهران'})
    radio_tehran_popular_duration = radio_tehran_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو تهران (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو تهران'})
    
    print("radio_gilan")
    radio_gilan=radio.query("channel == 'رادیو گیلان'")
    radio_gilan_visit=radio_gilan['تعداد بازدید'].sum()
    radio_gilan_duration=radio_gilan['مدت بازدید'].sum()
    radio_gilan_duration=round(radio_gilan_duration, 0)
    radio_gilan_content=radio_gilan.copy()
    radio_gilan_content.drop_duplicates(subset =['نام برنامه', 'channel'], keep = 'first', inplace = True)
    radio_gilan_content=len(radio_gilan_content)
    radio_gilan_pivot=radio_gilan.groupby(['نام برنامه','channel', 'type']).sum().reset_index()
    radio_gilan_pivot.sort_values('تعداد بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_gilan_popular_visit=radio_gilan_pivot.iloc[0:10 , [0, 3]]
    radio_gilan_pivot.sort_values('مدت بازدید', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_gilan_popular_duration=radio_gilan_pivot.iloc[0:10 , [0, 5]]
    
    radio_gilan_popular_visit = radio_gilan_popular_visit.rename(columns={'تعداد بازدید': 'تعداد بازدید رادیو ورزش', 'نام برنامه': 'محتواهای پربازدید رادیو گیلان'})
    radio_gilan_popular_duration = radio_gilan_popular_duration.rename(columns={'مدت بازدید': 'زمان بازدید رادیو ورزش (به دقیقه)', 'نام برنامه': 'محتواها با بیشترین زمان بازدید رادیو گیلان'})
    
    print("dataframe radio channels")
    radio_channels_statistics={'channel_name': ['رادیو اقتصاد',
                                                'رادیو آوا',
                                                'رادیو ایران',
                                                'رادیو پیام',
                                                'رادیو جوان',
                                         'رادیو سلامت',
                                         'رادیو صبا',
                                         'رادیو فرهنگ',
                                         'رادیو قرآن',
                                         'رادیو گفتگو',
                                         'رادیو معارف',
                                         'رادیو نمایش',
                                         'رادیو ورزش',
                                         'رادیو اردبیل',
                                         'رادیو یزد',
                                         'رادیو همدان',
                                         'رادیو مرکزی',
                                         'رادیو تلاوت',
                                         'رادیو تهران',
                                         'رادیو گیلان',],
           'channel_content': [radio_eghtesad_content, radio_ava_content, radio_iran_content, radio_payam_content, radio_javan_content,
                               radio_salamat_content, radio_saba_content, radio_farhang_content, radio_qoran_content, radio_goftego_content,
                               radio_maaref_content, radio_namayesh_content, radio_varzesh_content,
                               radio_ardebil_content, radio_yazd_content, radio_hamedan_content,
                               radio_markazi_content, radio_telavat_content, radio_tehran_content,
                               radio_gilan_content,],
           'channel_visit': [radio_eghtesad_visit, radio_ava_visit, radio_iran_visit, radio_payam_visit, radio_javan_visit,
                               radio_salamat_visit, radio_saba_visit, radio_farhang_visit, radio_qoran_visit, radio_goftego_visit,
                               radio_maaref_visit, radio_namayesh_visit, radio_varzesh_visit,
                               radio_ardebil_visit, radio_yazd_visit, radio_hamedan_visit,
                               radio_markazi_visit, radio_telavat_visit, radio_tehran_visit,
                               radio_gilan_visit,],
           'channel_duration': [radio_eghtesad_duration, radio_ava_duration, radio_iran_duration, radio_payam_duration, radio_javan_duration,
                               radio_salamat_duration, radio_saba_duration, radio_farhang_duration, radio_qoran_duration, radio_goftego_duration,
                               radio_maaref_duration, radio_namayesh_duration, radio_varzesh_duration,
                               radio_ardebil_duration, radio_yazd_duration, radio_hamedan_duration,
                               radio_markazi_duration, radio_telavat_duration, radio_tehran_duration,
                               radio_gilan_duration,],}
    radio_channels_statistics=pd.DataFrame(radio_channels_statistics, columns=['channel_name', 'channel_content', 'channel_visit', 'channel_duration'])
    radio_channels_statistics.sort_values('channel_visit', axis = 0, ascending = False, inplace = True, na_position ='last')
    radio_channels_statistics=radio_channels_statistics.rename(columns={'channel_name': 'نام شبکه', 'channel_content': 'تعداد محتوا', 'channel_visit': 'تعداد بازدید', 'channel_duration': 'مدت زمان بازدید (به دقیقه)'})
    
    radio_eghtesad_popular_visit = radio_eghtesad_popular_visit.reset_index()
    del radio_eghtesad_popular_visit['index']
    radio_eghtesad_popular_duration = radio_eghtesad_popular_duration.reset_index()
    del radio_eghtesad_popular_duration['index']
    
    radio_ava_popular_visit = radio_ava_popular_visit.reset_index()
    del radio_ava_popular_visit['index']
    radio_ava_popular_duration = radio_ava_popular_duration.reset_index()
    del radio_ava_popular_duration['index']
    
    radio_iran_popular_visit = radio_iran_popular_visit.reset_index()
    del radio_iran_popular_visit['index']
    radio_iran_popular_duration = radio_iran_popular_duration.reset_index()
    del radio_iran_popular_duration['index']
    
    radio_payam_popular_visit = radio_payam_popular_visit.reset_index()
    del radio_payam_popular_visit['index']
    radio_payam_popular_duration = radio_payam_popular_duration.reset_index()
    del radio_payam_popular_duration['index']
    
    radio_javan_popular_visit = radio_javan_popular_visit.reset_index()
    del radio_javan_popular_visit['index']
    radio_javan_popular_duration = radio_javan_popular_duration.reset_index()
    del radio_javan_popular_duration['index']
    
    radio_salamat_popular_visit = radio_salamat_popular_visit.reset_index()
    del radio_salamat_popular_visit['index']
    radio_salamat_popular_duration = radio_salamat_popular_duration.reset_index()
    del radio_salamat_popular_duration['index']
    
    radio_saba_popular_visit = radio_saba_popular_visit.reset_index()
    del radio_saba_popular_visit['index']
    radio_saba_popular_duration = radio_saba_popular_duration.reset_index()
    del radio_saba_popular_duration['index']
    
    radio_farhang_popular_visit = radio_farhang_popular_visit.reset_index()
    del radio_farhang_popular_visit['index']
    radio_farhang_popular_duration = radio_farhang_popular_duration.reset_index()
    del radio_farhang_popular_duration['index']
    
    radio_qoran_popular_visit = radio_qoran_popular_visit.reset_index()
    del radio_qoran_popular_visit['index']
    radio_qoran_popular_duration = radio_qoran_popular_duration.reset_index()
    del radio_qoran_popular_duration['index']
    
    radio_goftego_popular_visit = radio_goftego_popular_visit.reset_index()
    del radio_goftego_popular_visit['index']
    radio_goftego_popular_duration = radio_goftego_popular_duration.reset_index()
    del radio_goftego_popular_duration['index']
    
    radio_maaref_popular_visit = radio_maaref_popular_visit.reset_index()
    del radio_maaref_popular_visit['index']
    radio_maaref_popular_duration = radio_maaref_popular_duration.reset_index()
    del radio_maaref_popular_duration['index']
    
    radio_namayesh_popular_visit = radio_namayesh_popular_visit.reset_index()
    del radio_namayesh_popular_visit['index']
    radio_namayesh_popular_duration = radio_namayesh_popular_duration.reset_index()
    del radio_namayesh_popular_duration['index']
    
    radio_varzesh_popular_visit = radio_varzesh_popular_visit.reset_index()
    del radio_varzesh_popular_visit['index']
    radio_varzesh_popular_duration = radio_varzesh_popular_duration.reset_index()
    del radio_varzesh_popular_duration['index']
    
    radio_ardebil_popular_visit = radio_ardebil_popular_visit.reset_index()
    del radio_ardebil_popular_visit['index']
    radio_ardebil_popular_duration = radio_ardebil_popular_duration.reset_index()
    del radio_ardebil_popular_duration['index']
    
    radio_yazd_popular_visit = radio_yazd_popular_visit.reset_index()
    del radio_yazd_popular_visit['index']
    radio_yazd_popular_duration = radio_yazd_popular_duration.reset_index()
    del radio_yazd_popular_duration['index']
    
    radio_hamedan_popular_visit = radio_hamedan_popular_visit.reset_index()
    del radio_hamedan_popular_visit['index']
    radio_hamedan_popular_duration = radio_hamedan_popular_duration.reset_index()
    del radio_hamedan_popular_duration['index']
    
    radio_markazi_popular_visit = radio_markazi_popular_visit.reset_index()
    del radio_markazi_popular_visit['index']
    radio_markazi_popular_duration = radio_markazi_popular_duration.reset_index()
    del radio_markazi_popular_duration['index']
    
    radio_telavat_popular_visit = radio_telavat_popular_visit.reset_index()
    del radio_telavat_popular_visit['index']
    radio_telavat_popular_duration = radio_telavat_popular_duration.reset_index()
    del radio_telavat_popular_duration['index']
    
    radio_tehran_popular_visit = radio_tehran_popular_visit.reset_index()
    del radio_tehran_popular_visit['index']
    radio_tehran_popular_duration = radio_tehran_popular_duration.reset_index()
    del radio_tehran_popular_duration['index']
    
    radio_gilan_popular_visit = radio_gilan_popular_visit.reset_index()
    del radio_gilan_popular_visit['index']
    radio_gilan_popular_duration = radio_gilan_popular_duration.reset_index()
    del radio_gilan_popular_duration['index']
    

    radio_channels_popular_content=pd.DataFrame()
    radio_channels_popular_content=pd.concat([radio_eghtesad_popular_visit, radio_eghtesad_popular_duration,
                                              radio_ava_popular_visit, radio_ava_popular_duration,
                                              radio_iran_popular_visit, radio_iran_popular_duration,
                                              radio_payam_popular_visit, radio_payam_popular_duration,
                                              radio_javan_popular_visit, radio_javan_popular_duration,
                                              radio_salamat_popular_visit, radio_salamat_popular_duration,
                                              radio_saba_popular_visit, radio_saba_popular_duration,
                                              radio_farhang_popular_visit, radio_farhang_popular_duration,
                                              radio_qoran_popular_visit, radio_qoran_popular_duration,
                                              radio_goftego_popular_visit, radio_goftego_popular_duration,
                                              radio_maaref_popular_visit, radio_maaref_popular_duration,
                                              radio_namayesh_popular_visit, radio_namayesh_popular_duration,
                                              radio_varzesh_popular_visit, radio_varzesh_popular_duration,
                                              radio_ardebil_popular_visit, radio_ardebil_popular_duration,
                                              radio_yazd_popular_visit, radio_yazd_popular_duration,
                                              radio_hamedan_popular_visit, radio_hamedan_popular_duration,
                                              radio_markazi_popular_visit, radio_markazi_popular_duration,
                                              radio_telavat_popular_visit, radio_telavat_popular_duration,
                                              radio_tehran_popular_visit, radio_tehran_popular_duration,
                                              radio_gilan_popular_visit, radio_gilan_popular_duration,],axis=1)
    
    
    writer = pd.ExcelWriter('output/output.sending.hard/آمار رادیو.xlsx', engine='xlsxwriter')
    radio_channels_statistics.to_excel(writer, 'آمار شبکه های رادیویی', index=False)
    radio_channels_popular_content.to_excel(writer, 'محتواهای پربازدید', index=False)
    writer.save()
    
    print("End radio")
    
    
    
    return radio_channels_statistics, radio_channels_popular_content
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        