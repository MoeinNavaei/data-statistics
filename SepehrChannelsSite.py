    
def SepehrChannelsSite(site_channels_primary):
    
    import pandas as pd
    

#    sepehr_primary=sepehr_primary.rename(columns={"Event Action":"Action"})
#    sepehr_primary=sepehr_primary.rename(columns={"Avg. Session Duration":"Avg"})

#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('tv3', 'شبکه 3')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('varzesh', 'ورزش')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('ifilmfa', 'آی فیلم')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('tv1', 'شبکه 1')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('tv2', 'شبکه 2')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('irinn', 'خبر')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('namayesh', 'نمایش')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('nasim', 'نسیم')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('tv5', 'شبکه 5')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('tamasha', 'تماشا')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('shoma', 'شما')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('mostanad', 'مستند')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('pooya', 'پویا')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('omid', 'امید')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('amouzesh', 'آموزش')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('quran', 'قرآن')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('tv4', 'شبکه 4')
#    
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('ofogh', 'افق')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('salamat', 'سلامت')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('sepehr', 'سپهر')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('irankala', 'ایران کالا')
#    
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-ava', 'رادیو آوا')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-javan', 'رادیو جوان')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-payam', 'رادیو پیام')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-iran', 'رادیو ایران')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-varzesh', 'رادیو ورزش')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-maaref', 'رادیو معارف')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-farhang', 'رادیو فرهنگ')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-fasli', 'رادیو فصلی')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-talavat', 'رادیو تلاوت')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('rn-salamat', 'رادیو سلامت')
#    sepehr_primary['Action'] = sepehr_primary['Action'].str.replace('noor', 'استانی قم-نور')
#
#    sepehr_primary['Action'].replace('', 'NO', inplace=True)
#    sepehr_primary = sepehr_primary[~sepehr_primary.Action.str.contains("NO")]
##    sepehr_primary.dropna(inplace=True)
#
#    sepehr_primary_counter=sepehr_primary['Action']
#    length_sepehr_primary = len(sepehr_primary)
#    for i in range(0,length_sepehr_primary):
#         x_name_content=sepehr_primary_counter[i]
#         head, sep, tail = x_name_content.partition('-')
#         if head=="rn":
#             sepehr_primary['Action'].replace(sepehr_primary_counter[i], '', inplace=True)
#             sepehr_primary = sepehr_primary[sepehr_primary.Action != '']
#    sepehr_primary.to_excel('sepehr_primary4.xlsx', index=False)
#    sepehr_primary.insert(5, 'duration(hour)', '')
#    sepehr_primary['duration(hour)']=round(sepehr_primary.Sessions*sepehr_primary.Avg/60, 0)
#    
#    del sepehr_primary['% New Sessions']
#    del sepehr_primary['Pages / Session']
#    del sepehr_primary['Avg']
#
#    sepehr_final=pd.DataFrame()
#    sepehr_final.insert(0, 'channel_name', '')
#    sepehr_final.insert(1, 'نام برنامه اولیه', '')
#    sepehr_final.insert(2, 'نام برنامه', '')
#    sepehr_final.insert(3, 'تاریخ شروع', '')
#    sepehr_final.insert(4, 'تاریخ پایان', '')
#    sepehr_final.insert(5, 'مدت بازدید', '')
#    sepehr_final.insert(6, 'تعداد بازدید', '')
#    sepehr_final.insert(7, 'میانگین', '')
#    sepehr_final.insert(8, 'اپراتور', '')
#    sepehr_final.insert(9, 'ساعت', '')
#    sepehr_final.insert(10, 'تاریخ', '')
#    sepehr_final.insert(11, 'ردیف', '')
#    sepehr_final.insert(12, 'جنس', '')
#    sepehr_final.insert(13, 'tag', '')
#    sepehr_final.insert(14, 'type', '')
#    
#    sepehr_final['channel_name']=sepehr_primary['Action']
#    sepehr_final['تعداد بازدید']=sepehr_primary['Sessions']
#    sepehr_final['مدت بازدید']=sepehr_primary['duration(hour)']
#    sepehr_final['نام برنامه']="-"
#    sepehr_final['نام برنامه اولیه']="-"
#    sepehr_final['اپراتور']="سپهر"
#    sepehr_final['tag']="سایر"
#    
#    sepehr_final = sepehr_final.reset_index()
#    del sepehr_final['index']
#    length_sepehr_final=len(sepehr_final)
#    for i in range(length_sepehr_final):
#        try:
#            if sepehr_final.loc[i, 'channel_name'] == "شبکه 1":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "شبکه 3":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "ورزش":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "شبکه 2":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "خبر":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "آی فیلم":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "شبکه 5":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "نمایش":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "تماشا":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "نسیم":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "شبکه 4":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "شما":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "مستند":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "پویا":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "امید":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "آموزش":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "قرآن":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "افق":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "سلامت":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "سپهر":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "ایران کالا":
#                sepehr_final.loc[i, 'type'] = "سراسری"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو آوا":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو جوان":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو پیام":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو ایران":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو معارف":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو فرهنگ":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو فصلی":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو تلاوت":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "رادیو سلامت":
#                sepehr_final.loc[i, 'type'] = "رادیویی"
#            elif sepehr_final.loc[i, 'channel_name'] == "استانی قم-نور":
#                sepehr_final.loc[i, 'type'] = "استانی"
#        except: pass
#
#    sepehr_final=sepehr_final.rename(columns={"channel_name":"نام شبکه"})
#    sepehr_final=sepehr_final.rename(columns={"type":"نوع"})
    
    site_channels_final=pd.DataFrame()
    site_channels_final.insert(0, 'نام شبکه', '')
    site_channels_final.insert(1, 'نام برنامه اولیه', '')
    site_channels_final.insert(2, 'نام برنامه', '')
    site_channels_final.insert(3, 'تاریخ شروع', '')
    site_channels_final.insert(4, 'تاریخ پایان', '')
    site_channels_final.insert(5, 'مدت بازدید', '')
    site_channels_final.insert(6, 'تعداد بازدید', '')
    site_channels_final.insert(7, 'میانگین', '')
    site_channels_final.insert(8, 'اپراتور', '')
    site_channels_final.insert(9, 'ساعت', '')
    site_channels_final.insert(10, 'تاریخ', '')
    site_channels_final.insert(11, 'ردیف', '')
    site_channels_final.insert(12, 'جنس', '')
    site_channels_final.insert(13, 'tag', '')
    site_channels_final.insert(14, 'نوع', '')
    
    site_channels_final['نام شبکه']=site_channels_primary['نام شبکه']
    site_channels_final['تعداد بازدید']=site_channels_primary['تعداد بازدید']
    site_channels_final['مدت بازدید']=site_channels_primary['زمان بازدید (ساعت)']
    site_channels_final['نام برنامه']="-"
    site_channels_final['نام برنامه اولیه']="-"
    site_channels_final['اپراتور']="سایت شبکه ها"
    site_channels_final['tag']="سایر"
    site_channels_final['نوع']="سراسری"
    
#    sepehr_site_channels=pd.DataFrame()
#    sepehr_site_channels = sepehr_final.append([site_channels_final])
    sepehr_site_channels = site_channels_final.copy()
    return sepehr_site_channels
    
    
    
    
    
    
    
    
    
    
    
    
