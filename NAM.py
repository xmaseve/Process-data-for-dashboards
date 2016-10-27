# -*- coding: utf-8 -*-
"""
Created on Wed Oct 19 14:21:11 2016

@author: Qi Yi
"""

import pandas as pd
from pandas import ExcelWriter

data = pd.read_excel('C:\\Users\\Qi Yi\\Desktop\Task11\WUG ad report(NAM).xlsx', sheetname="Data")   
WUGdata= data[["Ad","Description line 1","Description line 2","Final URL","Conversions","CTR","Cost","Month","Week"]]
WUGdata["Product"] = "WUG"
WUG_product = list(WUGdata["Product"].unique())

FT = pd.read_excel('C:\\Users\\Qi Yi\\Desktop\Task11\FTP ad report(NAM).xlsx', sheetname="Data")
FTdata = FT[["Ad","Description line 1","Description line 2","Final URL","Conversions","CTR","Cost","Month","Week","Product"]]
FT_product = list(FT["Product"].unique())
month = list(WUGdata["Month"].unique())
week = list(WUGdata["Week"].unique())


def process_Month(FTdata,FT_product,month):
    top10 = []
    sum_top10 = []
    #sum_other = []
    for prod in FT_product:
        a = FTdata[FTdata["Product"].str.contains(prod)]
        for mon in month:
            b = a[a["Month"].str.contains(mon)]
            c = b.groupby(["Ad","Description line 1","Description line 2","Final URL"]).sum()
            d = c.sort_values(["Conversions"], ascending=False)
            d10 = d.head(10)
            d10["Month"] = mon
            d10["Product"] = prod
            top10.append(d10)
            
            top10_sum = d.head(10).sum()
            other_sum = d.sum() - top10_sum
            top10_sum["Month"] = mon
            top10_sum["Product"] = prod
            top10_sum["Type"] = "Top 10"
            sum_top10.append(top10_sum)
            other_sum["Month"] = mon
            other_sum["Product"] = prod
            other_sum["Type"] = "Others"
            sum_top10.append(other_sum)
    return top10, sum_top10 #, sum_other
    
top10, sum_top10 = process_Month(WUGdata,WUG_product,month)
top10 = pd.concat(top10)
sum_top10 = pd.DataFrame(sum_top10)
#sum_other = pd.DataFrame(sum_other)
#other = pd.concat([sum_top10, sum_other], axis=0)
def process_Week(FTdata,FT_product,week):
    top10 = []
    sum_top10 = []
    #sum_other = []
    for prod in FT_product:
        a = FTdata[FTdata["Product"].str.contains(prod)]
        for wk in week:
            b = a[a["Week"] == wk]
            c = b.groupby(["Ad","Description line 1","Description line 2","Final URL"]).sum()
            d = c.sort_values(["Conversions"], ascending=False)
            d10 = d.head(10)
            d10["Week"] = wk
            d10["Product"] = prod
            top10.append(d10)
            
            top10_sum = d.head(10).sum()
            other_sum = d.sum() - top10_sum
            top10_sum["Week"] = str(wk)
            top10_sum["Product"] = prod
            top10_sum["Type"] = "Top 10"
            sum_top10.append(top10_sum)
            other_sum["Week"] = str(wk)
            other_sum["Product"] = prod
            other_sum["Type"] = "Others"
            sum_top10.append(other_sum)
    return top10, sum_top10 #, sum_other

top10, sum_top10 = process_Week(WUGdata,WUG_product,week)
top10 = pd.concat(top10)
sum_top10 = pd.DataFrame(sum_top10)    
    
writer = ExcelWriter("C:\\Users\\Qi Yi\\Desktop\\Ad data week(NAM).xlsx") #change the file path 
sum_top10.to_excel(writer, "Data for Graph")  
top10.to_excel(writer, "Top 10 Data") 
writer.save()        
#NAMdata = data[data["Campaign"].str.contains("NAM")]
#NAM = NAMdata[NAMdata["Keyword state"].str.contains("enabled")]               
             
'''              
camp = NAMdata["Campaign"].values.tolist()                        
FTaccount = NAMdata[NAMdata["Campaign"].str.contains("IFT")]
WUGaccount = NAMdata[NAMdata["Campaign"].str.contains("WUG")]
WUGaccount["Product"] = "WUG"
moveit = FTaccount[FTaccount["Campaign"].str.contains("MOVEit")]
others = FTaccount[~FTaccount["Campaign"].str.contains("MOVEit")]
moveit["Product"] = "MOVEit"
others["Product"] = "Others"    
NAM = pd.concat([WUGaccount, moveit, others])
                 
                   
FTcampaigns = list(FTaccount["Campaign"].unique())
WUGcampaigns = list(WUGaccount["Campaign"].unique())
FTcampaigns.sort()
WUGcampaigns.sort()
len(WUGcampaigns)
searchterms = searchquery.values.tolist()
'''

               
NAMdata.to_excel('C:\\Users\\Qi Yi\\Desktop\WUG ad report(NAM).xlsx')               