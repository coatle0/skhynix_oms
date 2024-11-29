import pandas as pd
import os
import openpyxl
import fnmatch
from datetime import datetime
import shutil
import xlwings as xw

today_str = datetime.today().strftime('%Y-%m-%d')

#url = "c:/skhynix_oms/" #파일이 담긴 폴더의 경로명
url = os.getcwd()
 # 위 폴더의 모든 파일을 리스트로
today_str = url+'/'+today_str
print(today_str)

if not os.path.exists(today_str):
    os.mkdir(today_str)

file_log = 'oms_log.xlsx'
file_check = 'oms_sbl_check.xlsx'

#device manage file
fn_dev_man = 'device_man.xlsx'

#format file
fn_format = 'skhynix_format.xlsx'

#list file
file_list = os.listdir(url)

file_list_arch1 = [file for file in file_list if fnmatch.fnmatch(file,'*xlsx*')]

if 'test2.xlsx' in file_list_arch1:
    file_list_arch1.remove('test2.xlsx')    


file_list_arch2 = [file for file in file_list if fnmatch.fnmatch(file,'*pdf*')]

file_list_xlsx = [file for file in file_list if fnmatch.fnmatch(file,'*yield*')]

#read log file
df_log = pd.read_excel(file_log, engine='openpyxl')
df_check_final = pd.read_excel(file_check, engine='openpyxl')

# read SB limit file
file_limit = 'SK Hynix limit file_20220907.xlsx'
df_limit = pd.read_excel(file_limit, engine='openpyxl')

#read device manage file
df_dev_man = pd.read_excel(fn_dev_man, engine='openpyxl')

#read format file

app = xw.App(visible=False)
wb_format = xw.Book(fn_format)
sh_format = xw.sheets[0]

df = pd.DataFrame([])
df_log_update = pd.DataFrame([])

for fn in file_list_xlsx:
    df1 = pd.read_excel(fn, engine='openpyxl')
    df = pd.concat([df, df1])

# start real change
#identify device
    
for device in df_dev_man['device part number']:
    df_device = df[df['Part']==device]
    df_device_prop = df_dev_man[df_dev_man['device part number']==device]
    #check data frame for devcie is empty
    if not df_device.empty:
        df_eval_prop1=eval(df_device_prop['sbl_str'].iloc[0])
        df_eval_prop2 = list(df_eval_prop1)
        df_device_buf=df_device[df_eval_prop2] 
        print(device)
        #print(df_device_prop)
        print(df_device_buf)
        len_device = len(df_device_buf)
        df_log_update[device] = len_device





#need to debug
df_log=pd.concat([df_log,df_log_update],columns=df_log.columns),axis=0)
df_log.to_excel(file_log,index=False)

#need to check this operation
#move column name
df_rcd_g3.columns = df_rcd_g3_ms.columns
df_rcd_g2.columns = df_rcd_g2_ms.columns
df_rcd.columns = df_rcd_ms.columns
df_spd_hub.columns = df_spdhub_ms.columns
df_ts.columns = df_ts_ms.columns
#PMIC D1
df_server_pmic_b1.columns = df_spmicbig1_ms.columns
df_server_pmic_b.columns = df_spmicbig_ms.columns
df_server_pmic_s.columns = df_spmicsmall_ms.columns
df_client_pmic.columns = df_cpmic_ms.columns


#process yield limit

sbl_latest_idx = df_limit.shape[1]-1

#df for SBL check result
df_sbl_check = pd.DataFrame([])
df_sbl_tmp = pd.DataFrame([]) 
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
#check client pmic limit matching
df_part_buf = df_client_pmic
df_limit_buf = df_limit_client_pmic
col_idx_ori=df_part_buf.columns.get_loc('YIELD')

for i in range (0,len_client_pmic):
    print("test client PMIC")
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl2
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]
    #print(SBL)
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB2']=df_part_buf.iloc[i,(col_idx_ori+1)]
        df_sbl_tmp['SB2 NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #set sbl4
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB4']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB4 NG'] = SBL
        sbl_error_detected = 1
    #set sbl7
    SBL=df_limit_buf.iloc[4,sbl_latest_idx]
    if df_part_buf.iloc[i,9] > SBL :
        df_sbl_tmp['SB7']=df_part_buf.iloc[i,col_idx_ori+4]
        df_sbl_tmp['SB7 NG'] = SBL
        sbl_error_detected = 1
    #set sbl8
    SBL=df_limit_buf.iloc[5,sbl_latest_idx]
    #print(SBL)
    if df_part_buf.iloc[i,10] > SBL :
        df_sbl_tmp['SB8']=df_part_buf.iloc[i,col_idx_ori+5]
        df_sbl_tmp['SB8 NG'] = SBL
        sbl_error_detected = 1
    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        #print(df_sbl_tmp)
        df_sbl_tmp['date'] = today_str
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        print("SBL error detected")
        print(df_sbl_check)
        #print(df_sbl_check)
    sbl_error_detected = 0

#check RCD Gen3 limit matching



df_sbl_tmp = pd.DataFrame([])
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
df_part_buf = df_rcd_g3
df_limit_buf = df_limit_rcd_g3
col_idx_ori=df_part_buf.columns.get_loc('YIELD')
for i in range (0,len_rcd_g3):
    print("RCD Gen3 check")
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori+0]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl2
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB2']=df_part_buf.iloc[i,col_idx_ori+1]
        df_sbl_tmp['SB2 NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #set sbl6
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB6']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB6 NG'] = SBL
        sbl_error_detected = 1
    #set sbl10
    SBL=df_limit_buf.iloc[4,sbl_latest_idx]
    if df_part_buf.iloc[i,9] > SBL :
        df_sbl_tmp['SB10']=df_part_buf.iloc[i,col_idx_ori+4]
        df_sbl_tmp['SB10 NG'] = SBL
        sbl_error_detected = 1
    #set sbl11
    SBL=df_limit_buf.iloc[5,sbl_latest_idx]
    if df_part_buf.iloc[i,10] > SBL :
        df_sbl_tmp['SB11']=df_part_buf.iloc[i,col_idx_ori+5]
        df_sbl_tmp['SB11 NG'] = SBL
        sbl_error_detected = 1
    #set sbl12
    SBL=df_limit_buf.iloc[6,sbl_latest_idx]
    if df_part_buf.iloc[i,11] > SBL :
        df_sbl_tmp['SB12']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB12 NG'] = SBL
        sbl_error_detected = 1        
    #set sbl13
    SBL=df_limit_buf.iloc[7,sbl_latest_idx]
    if df_part_buf.iloc[i,12] > SBL :
        df_sbl_tmp['SB13']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB13 NG'] = SBL
        sbl_error_detected = 1        
    #set sbl14
    SBL=df_limit_buf.iloc[8,sbl_latest_idx]
    if df_part_buf.iloc[i,13] > SBL :
        df_sbl_tmp['SB14']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB14 NG'] = SBL
        sbl_error_detected = 1        
    #set sbl15
    SBL=df_limit_buf.iloc[9,sbl_latest_idx]
    if df_part_buf.iloc[i,14] > SBL :
        df_sbl_tmp['SB15']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB15 NG'] = SBL
        sbl_error_detected = 1        


    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        df_sbl_tmp['date'] = today_str
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        print("SBL error detected")
        print(df_sbl_check)
    sbl_error_detected = 0

#check RCD Gen2 limit matching

df_sbl_tmp = pd.DataFrame([])
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
df_part_buf = df_rcd_g2
df_limit_buf = df_limit_rcd_g2
col_idx_ori=df_part_buf.columns.get_loc('YIELD')
for i in range (0,len_rcd_g2):
    print("RCD Gen2 check")
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori+0]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl2
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB2']=df_part_buf.iloc[i,col_idx_ori+1]
        df_sbl_tmp['SB2 NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #set sbl5
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB5']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB5 NG'] = SBL
        sbl_error_detected = 1
    #set sbl7
    SBL=df_limit_buf.iloc[4,sbl_latest_idx]
    if df_part_buf.iloc[i,9] > SBL :
        df_sbl_tmp['SB7']=df_part_buf.iloc[i,col_idx_ori+4]
        df_sbl_tmp['SB7 NG'] = SBL
        sbl_error_detected = 1
    #set sbl11
    SBL=df_limit_buf.iloc[5,sbl_latest_idx]
    if df_part_buf.iloc[i,10] > SBL :
        df_sbl_tmp['SB11']=df_part_buf.iloc[i,col_idx_ori+5]
        df_sbl_tmp['SB11 NG'] = SBL
        sbl_error_detected = 1
    #set sbl12
    SBL=df_limit_buf.iloc[6,sbl_latest_idx]
    if df_part_buf.iloc[i,11] > SBL :
        df_sbl_tmp['SB12']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB12 NG'] = SBL
        sbl_error_detected = 1
    #set sbl13
    SBL=df_limit_buf.iloc[7,sbl_latest_idx]
    if df_part_buf.iloc[i,12] > SBL :
        df_sbl_tmp['SB13']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB13 NG'] = SBL
        sbl_error_detected = 1        
    #set sbl14
    SBL=df_limit_buf.iloc[8,sbl_latest_idx]
    if df_part_buf.iloc[i,13] > SBL :
        df_sbl_tmp['SB14']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB14 NG'] = SBL
        sbl_error_detected = 1        
    #set sbl15
    SBL=df_limit_buf.iloc[9,sbl_latest_idx]
    if df_part_buf.iloc[i,14] > SBL :
        df_sbl_tmp['SB15']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB15 NG'] = SBL
        sbl_error_detected = 1        
        
    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        df_sbl_tmp['date'] = today_str
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        print("SBL error detected")
        print(df_sbl_check)
    sbl_error_detected = 0

#check RCD limit matching

df_sbl_tmp = pd.DataFrame([])
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
df_part_buf = df_rcd
df_limit_buf = df_limit_rcd
col_idx_ori=df_part_buf.columns.get_loc('YIELD')
for i in range (0,len_rcd):
    print("RCD check")
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori+0]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+1]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #set sbl5
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB5']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB5 NG'] = SBL
        sbl_error_detected = 1
    #set sbl7
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB7']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB7 NG'] = SBL
        sbl_error_detected = 1
    #set sbl9
    SBL=df_limit_buf.iloc[4,sbl_latest_idx]
    if df_part_buf.iloc[i,9] > SBL :
        df_sbl_tmp['SB9']=df_part_buf.iloc[i,col_idx_ori+4]
        df_sbl_tmp['SB9 NG'] = SBL
        sbl_error_detected = 1
    #set sbl10
    SBL=df_limit_buf.iloc[5,sbl_latest_idx]
    if df_part_buf.iloc[i,10] > SBL :
        df_sbl_tmp['SB10']=df_part_buf.iloc[i,col_idx_ori+5]
        df_sbl_tmp['SB10 NG'] = SBL
        sbl_error_detected = 1
    #set sbl12
    SBL=df_limit_buf.iloc[6,sbl_latest_idx]
    if df_part_buf.iloc[i,11] > SBL :
        df_sbl_tmp['SB12']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB12 NG'] = SBL
        sbl_error_detected = 1        
    #set sbl13
    SBL=df_limit_buf.iloc[7,sbl_latest_idx]
    if df_part_buf.iloc[i,12] > SBL :
        df_sbl_tmp['SB13']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB13 NG'] = SBL
        sbl_error_detected = 1        
    #set sbl14
    SBL=df_limit_buf.iloc[8,sbl_latest_idx]
    if df_part_buf.iloc[i,13] > SBL :
        df_sbl_tmp['SB14']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB14 NG'] = SBL
        sbl_error_detected = 1        
    #set sbl12
    SBL=df_limit_buf.iloc[9,sbl_latest_idx]
    if df_part_buf.iloc[i,14] > SBL :
        df_sbl_tmp['SB15']=df_part_buf.iloc[i,col_idx_ori+6]
        df_sbl_tmp['SB15 NG'] = SBL
        sbl_error_detected = 1        

    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        df_sbl_tmp['date'] = today_str
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        print("SBL error detected")
        print(df_sbl_check)
    sbl_error_detected = 0

#check SPH Hub limit matching

df_sbl_tmp = pd.DataFrame([])
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
df_part_buf = df_spd_hub
df_limit_buf = df_limit_spd_hub
col_idx_ori=df_part_buf.columns.get_loc('YIELD')
for i in range (0,len_spd_hub):
    print("test SPD Hub")
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori+0]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl2
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]*100
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB2']=df_part_buf.iloc[i,col_idx_ori+1]
        df_sbl_tmp['SB2 NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]*100
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #set sbl4
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]*100
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB4']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB4 NG'] = SBL
        sbl_error_detected = 1
    #set sbl6
    SBL=df_limit_buf.iloc[4,sbl_latest_idx]*100
    if df_part_buf.iloc[i,9] > SBL :
        df_sbl_tmp['SB6']=df_part_buf.iloc[i,col_idx_ori+4]
        df_sbl_tmp['SB6 NG'] = SBL
        sbl_error_detected = 1
    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        df_sbl_tmp['date'] = today_str
        #print(df_sbl_tmp)
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        #print(df_sbl_check)
        print("SBL error detected")
        print(df_sbl_check)
    sbl_error_detected = 0
#check ts limit matching

df_sbl_tmp = pd.DataFrame([])
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
df_part_buf = df_ts
df_limit_buf = df_limit_ts
col_idx_ori=df_part_buf.columns.get_loc('YIELD')
for i in range (0,len_ts):
    print("test TS")
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori+0]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl1
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB1']=df_part_buf.iloc[i,col_idx_ori+1]
        df_sbl_tmp['SB1 NG'] = SBL
        sbl_error_detected = 1
    #set sbl2
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB2']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB2 NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        df_sbl_tmp['date'] = today_str
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        print("SBL error detected")
        print(df_sbl_check)
    sbl_error_detected = 0

#check server pmic D1 limit matching
df_sbl_tmp = pd.DataFrame([])
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
df_part_buf = df_server_pmic_b1
df_limit_buf = df_limit_server_pmic_d1
col_idx_ori=df_part_buf.columns.get_loc('YIELD')
for i in range (0,len_server_pmic_b1):
    print("test big pmic D1")
#for i in range (0,1):
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori+0]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl2
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB2']=df_part_buf.iloc[i,col_idx_ori+1]
        df_sbl_tmp['SB2 NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #set sbl4
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB4']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB4 NG'] = SBL
        sbl_error_detected = 1
    #set sbl8
    SBL=df_limit_buf.iloc[4,sbl_latest_idx]
    if df_part_buf.iloc[i,9] > SBL :
        df_sbl_tmp['SB8']=df_part_buf.iloc[i,col_idx_ori+4]
        df_sbl_tmp['SB8 NG'] = SBL
        sbl_error_detected = 1
    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        df_sbl_tmp['date'] = today_str
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        print("SBL error detected")
        print(df_sbl_check)
    sbl_error_detected = 0

#check server pmic big limit matching

df_sbl_tmp = pd.DataFrame([])
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
df_part_buf = df_server_pmic_b
df_limit_buf = df_limit_server_pmic
col_idx_ori=df_part_buf.columns.get_loc('YIELD')
for i in range (0,len_server_pmic_b):
    print("test big pmic")
#for i in range (0,1):
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori+0]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl2
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB2']=df_part_buf.iloc[i,col_idx_ori+1]
        df_sbl_tmp['SB2 NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #set sbl4
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB4']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB4 NG'] = SBL
        sbl_error_detected = 1
    #set sbl8
    SBL=df_limit_buf.iloc[4,sbl_latest_idx]
    if df_part_buf.iloc[i,9] > SBL :
        df_sbl_tmp['SB8']=df_part_buf.iloc[i,col_idx_ori+4]
        df_sbl_tmp['SB8 NG'] = SBL
        sbl_error_detected = 1
    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        df_sbl_tmp['date'] = today_str
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        print("SBL error detected")
        print(df_sbl_check)
    sbl_error_detected = 0
#check server pmic small limit matching

df_sbl_tmp = pd.DataFrame([])
sbl_error_detected = 0
df_part_buf = pd.DataFrame([])
df_limit_buf = pd.DataFrame([])
df_part_buf = df_server_pmic_s
df_limit_buf = df_limit_server_pmic
col_idx_ori=df_part_buf.columns.get_loc('YIELD')
for i in range (0,len_server_pmic_s):
    print("test small pmic")
    #check yield
    df_sbl_tmp = df_part_buf[['Part','Asm_lot_num','Datecode']][i:i+1]
    #set Yield
    SBL=df_limit_buf.iloc[0,sbl_latest_idx]*100
    if df_part_buf.iloc[i,5] < SBL :
        df_sbl_tmp['YIELD']=df_part_buf.iloc[i,col_idx_ori+0]
        df_sbl_tmp['Yield limit NG'] = SBL
        sbl_error_detected = 1
    #set sbl2
    SBL=df_limit_buf.iloc[1,sbl_latest_idx]
    if df_part_buf.iloc[i,6] > SBL :
        df_sbl_tmp['SB2']=df_part_buf.iloc[i,col_idx_ori+1]
        df_sbl_tmp['SB2 NG'] = SBL
        sbl_error_detected = 1
    #set sbl3
    SBL=df_limit_buf.iloc[2,sbl_latest_idx]
    if df_part_buf.iloc[i,7] > SBL :
        df_sbl_tmp['SB3']=df_part_buf.iloc[i,col_idx_ori+2]
        df_sbl_tmp['SB3 NG'] = SBL
        sbl_error_detected = 1
    #set sbl4
    SBL=df_limit_buf.iloc[3,sbl_latest_idx]
    if df_part_buf.iloc[i,8] > SBL :
        df_sbl_tmp['SB4']=df_part_buf.iloc[i,col_idx_ori+3]
        df_sbl_tmp['SB4 NG'] = SBL
        sbl_error_detected = 1
     #set sbl8
    SBL=df_limit_buf.iloc[4,sbl_latest_idx]
    if df_part_buf.iloc[i,9] > SBL :
        df_sbl_tmp['SB8']=df_part_buf.iloc[i,col_idx_ori+4]
        df_sbl_tmp['SB8 NG'] = SBL
        sbl_error_detected = 1
    #check datecode
    if not (str(df_part_buf.iloc[i,2])[0] == '2') or not(len(str(df_part_buf.iloc[i,2])) ==4):
        df_sbl_tmp['Datecode NG'] = 'O'
        sbl_error_detected = 1
    if sbl_error_detected:
        df_sbl_tmp['date'] = today_str
        df_sbl_check = pd.concat([df_sbl_check,df_sbl_tmp])
        print("SBL error detected")
        print(df_sbl_check)
    sbl_error_detected = 0


df_check_final= pd.concat([df_check_final,df_sbl_check])
df_check_final.to_excel(file_check,index=False)

#merge master file and updated file
df_rcd_g3 = pd.concat([df_rcd_g3_ms,df_rcd_g3])
df_rcd_g2 = pd.concat([df_rcd_g2_ms,df_rcd_g2])
df_rcd = pd.concat([df_rcd_ms,df_rcd])
df_spd_hub = pd.concat([df_spdhub_ms,df_spd_hub])
df_ts = pd.concat([df_ts_ms,df_ts])

df_server_pmic_b1 = pd.concat([df_spmicbig1_ms,df_server_pmic_b1])
df_server_pmic_b = pd.concat([df_spmicbig_ms,df_server_pmic_b])
df_server_pmic_s = pd.concat([df_spmicsmall_ms,df_server_pmic_s])
df_client_pmic = pd.concat([df_cpmic_ms,df_client_pmic])




#reformat 
#Gen3 RCD

df_rcd_g3_rs=df_rcd_g3.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_rcd_g3_rs = df_rcd_g3_rs.dropna(subset=['Part'])
df_rcd_g3_rs['Material Gr.'] = mg_rcd_g3
df_rcd_g3_rs['Sapcode'] = sap_rcd_g3
df_rcd_g3_rs[['USL','LSL','Unit']] = [0,0,'%']
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='YIELD','USL']=df_limit_rcd_g3.iloc[0,sbl_latest_idx]*100
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='YIELD','USL'] = '4sigmas'

df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB2(SCAN)','USL']=df_limit_rcd_g3.iloc[1,sbl_latest_idx]
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB3(Open/short)','USL']=df_limit_rcd_g3.iloc[2,sbl_latest_idx]
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB6(Power Short)','USL']=df_limit_rcd_g3.iloc[3,sbl_latest_idx]
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB10(Function)','USL']=df_limit_rcd_g3.iloc[4,sbl_latest_idx]
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB11(Leakage_DRST)','USL']=df_limit_rcd_g3.iloc[5,sbl_latest_idx]
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB12(Leakages_Input_Post)','USL']=df_limit_rcd_g3.iloc[6,sbl_latest_idx]
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB13(Leakages_Output_Post)','USL']=df_limit_rcd_g3.iloc[7,sbl_latest_idx]
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB14(Leakages_Input_Pre)','USL']=df_limit_rcd_g3.iloc[8,sbl_latest_idx]
df_rcd_g3_rs.loc[df_rcd_g3_rs['variable']=='SB15(Leakages_Output_Pre)','USL']=df_limit_rcd_g3.iloc[9,sbl_latest_idx]

#Gen2 RCD
df_rcd_g2_rs=df_rcd_g2.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_rcd_g2_rs = df_rcd_g2_rs.dropna(subset=['Part'])
df_rcd_g2_rs['Material Gr.'] = mg_rcd_g2
df_rcd_g2_rs['Sapcode'] = sap_rcd_g2
df_rcd_g2_rs[['USL','LSL','Unit']] = [0,0,'%']
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='YIELD','USL']=df_limit_rcd_g2.iloc[0,sbl_latest_idx]*100
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='YIELD','USL'] = '4sigmas'

df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB2(SCAN)','USL']=df_limit_rcd_g2.iloc[1,sbl_latest_idx]
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB3(Open)','USL']=df_limit_rcd_g2.iloc[2,sbl_latest_idx]
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB5(Short)','USL']=df_limit_rcd_g2.iloc[3,sbl_latest_idx]
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB7(IDD6S)','USL']=df_limit_rcd_g2.iloc[4,sbl_latest_idx]
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB11(Function)','USL']=df_limit_rcd_g2.iloc[5,sbl_latest_idx]
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB12(Leakages_Input_post)','USL']=df_limit_rcd_g2.iloc[6,sbl_latest_idx]
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB13(Leakages_Output_post)','USL']=df_limit_rcd_g2.iloc[7,sbl_latest_idx]
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB14(Leakages_Input_pre)','USL']=df_limit_rcd_g2.iloc[8,sbl_latest_idx]
df_rcd_g2_rs.loc[df_rcd_g2_rs['variable']=='SB15(Leakages_Output_pre)','USL']=df_limit_rcd_g2.iloc[9,sbl_latest_idx]


#Gen1 RCD
df_rcd_rs=df_rcd.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_rcd_rs = df_rcd_rs.dropna(subset=['Part'])
df_rcd_rs['Material Gr.'] = mg_rcd
df_rcd_rs['Sapcode'] = sap_rcd
df_rcd_rs[['USL','LSL','Unit']] = [0,0,'%']

df_rcd_rs.loc[df_rcd_rs['variable']=='YIELD','USL']=df_limit_rcd.iloc[0,sbl_latest_idx]*100
df_rcd_rs.loc[df_rcd_rs['variable']=='YIELD','USL']='4sigmas'
df_rcd_rs.loc[df_rcd_rs['variable']=='SB3(Open)','USL']=df_limit_rcd.iloc[1,sbl_latest_idx]
df_rcd_rs.loc[df_rcd_rs['variable']=='SB5(Short)','USL']=df_limit_rcd.iloc[2,sbl_latest_idx]
df_rcd_rs.loc[df_rcd_rs['variable']=='SB7(IDD6S)','USL']=df_limit_rcd.iloc[3,sbl_latest_idx]
df_rcd_rs.loc[df_rcd_rs['variable']=='SB9(Function)','USL']=df_limit_rcd.iloc[4,sbl_latest_idx]
df_rcd_rs.loc[df_rcd_rs['variable']=='SB10(SCAN)','USL']=df_limit_rcd.iloc[5,sbl_latest_idx]
df_rcd_rs.loc[df_rcd_rs['variable']=='SB12(Leakages_Input_Post)','USL']=df_limit_rcd.iloc[6,sbl_latest_idx]
df_rcd_rs.loc[df_rcd_rs['variable']=='SB13(Leakages_Output_Post)','USL']=df_limit_rcd.iloc[7,sbl_latest_idx]
df_rcd_rs.loc[df_rcd_rs['variable']=='SB14(Leakages_Input_Pre)','USL']=df_limit_rcd.iloc[8,sbl_latest_idx]
df_rcd_rs.loc[df_rcd_rs['variable']=='SB15(Leakages_Output_Pre)','USL']=df_limit_rcd.iloc[9,sbl_latest_idx]

#SPD Hub
df_spd_hub_rs=df_spd_hub.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_spd_hub_rs = df_spd_hub_rs.dropna(subset=['Part'])
df_spd_hub_rs['Material Gr.'] = mg_spd_hub
df_spd_hub_rs['Sapcode'] = sap_spd_hub
df_spd_hub_rs[['USL','LSL','Unit']] = [0,0,'%']

df_spd_hub_rs.loc[df_spd_hub_rs['variable']=='YIELD','USL']=df_limit_spd_hub.iloc[0,sbl_latest_idx]*100
df_spd_hub_rs.loc[df_spd_hub_rs['variable']=='YIELD','USL']='4sigmas'
df_spd_hub_rs.loc[df_spd_hub_rs['variable']=='SB2(Open)','USL']=df_limit_spd_hub.iloc[1,sbl_latest_idx]
df_spd_hub_rs.loc[df_spd_hub_rs['variable']=='SB3(Short)','USL']=df_limit_spd_hub.iloc[2,sbl_latest_idx]
df_spd_hub_rs.loc[df_spd_hub_rs['variable']=='SB4(Leakage)','USL']=df_limit_spd_hub.iloc[3,sbl_latest_idx]
df_spd_hub_rs.loc[df_spd_hub_rs['variable']=='SB6(Function)','USL']=df_limit_spd_hub.iloc[4,sbl_latest_idx]

#TS
df_ts_rs=df_ts.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_ts_rs = df_ts_rs.dropna(subset=['Part'])
df_ts_rs['Material Gr.'] = mg_ts
df_ts_rs['Sapcode'] = sap_ts
df_ts_rs[['USL','LSL','Unit']] = [0,0,'%']

df_ts_rs.loc[df_ts_rs['variable']=='YIELD','USL']=df_limit_ts.iloc[0,sbl_latest_idx]*100
df_ts_rs.loc[df_ts_rs['variable']=='YIELD','USL']='4sigmas'
df_ts_rs.loc[df_ts_rs['variable']=='SB1(Cont.)','USL']=df_limit_ts.iloc[1,sbl_latest_idx]
df_ts_rs.loc[df_ts_rs['variable']=='SB2(Leakage)','USL']=df_limit_ts.iloc[2,sbl_latest_idx]
df_ts_rs.loc[df_ts_rs['variable']=='SB3(Function)','USL']=df_limit_ts.iloc[3,sbl_latest_idx]

#PMIC D1
df_server_pmic_b1_rs=df_server_pmic_b1.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_server_pmic_b1_rs = df_server_pmic_b1_rs.dropna(subset=['Part'])
df_server_pmic_b1_rs['Material Gr.'] = mg_server_pmic_b1
df_server_pmic_b1_rs['Sapcode'] = sap_server_pmic_b1
df_server_pmic_b1_rs[['USL','LSL','Unit']] = [0,0,'%']

df_server_pmic_b1_rs.loc[df_server_pmic_b1_rs['variable']=='YIELD','USL']=df_limit_server_pmic_d1.iloc[0,sbl_latest_idx]*100
df_server_pmic_b1_rs.loc[df_server_pmic_b1_rs['variable']=='YIELD','USL']='4sigmas'
df_server_pmic_b1_rs.loc[df_server_pmic_b1_rs['variable']=='SB2(Open)','USL']=df_limit_server_pmic_d1.iloc[1,sbl_latest_idx]
df_server_pmic_b1_rs.loc[df_server_pmic_b1_rs['variable']=='SB3(Short)','USL']=df_limit_server_pmic_d1.iloc[2,sbl_latest_idx]
df_server_pmic_b1_rs.loc[df_server_pmic_b1_rs['variable']=='SB4(SCAN)','USL']=df_limit_server_pmic_d1.iloc[3,sbl_latest_idx]
df_server_pmic_b1_rs.loc[df_server_pmic_b1_rs['variable']=='SB8(Leakage)','USL']=df_limit_server_pmic_d1.iloc[4,sbl_latest_idx]

df_server_pmic_b_rs=df_server_pmic_b.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_server_pmic_b_rs = df_server_pmic_b_rs.dropna(subset=['Part'])
df_server_pmic_b_rs['Material Gr.'] = mg_server_pmic_b
df_server_pmic_b_rs['Sapcode'] = sap_server_pmic_b
df_server_pmic_b_rs[['USL','LSL','Unit']] = [0,0,'%']

df_server_pmic_b_rs.loc[df_server_pmic_b_rs['variable']=='YIELD','USL']=df_limit_server_pmic.iloc[0,sbl_latest_idx]*100
df_server_pmic_b_rs.loc[df_server_pmic_b_rs['variable']=='YIELD','USL']='4sigmas'
df_server_pmic_b_rs.loc[df_server_pmic_b_rs['variable']=='SB2(Open)','USL']=df_limit_server_pmic.iloc[1,sbl_latest_idx]
df_server_pmic_b_rs.loc[df_server_pmic_b_rs['variable']=='SB3(Short)','USL']=df_limit_server_pmic.iloc[2,sbl_latest_idx]
df_server_pmic_b_rs.loc[df_server_pmic_b_rs['variable']=='SB4(SCAN)','USL']=df_limit_server_pmic.iloc[3,sbl_latest_idx]
df_server_pmic_b_rs.loc[df_server_pmic_b_rs['variable']=='SB8(Leakage)','USL']=df_limit_server_pmic.iloc[4,sbl_latest_idx]


df_server_pmic_s_rs=df_server_pmic_s.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_server_pmic_s_rs = df_server_pmic_s_rs.dropna(subset=['Part'])
df_server_pmic_s_rs['Material Gr.'] = mg_server_pmic_s
df_server_pmic_s_rs['Sapcode'] = sap_server_pmic_s
df_server_pmic_s_rs[['USL','LSL','Unit']] = [0,0,'%']

df_server_pmic_s_rs.loc[df_server_pmic_s_rs['variable']=='YIELD','USL']=df_limit_server_pmic.iloc[0,sbl_latest_idx]*100
df_server_pmic_s_rs.loc[df_server_pmic_s_rs['variable']=='YIELD','USL']='4sigmas'
df_server_pmic_s_rs.loc[df_server_pmic_s_rs['variable']=='SB2(Open)','USL']=df_limit_server_pmic.iloc[1,sbl_latest_idx]
df_server_pmic_s_rs.loc[df_server_pmic_s_rs['variable']=='SB3(Short)','USL']=df_limit_server_pmic.iloc[2,sbl_latest_idx]
df_server_pmic_s_rs.loc[df_server_pmic_s_rs['variable']=='SB4(SCAN)','USL']=df_limit_server_pmic.iloc[3,sbl_latest_idx]
df_server_pmic_s_rs.loc[df_server_pmic_s_rs['variable']=='SB8(Leakage)','USL']=df_limit_server_pmic.iloc[4,sbl_latest_idx]



df_client_pmic_rs=df_client_pmic.melt(id_vars=['Part','Asm_lot_num','Datecode','Picked_qty','SO','SHIP_date','COO','PO_number'])
df_client_pmic_rs = df_client_pmic_rs.dropna(subset=['Part'])
df_client_pmic_rs['Material Gr.'] = mg_client_pmic
df_client_pmic_rs['Sapcode'] = sap_client_pmic
df_client_pmic_rs[['USL','LSL','Unit']] = [0,0,'%']

df_client_pmic_rs.loc[df_client_pmic_rs['variable']=='YIELD','USL']=df_limit_client_pmic.iloc[0,sbl_latest_idx]*100
df_client_pmic_rs.loc[df_client_pmic_rs['variable']=='YIELD','USL']='4sigmas'
df_client_pmic_rs.loc[df_client_pmic_rs['variable']=='SB2(Open)','USL']=df_limit_client_pmic.iloc[1,sbl_latest_idx]
df_client_pmic_rs.loc[df_client_pmic_rs['variable']=='SB3(Short)','USL']=df_limit_client_pmic.iloc[2,sbl_latest_idx]
df_client_pmic_rs.loc[df_client_pmic_rs['variable']=='SB4(SCAN)','USL']=df_limit_client_pmic.iloc[3,sbl_latest_idx]
df_client_pmic_rs.loc[df_client_pmic_rs['variable']=='SB7(IDD)','USL']=df_limit_client_pmic.iloc[4,sbl_latest_idx]
df_client_pmic_rs.loc[df_client_pmic_rs['variable']=='SB8(Leakage)','USL']=df_limit_client_pmic.iloc[5,sbl_latest_idx]

#merget data
df_all=pd.concat([df_rcd_g3_rs,df_rcd_g2_rs,df_rcd_rs,df_spd_hub_rs,df_ts_rs,df_server_pmic_b_rs,df_server_pmic_b1_rs,df_server_pmic_s_rs,df_client_pmic_rs])

df_all_1=df_all.drop(['Datecode','Picked_qty','SO','COO','PO_number'],axis=1)
df_all_1['BP']='Reensas'
df_all_1['Pass/Fail'] ='=IF([@Value]="","-",IF(AND([@USL]="",[@LSL]=""),"P",IF([@USL]="",IF([@Value]<[@LSL],"F","P"),IF([@LSL]="",IF([@Value]>[@USL],"F","P"),IF([@Value]>[@USL],"F",IF([@Value]<[@LSL],"F","P"))))))'
df_all_2 = df_all_1[['BP','Material Gr.','Part','Sapcode','Asm_lot_num','variable','Unit','USL','LSL','value','Pass/Fail','SHIP_date']]
df_all_2=df_all_2.rename(columns={'Part':'Material Name','Asm_lot_num':'Lot No','variable':'Item','SHIP_date':'Registration Date'})
df_all_2.to_excel(excel_writer="test2.xlsx",index=False)

with pd.ExcelWriter(file_list_master[0]) as writer:
    df_rcd_g2.to_excel(writer,sheet_name=rcd_g2,index=False)
    df_rcd.to_excel(writer,sheet_name=rcd,index=False)
    df_spd_hub.to_excel(writer,sheet_name=spd_hub,index=False)
    df_ts.to_excel(writer,sheet_name=ts,index=False)
    df_server_pmic_b.to_excel(writer,sheet_name=server_pmic_b,index=False)
    df_server_pmic_b1.to_excel(writer,sheet_name=server_pmic_b1,index=False)

    df_server_pmic_s.to_excel(writer,sheet_name=server_pmic_s,index=False)
    df_client_pmic.to_excel(writer,sheet_name=client_pmic,index=False)
    
#os.mkdir(today_str)

# move PDF file
for g in file_list_arch2:
    shutil.move(url+ g, today_str)
# move excel files
for g in file_list_arch1:
    shutil.move(url+ g, today_str)

# get back master file
shutil.move(today_str + '/'+file_list_master[0] , url)

#get back log file
shutil.move(today_str + '/'+file_log , url)

#get back log file
shutil.move(today_str + '/'+file_check , url)

#get limit file
shutil.move(today_str + '/'+file_limit , url)

