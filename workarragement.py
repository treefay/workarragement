import pandas as pd
import datetime as dt
import os
import platform as pt

#CRCA
#ARBPL 工作中心
#KAPAZ 每日最长工作时间

#MARA
#MATNR 物料编码
#MATKL 物料组

#PLAF
#PLNUM 计划订单号
#MATNR 物料编码
#PLWRK 生产工厂
#GSMNG 计划订单数量
#PSTTR 订单开始日期
#PEDTR 订单完成日期

#PLPO
#MATNR 物料编码
#ARBPL 工作中心
#TIME  单位物料生产时间

#需求
#PLAF+
#PLNUM 计划订单号
#MATNR 物料编码
#PLWRK 生产工厂
#GSMNG 计划订单数量
#PSTTR 订单开始日期
#PEDTR 订单完成日期
#ARBPL 工作中心
#MATKL 物料组
#WKBEGIN 工作开始时间
#WKEND   工作结束时间

def check_plaf_file(vPLAF):
    #检查打开的plaf文件是否为空，为空则抛出异常
    #检查打开的plaf文件是否包含应有列，缺少列则抛出异常
    pass

def check_mara_file(vMARA):
    #检查打开的mara文件是否为空，为空则抛出异常
    #检查打开的mara文件是否包含应有列，缺少列则抛出异常
    pass

def decoOUTPUT(func):
    def starline(*args,**kwargs):
        print("***************************************************")
        func(*args,**kwargs)
        print("***************************************************")
    return starline

def getworktime(wtime):
    work_date_str=dt.datetime.strftime(wtime,'%Y/%m/%d')
    eveyday_work_start_time_str=work_date_str+" 08:00:00"
    eveyday_work_start_time=dt.datetime.strptime(eveyday_work_start_time_str,'%Y/%m/%d %H:%M:%S')
    return eveyday_work_start_time

def checkArragementStartTime(stime):
    everyday_workstart=getworktime(stime)
    if everyday_workstart >= stime:
        return everyday_workstart
    else:
        return stime

def inputWorkStartDate():
    print("***************************************************")
    input_date=input("Enter the Work arrangement start date,format %Y/%m/%d %H:%M:%S like '2020/02/01 08:00:00':")
    try:
        st_date=dt.datetime.strptime(input_date,'%Y/%m/%d %H:%M:%S')
        return st_date,1
    except:
        print("The date format your inputed is not correct ,Please input it again")
        return "error",0

def checkResult(result,errorinfo):
    if result == 0:
        print("Error:"+errorinfo)
        exit(1)

def dropDFColumn(df,dlist):
    for ind in dlist:
        df = df.drop(ind,axis=1)
    return df

def checkFileExistsAndRead(vfile):
    if os.access(vfile, os.R_OK) == False :
        print("Error:File ",vfile," not exists!")
        return 0,"nofile"
    else:
        print("Processing : File Check ", vfile, " exists!")
        rfile = pd.read_csv(vfile)
        return 1,rfile

def openFileInput():
    print("***************************************************")
    open_file_current_dir=input("IF all files are in the same folder with workarrangement.py script ,Please enter 'Y'   \n"
                                "else enter 'N' for entering the absolute directory for each file  \n"
                                "Input :")
    if open_file_current_dir == "Y" or open_file_current_dir == "y":
        sys_version = pt.system()
        current_dir = os.path.abspath('.')
        if sys_version == "Windows":
            print("Notice: OS is Windows!!!")
            c_plaf = current_dir + '\\PLAF.csv'
            c_plpo = current_dir + '\\PLPO.csv'
            c_mara = current_dir + '\\MARA.csv'
            c_crca = current_dir + '\\CRCA.csv'
        elif sys_version == "Linux":
            print("Notice: OS is Linux!!!")
            c_plaf = current_dir + '/PLAF.csv'
            c_plpo = current_dir + '/PLPO.csv'
            c_mara = current_dir + '/MARA.csv'
            c_crca = current_dir + '/CRCA.csv'
        elif sys_version == "Darwin":
            print("Notice: OS is Darwin!!!")
            c_plaf = current_dir + '/PLAF.csv'
            c_plpo = current_dir + '/PLPO.csv'
            c_mara = current_dir + '/MARA.csv'
            c_crca = current_dir + '/CRCA.csv'
        else:
            print("Error: system ",sys_version," is not supported. \n"
                  "please rerun the script and enter the absolute directory for each file")
            exit(1)
        open_flag1,o_plaf=checkFileExistsAndRead(c_plaf)
        open_flag2,o_plpo=checkFileExistsAndRead(c_plpo)
        open_flag3,o_mara=checkFileExistsAndRead(c_mara)
        open_flag4,o_crca=checkFileExistsAndRead(c_crca)
        if open_flag1 * open_flag2 * open_flag3 * open_flag4 == 0:
            print("Error: Some file is not in the same folder with workarrangement.py script \n"
                  "please re-run the script and enter the absolute directory for each file")
            return 0,"nofile","nofile","nofile","nofile"
        else:
            return 1,o_plaf,o_plpo,o_mara,o_crca
    elif open_file_current_dir == "N" or open_file_current_dir == "n":
        input_plaf = input("PLAF :")
        open_flag1,o_plaf=checkFileExistsAndRead(input_plaf)
        checkResult(open_flag1,"Read file PLAF.csv failed.")

        input_plpo = input("PLPO :")
        open_flag2,o_plpo=checkFileExistsAndRead(input_plpo)
        checkResult(open_flag2,"Read file PLPO.csv failed.")

        input_mara = input("MARA :")
        open_flag3,o_mara=checkFileExistsAndRead(input_mara)
        checkResult(open_flag3,"Read file MARA.csv failed.")

        input_crca = input("CRCA :")
        open_flag4,o_crca=checkFileExistsAndRead(input_crca)
        checkResult(open_flag4,"Read file CRCA.csv failed.")

        return 1, o_plaf, o_plpo, o_mara, o_crca
    else:
        print("Error :Your input is neither 'Y' nor 'N' .\n"
              "please re-run the script and enter the absolute directory for each file")
        exit(1)

def inputOutputFileDir():
    print("***************************************************")
    odir = input("Enter the output file directory :")
    ofile = outputFile(odir)
    return ofile

def outputFile(odir):
    current_dir = os.path.abspath('.')
    now = dt.datetime.now()
    sys_version = pt.system()
    ofile = "workarrangement" + now.strftime("%Y%m%d%H%M%S") + ".csv"
    if os.path.isdir(odir) == False or os.access(odir,os.W_OK) == False:
        print("Error:",odir,"not exists or have no write privilege.\n"
                            "Arrangement file will be output to ",current_dir)
        if sys_version == "Windows":
            ofile = current_dir + "\\" + ofile
        elif sys_version == "Linux":
            ofile = current_dir + "/" + ofile
        elif sys_version == "Darwin":
            ofile = current_dir + "/" + ofile
        else:
            print("Error: system ",sys_version," is not supported. \n")
            exit(1)
    else:
        if sys_version == "Windows":
            ofile = odir + "\\" + ofile
        elif sys_version == "Linux":
            ofile = odir + "/" + ofile
        elif sys_version == "Darwin":
            ofile = odir + "/" + ofile
        else:
            print("Error: system ", sys_version, " is not supported. \n")
            exit(1)
    print("Output file name : ",ofile)
    return ofile

def workArrangement(plaf,plpo,mara,crca,wkbg_date):
    # 初始化输出文件
    plaf_output = pd.DataFrame()
    # yyyy/mm/dd转换成日期
    plaf['PEDTR'] = pd.to_datetime(plaf['PEDTR'], format='%Y/%m/%d')
    plaf['PSTTR'] = pd.to_datetime(plaf['PSTTR'], format='%Y/%m/%d')
    plaf['GSMNG'] = plaf['GSMNG'].astype(float)
    try:
        plaf_plus = pd.merge(plaf, mara, on='MATNR', how='outer')
    except:
        print("Error: Merge plaf and mara failed.")
        exit(1)

    # 产生物料组每日产出量MATPD
    try:
        plca = pd.merge(plpo, crca, on='ARBPL', how='outer')
    except:
        print("Error: Merge plpo and crca failed.")
        exit(1)

    try:
        plaf_plus = pd.merge(plaf_plus, plca, on='MATNR', how='outer')
    except:
        print("Error: Merge plaf_plus and plca failed.")
        exit(1)

    plaf_plus['GSMNG'] = plaf_plus['GSMNG'].astype(float)

    # 由于工作中心为并行工作，故按照工作中心分别排期

    ARBPL_DIS = list(plaf_plus['ARBPL'].drop_duplicates())
    for ind in ARBPL_DIS:
        sub_plaf = plaf_plus[(plaf_plus['ARBPL'] == ind)].sort_values(by=['PEDTR', 'PSTTR', 'MATKL'])
        next_begin_date = wkbg_date
        sub_plaf['begin_time'] = next_begin_date
        sub_plaf['end_time'] = next_begin_date
        sub_plaf['delay'] = None
        for index, row in sub_plaf.iterrows():
            sub_plaf.at[index, 'begin_time'] = next_begin_date
            work_used_day = row['GSMNG'] * row['TIME'] // row['KAPAZ']
            work_used_hour = row['GSMNG'] * row['TIME'] % row['KAPAZ']
            everyday_workstart = getworktime(sub_plaf.at[index, 'begin_time'])
            sub_plaf.at[index, 'end_time'] = sub_plaf.at[index, 'begin_time'] + dt.timedelta(
                hours=work_used_hour) + dt.timedelta(days=work_used_day)

            if sub_plaf.at[index, 'begin_time'] + dt.timedelta(
                    hours=work_used_hour) > everyday_workstart + dt.timedelta(hours=row['KAPAZ']):
                work_used_day = work_used_day + 1
                everyday_workend = everyday_workstart + dt.timedelta(hours=row['KAPAZ'])
                work_time_left = getworktime(sub_plaf.at[index, 'begin_time']) + dt.timedelta(
                    hours=work_used_hour) - everyday_workend
                sub_plaf.at[index, 'end_time'] = sub_plaf.at[index, 'begin_time'] + dt.timedelta(
                    days=work_used_day) + work_time_left
            else:
                sub_plaf.at[index, 'end_time'] = sub_plaf.at[index, 'begin_time'] + dt.timedelta(
                    days=work_used_day) + dt.timedelta(hours=work_used_hour)

            if sub_plaf.at[index, 'end_time'] > row['PEDTR']:
                sub_plaf.at[index, 'delay'] = 'Yes'
            else:
                sub_plaf.at[index, 'delay'] = 'No'
            next_begin_date = sub_plaf.at[index, 'end_time']
        plaf_output = plaf_output.append(sub_plaf)

    # 删除多余列
    plaf_output = dropDFColumn(plaf_output, ['TIME', 'KAPAZ', 'MATKL'])
    # 创建输出excel
    try:
        plaf_output.to_csv(output_file)
    except:
        return 0
    return 1


if __name__ == '__main__':

    #读取四种csv
    open_flag=0
    while open_flag == 0:
        open_flag,plaf,plpo,mara,crca=openFileInput()

    #plaf = pd.read_csv('/Users/treefay/Documents/mem/python/workarragement/PLAF.csv')
    #mara= pd.read_csv('/Users/treefay/Documents/mem/python/workarragement/MARA.csv')
    #plpo = pd.read_csv('/Users/treefay/Documents/mem/python/workarragement/PLPO.csv')
    #crca = pd.read_csv('/Users/treefay/Documents/mem/python/workarragement/CRCA.csv')

    output_file=inputOutputFileDir()

    result_code=0
    while result_code == 0:
        wkbg_date,result_code = inputWorkStartDate()
        checkResult(result_code,"Notice:Work arrangement start date input error.\n"
                                "enter 'ctrl + D' to stop the script.")

    wkbg_date = checkArragementStartTime(wkbg_date)
    print("Notice：Work will be arranged on date: ",wkbg_date)

    result_code=workArrangement(plaf, plpo, mara, crca, wkbg_date)
    if result_code == 1:
        print("Success: Work arrangement run succeed.")
    else:
        print("Error: Work arrangement run failed.")
