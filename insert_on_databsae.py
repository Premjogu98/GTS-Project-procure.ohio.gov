import os
import sys
import time
from datetime import datetime

import mysql.connector

import global_var


def Local_connection():
    a = 0
    while a == 0:
        try:
            File_Location = open("D:\\0 PYTHON EXE SQL CONNECTION & DRIVER PATH\\procure.ohio.gov\\Location For Database & Driver.txt", "r")
            TXT_File_AllText = File_Location.read()
            Local_host = str(TXT_File_AllText).partition("Local_host=")[2].partition(",")[0].strip()
            Local_user = str(TXT_File_AllText).partition("Local_user=")[2].partition(",")[0].strip()
            Local_password = str(TXT_File_AllText).partition("Local_password=")[2].partition(",")[0].strip()
            Local_db = str(TXT_File_AllText).partition("Local_db=")[2].partition(",")[0].strip()
            Local_charset = str(TXT_File_AllText).partition("Local_charset=")[2].partition("\")")[0].strip()
            mydb_Local = mysql.connector.connect(
                host=str(Local_host) ,
                user=str(Local_user) ,
                passwd=str(Local_password) ,
                database=str(Local_db) ,
                charset=str(Local_charset))
            print('SQL Connected Local_connection')
            a = 1
            return mydb_Local
        except mysql.connector.ProgrammingError as e:
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname ,
                  "\n" , exc_tb.tb_lineno)
            a = 0
            time.sleep(10)


def L2L_connection():
    a3 = 0
    while a3 == 0:
        try:
            File_Location = open("D:\\0 PYTHON EXE SQL CONNECTION & DRIVER PATH\\procure.ohio.gov\\Location For Database & Driver.txt" , "r")
            TXT_File_AllText = File_Location.read()

            L2L_host = str(TXT_File_AllText).partition("L2L_host=")[2].partition(",")[0].strip()
            L2L_user = str(TXT_File_AllText).partition("L2L_user=")[2].partition(",")[0].strip()
            L2L_password = str(TXT_File_AllText).partition("L2L_password=")[2].partition(",")[0].strip()
            L2L_db = str(TXT_File_AllText).partition("L2L_db=")[2].partition(",")[0].strip()
            L2L_charset = str(TXT_File_AllText).partition("L2L_charset=")[2].partition("\")")[0].strip()

            mydb_L2L = mysql.connector.connect(
                host=str(L2L_host) ,
                user=str(L2L_user) ,
                passwd=str(L2L_password) ,
                database=str(L2L_db) ,
                charset=str(L2L_charset))
            print('SQL Connected L2L_connection')
            print("Live Tender")
            a3 = 1
            return mydb_L2L
        except mysql.connector.ProgrammingError as e:
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname ,
                  "\n" , exc_tb.tb_lineno)
            time.sleep(10)
            a3 = 0


def check_Duplication(get_htmlSource, SegFeild):
    mydb_Local = Local_connection()
    mycursorLocal = mydb_Local.cursor()
    a1 = 0
    while a1 == 0:
        try:
            if SegFeild[13] != '' and SegFeild[24] != '' and SegFeild[7] != '':
                commandText = "SELECT Posting_Id from Tenders where tender_notice_no = '" + str(SegFeild[13]) + "' and Country = '" + str(SegFeild[7]) + "' and doc_last= '" + str(SegFeild[24]) + "'"
            elif SegFeild[13] != "" and SegFeild[7] != "":
                commandText = "SELECT Posting_Id from Tenders where tender_notice_no = '" + str(SegFeild[13]) + "' and Country = '" + str(SegFeild[7]) + "'"
            elif SegFeild[19] != "" and SegFeild[24] != "" and SegFeild[7] != "":
                commandText = "SELECT Posting_Id from Tenders where short_desc = '" + str(SegFeild[19]) + "' and doc_last = '" + SegFeild[24] + "' and Country = '" + SegFeild[7] + "'"
            else:
                commandText = "SELECT Posting_Id from Tenders where short_desc = '" + str(SegFeild[19]) + "' and Country = '" + str(SegFeild[7]) + "'"
            mycursorLocal.execute(commandText)
            results = mycursorLocal.fetchall()
            a1 = 1
            print("Code Reached On check_Duplication")
            return results
        except Exception as e:
            mydb_L2L = Local_connection()
            mycursorL2L = mydb_L2L.cursor()
            Function_name: str = sys._getframe().f_code.co_name
            Error: str = str(e)
            sql1 = "INSERT INTO ErrorLog(Error_Message,Function_Name,Exe_Name) VALUES('" + str(Error).replace("'" , "") + "','" + str(Function_name).replace("'" , "") + "','" + str(SegFeild[31]) + "')"
            mycursorL2L.execute(sql1)
            mydb_L2L.commit()
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : ", sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname ,"\n" , exc_tb.tb_lineno)
            time.sleep(10)
            a1 = 0


def insert_in_Local(get_htmlSource , SegFeild):
    mydb_Local = Local_connection()
    mycursorLocal = mydb_Local.cursor()
    results = check_Duplication(get_htmlSource , SegFeild)
    if len(results) > 0:
        print('Duplicate Tender')
        global_var.duplicate += 1
        return 1
    else:
        MyLoop = 0
        while MyLoop == 0:
            sql = "INSERT INTO Tenders(EMail,add1,Country,Maj_Org,tender_notice_no,notice_type,Tenders_details,short_desc,est_cost,currency,doc_cost,doc_last,earnest_money,Financier,tender_doc_file,source)VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
            val = (str(SegFeild[1]) , str(SegFeild[2]) , str(SegFeild[7]) , str(SegFeild[12]) , str(SegFeild[13]) , str(SegFeild[14]),
                   str(SegFeild[18]) , str(SegFeild[19]) , str(SegFeild[20]) , str(SegFeild[21]) , str(SegFeild[22]), str(SegFeild[24]),str(SegFeild[26]) ,str(SegFeild[27]),
                   str(SegFeild[28]) , str(SegFeild[31]))
            try:
                mycursorLocal.execute(sql , val)
                mydb_Local.commit()
                global_var.inserted += 1
                print("Code Reached On insert_in_Local")
                MyLoop = 1
            except Exception as e:
                mydb_L2L = L2L_connection()
                mycursorL2L = mydb_L2L.cursor()
                Function_name :str = sys._getframe().f_code.co_name
                Error : str = str(e)
                sql1 = "INSERT INTO ErrorLog(Error_Message,Function_Name,Exe_Name) VALUES('" + str(Error).replace("'","") + "','" + str(Function_name).replace("'","")+ "','"+str(SegFeild[31])+"')"
                mycursorL2L.execute(sql1)
                mydb_L2L.commit()
                exc_type , exc_obj , exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" ,
                      fname , "\n" , exc_tb.tb_lineno)
                MyLoop = 0
                time.sleep(10)
    create_filename(get_htmlSource , SegFeild)


def create_filename(get_htmlSource , SegFeild):
    a = 0
    while a == 0:
        try:
            basename = "PY544"
            Current_dateTime = datetime.now().strftime("%Y%m%d%H%M%S%f")
            Fileid = "".join([basename , Current_dateTime])
            File_path = "Z:\\" + Fileid + ".html"
            file1 = open(File_path , "w" , encoding='utf-8')
            Final_Doc = "<HTML><style style=\"text/css\"> td{ padding:5px;}</style><BODY>" + get_htmlSource + "</BODY></HTML>"
            Final_Doc = "<head><meta content=\"text/html; charset=utf-8\" http-equiv=\"Content-Type\" /><title>Tender Document</title>" +\
                        "<link rel=\"shortcut icon\" type=\"image/png\" href=\"https://www.tendersontime.com/favicon.ico\"/></head>"+\
                        "<Blockquote style='border:1px solid; padding:10px;'>" + Final_Doc + "</Blockquote>"
            Final_HTML_Document = Final_Doc + "</body>"

            file1.write(str(Final_HTML_Document))
            file1.close()
            print("Code Reached On create_filename")
            insert_L2L(SegFeild, Fileid)
            a = 1
        except Exception as e:
            mydb_L2L = L2L_connection()
            mycursorL2L = mydb_L2L.cursor()
            Function_name: str = sys._getframe().f_code.co_name
            Error: str = str(e)
            sql1 = "INSERT INTO ErrorLog(Error_Message,Function_Name,Exe_Name) VALUES('" + str(Error).replace("'" ,
                                                                                                              "") + "','" + str(
                Function_name).replace("'" , "") + "','" + str(SegFeild[31]) + "')"
            mycursorL2L.execute(sql1)
            mydb_L2L.commit()
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" ,
                  fname , "\n" , exc_tb.tb_lineno)
            a = 0
            time.sleep(10)


def insert_L2L(SegFeild , Fileid):
    ncb_icb = "icb"
    dms_entrynotice_tblstatus = "1"
    added_on = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    search_id = "1"
    cpv_userid = ""
    dms_entrynotice_tblquality_status = '1'
    quality_id = '1'
    quality_addeddate = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    Col1 = 'https://procure.ohio.gov/'
    if SegFeild[7] == "IN":
        Col2 = str(SegFeild[26]) + " * " + str(SegFeild[20])  # For India Only Other Wise Blank
    else:
        Col2 = ''
    Col3 = ''
    Col4 = ''
    Col5 = ''
    file_name = "D:\\Tide\\DocData\\" + Fileid + ".html"
    dms_downloadfiles_tbluser_id = 'DWN1011566'
    # Europe-DWN2554488,India-DWN00541021,Asia-DWN5046627,Africa-DWN302520,North America-DWN1011566,
    # South America-DWN1456891,Semi-Auto-DWN30531073,MFA-DWN0654200
    dms_downloadfiles_tblstatus = '1'
    dms_downloadfiles_tblsave_status = '1'
    selector_id = ''
    select_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    dms_downloadfiles_tbldatatype = "A"
    dms_entrynotice_tblnotice_type = '2'
    dms_entrynotice_tbl_cqc_status = '1'
    file_id = Fileid
    mydb_L2L = L2L_connection()
    mycursorL2L = mydb_L2L.cursor()
    if SegFeild[12] != "" and SegFeild[19] != "" and SegFeild[24] != "" and SegFeild[7] != "" and SegFeild[2] != "":
        dms_entrynotice_tblcompulsary_qc = "2"
    else:
        dms_entrynotice_tblcompulsary_qc = "1"
        global_var.QC_Tenders += 1
        sql = "INSERT INTO QCTenders(Source,tender_notice_no,short_desc,doc_last,Maj_Org,Address,doc_path,Country)VALUES(%s,%s,%s,%s,%s,%s,%s,%s) "
        val = (str(SegFeild[31]) , str(SegFeild[13]) , str(SegFeild[19]) , str(SegFeild[24]) , str(SegFeild[12]) ,str(SegFeild[2]) , "http://tottestupload3.s3.amazonaws.com/" + file_id + ".html" , str(SegFeild[7]))
        a4 = 0
        while a4 == 0:
            try:
                mycursorL2L.execute(sql , val)
                mydb_L2L.commit()
                a4 = 1
                print("Code Reached On QCTenders")
            except Exception as e:
                Function_name: str = sys._getframe().f_code.co_name
                Error: str = e
                sql1 = "INSERT INTO ErrorLog(Error_Message,Function_Name,Exe_Name) VALUES('" + str(Error).replace("'" , "") + "','" + str(Function_name).replace("'" , "") + "','" + str(SegFeild[31]) + "')"
                mycursorL2L.execute(sql1)
                mydb_L2L.commit()
                exc_type , exc_obj , exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" ,fname ,"\n" , exc_tb.tb_lineno)
                a4 = 0
                time.sleep(10)

    sql = "INSERT INTO Final_Tenders(notice_no,file_id,purchaser_name,deadline,country,description,purchaser_address,purchaser_email,purchaser_url,purchaser_emd,purchaser_value,financier,deadline_two,tender_details,ncbicb,status,added_on,search_id,cpv_value,cpv_userid,quality_status,quality_id,quality_addeddate,source,tender_doc_file,Col1,Col2,Col3,Col4,Col5,file_name,user_id,status_download_id,save_status,selector_id,select_date,datatype,compulsary_qc,notice_type,cqc_status,DocCost,DocLastDate)VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
    val = (str(SegFeild[13]) , file_id , str(SegFeild[12]) , str(SegFeild[24]) , str(SegFeild[7]) , str(SegFeild[19]) ,str(SegFeild[2]) ,str(SegFeild[1]) , str(SegFeild[8]) , str(SegFeild[26]) , str(SegFeild[20]) , str(SegFeild[27]) ,str(SegFeild[24]) , str(SegFeild[18]) , ncb_icb , dms_entrynotice_tblstatus , str(added_on) , search_id ,str(SegFeild[36]) ,cpv_userid , dms_entrynotice_tblquality_status , quality_id , str(quality_addeddate) , str(SegFeild[31]) ,str(SegFeild[28]) ,Col1 , Col2 , Col3 , Col4 , Col5 ,file_name , dms_downloadfiles_tbluser_id , dms_downloadfiles_tblstatus , dms_downloadfiles_tblsave_status ,selector_id , str(select_date) , dms_downloadfiles_tbldatatype ,dms_entrynotice_tblcompulsary_qc , dms_entrynotice_tblnotice_type , dms_entrynotice_tbl_cqc_status ,str(SegFeild[22]) , str(SegFeild[41]))
    a5 = 0
    while a5 == 0:
        try:
            mycursorL2L.execute(sql , val)
            mydb_L2L.commit()
            print("Code Reached On insert_L2L")
            a5 = 1
        except Exception as e:
            Function_name: str = sys._getframe().f_code.co_name
            Error: str = str(e)
            sql1 = "INSERT INTO ErrorLog(Error_Message,Function_Name,Exe_Name) VALUES('" + str(Error).replace("'" , "") + "','" + str(Function_name).replace("'" , "") + "','" + str(SegFeild[31]) + "')"
            mycursorL2L.execute(sql1)
            mydb_L2L.commit()
            exc_type , exc_obj , exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" ,fname ,"\n" , exc_tb.tb_lineno)
            a5 = 0
            time.sleep(10)
