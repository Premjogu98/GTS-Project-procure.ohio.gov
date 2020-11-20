import ctypes
import html
import os
import string
import sys
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import global_var
from insert_on_databsae import insert_in_Local
import wx
app = wx.App()

def choromedriver():

    From_date = global_var.Fromdate.partition("Date (FROM)")[2].partition("00:")[0].strip()
    # 2019-07-24 # Tender Actual Date Format
    Year = From_date[0:4]
    Month = From_date[5:7]
    Day = From_date[8:10]
    From_date = Month + '/' + Day + '/' + Year

    To_date = global_var.Todate.partition("Date (TO)")[2].partition("00:")[0].strip()
    # 2019-07-24 # Tender Actual Date Format
    Year = To_date[0:4]
    Month = To_date[5:7]
    Day = To_date[8:10]
    To_date = Month + '/' + Day + '/' + Year

    # OG_URL = 'https://procure.ohio.gov/proc/searchProcOppsResults.asp?t1=0&IN=&DBN=&SK=&MT=All&KST=All%20Words&OT=0&OSTAT=Active&SDT=POST&SD=10/25/2019&ED=10/29/2019&A=All&AT=All&OTT=All&SDTT=Search%20by%20Posted%20Date&MTT=&OSTATT=Active'

    Custom_URL = "https://procure.ohio.gov/proc/searchProcOppsResults.asp?t1=0&IN=&DBN=&SK=&MT=All&KST=All%20Words&OT=0&OSTAT=Active&SDT=POST&SD="+str(From_date)+"&ED="+str(To_date)+"&A=All&AT=All&OTT=All&SDTT=Search%20by%20Posted%20Date&MTT=&OSTATT=Active"
    chrome_options = Options()
    chrome_options.add_extension('C:\\BrowsecVPN.crx')
    browser = webdriver.Chrome(executable_path=str(f"C:\\chromedriver.exe"),chrome_options=chrome_options)
    browser.maximize_window()
    # browser.get("""https://chrome.google.com/webstore/detail/browsec-vpn-free-and-unli/omghfjlpggmjjaagoclmmobgdodcjboh?hl=en" ping="/url?sa=t&amp;source=web&amp;rct=j&amp;url=https://chrome.google.com/webstore/detail/browsec-vpn-free-and-unli/omghfjlpggmjjaagoclmmobgdodcjboh%3Fhl%3Den&amp;ved=2ahUKEwivq8rjlcHmAhVtxzgGHZ-JBMgQFjAAegQIAhAB""")
    wx.MessageBox(' -_-  Add Extension and Select Proxy Between 10 SEC -_- ', 'Info', wx.OK | wx.ICON_WARNING)
    time.sleep(15)  # WAIT UNTIL CHANGE THE MANUAL VPN SETtING
    browser.get(Custom_URL)
    clicking_process(browser)


def clicking_process(browser):

    try:
        for Record_not_found in browser.find_elements_by_xpath('//*[@class="red"]'):
            Record_not_found = Record_not_found.get_attribute('innerText').strip()
            if Record_not_found == "No records found!":
                ctypes.windll.user32.MessageBoxW(0, " ♥ Record Not Found ! ♥ " "procure.ohio.gov", 1)
                browser.close()
                sys.exit()
            break
    except:pass
    List_Of_Tender_Href = []
    for Tender_href in browser.find_elements_by_xpath('//*[@id="full"]/table/tbody//td[7]/a'):
        Tender_href = Tender_href.get_attribute('href')
        List_Of_Tender_Href.append(Tender_href)
    if len(List_Of_Tender_Href) == 0:
        ctypes.windll.user32.MessageBoxW(0, " ♥ Record Not Found ! ♥ " "procure.ohio.gov", 1)
        browser.close()
        sys.exit()
    a = True
    while a == True:
        try:
            for href in List_Of_Tender_Href:
                get_htmlsource = ""
                browser.get(href)
                global_var.Total += 1
                SegFeild = []
                for data in range(45):
                    SegFeild.append('')

                for get_htmlsource in browser.find_elements_by_xpath('//*[@id="left"]'):
                    get_htmlsource = get_htmlsource.get_attribute('outerHTML').replace('src="/img/', 'src="https://procure.ohio.gov/img/').replace('href="glossary', 'href="https://procure.ohio.gov/proc/glossary').replace('SITE VISIT ', '').replace('1</span>', '').replace('<h2>Alerts</h2>', '').replace('back</a>', '</a>').replace('alt="Get Acrobat Reader"', '') \
                        .replace('href="submitDocumentInquiries', 'href="https://procure.ohio.gov/proc/submitDocumentInquiries').replace('href="viewDocQuestions', 'href="https://procure.ohio.gov/proc/viewDocQuestions')
                    break

                Title = get_htmlsource.partition("Opportunity Title</h2>")[2].partition("</p>")[0].strip()
                Title = Title.partition("<b>")[2].partition("</b>")[0]
                Title = string.capwords(str(Title.strip()))
                Title = html.unescape(str(Title))
                SegFeild[19] = Title.strip()

                SegFeild[2] = 'Ohio, United States of America'

                Tender_Detail = get_htmlsource.partition("Opportunity Description</h2>")[2].partition("</p>")[0].strip()
                Tender_Detail = Tender_Detail.partition("<p>")[2]
                Tender_Detail = html.unescape(str(Tender_Detail)).rstrip(' ')
                Tender_Detail = string.capwords(str(Tender_Detail.strip())).replace('<br>', ' ')

                Opportunity_Type = get_htmlsource.partition("Opportunity Type:</a></div>")[2].partition("</div>")[0].strip()
                Opportunity_Type = Opportunity_Type.partition("\">")[2]
                Opportunity_Type = html.unescape(str(Opportunity_Type)).rstrip(' ')

                Opportunity_Status = get_htmlsource.partition("Opportunity Status:</a></div>")[2].partition("</div>")[0].strip()
                Opportunity_Status = Opportunity_Status.partition("\">")[2]
                Opportunity_Status = html.unescape(str(Opportunity_Status)).rstrip(' ')

                Bid_Number_Detail = get_htmlsource.partition("Document/Bid #:</a></div>")[2].partition("</div>")[0].strip()
                Bid_Number_Detail = Bid_Number_Detail.partition("\">")[2]
                Bid_Number_Detail = html.unescape(str(Bid_Number_Detail)).rstrip(' ')

                Index = get_htmlsource.partition("Index #:</a></div>")[2].partition("</div>")[0].strip()
                Index = Index.partition("\">")[2]
                Index = html.unescape(str(Index)).rstrip(' ')

                Requesting_Agency = get_htmlsource.partition("Requesting Agency:</a></div>")[2].partition("</div>")[0].strip()
                Requesting_Agency = Requesting_Agency.partition("\">")[2]
                Requesting_Agency = html.unescape(str(Requesting_Agency)).rstrip(' ')

                Issued_By = get_htmlsource.partition("Issued By:</a></div>")[2].partition("</div>")[0].strip()
                Issued_By = Issued_By.partition("\">")[2].rstrip(' ')
                Issued_By = html.unescape(str(Issued_By)).rstrip(' ')

                Posted_Date = get_htmlsource.partition("Posted Date:</a></div>")[2].partition("</div>")[0].strip()
                Posted_Date = Posted_Date.partition("\">")[2]

                Inquiry_Period = get_htmlsource.partition("Inquiry Period:</a></div>")[2].partition("</div>")[0].strip()
                Inquiry_Period = Inquiry_Period.partition("\">")[2]
                Inquiry_Period = html.unescape(str(Inquiry_Period)).rstrip(' ')

                Opening_Date = get_htmlsource.partition("Opening Date:</a></div>")[2].partition("</div>")[0].strip()
                Opening_Date = Opening_Date.partition("\">")[2].strip()

                MBE_Set_Aside = get_htmlsource.partition("MBE Set Aside:</a></div>")[2].partition("</div>")[0].strip()
                MBE_Set_Aside = MBE_Set_Aside.partition("\">")[2].strip()

                Open_Market = get_htmlsource.partition("Open Market:</a></div>")[2].partition("</div>")[0].strip()
                Open_Market = Open_Market.partition("\">")[2].strip()

                Tender_Details = "Tender Details:  "+Tender_Detail.strip()+"<br>""\n"+"Opportunity Type:  "+Opportunity_Type.strip()+"<br>""\n""Opportunity Status:  "+Opportunity_Status.strip()+"<br>""\n""Document/Bid:  "+Bid_Number_Detail.strip() \
                               + "<br>""\n""Index:  "+Index.strip() + "<br>""\n""Requesting Agency:  "+Requesting_Agency.strip() + "<br>""\n""Issued By:  "+Issued_By.strip()+ "<br>\n""Posted Date:  "+Posted_Date.strip() \
                               + "<br>""\n""Inquiry_Period:  " + Inquiry_Period.strip() + "<br>""\n""Opening Date:  " + Opening_Date.strip() + "<br>""\n""MBE Set Aside:  " + MBE_Set_Aside.strip() + "<br>""\n""Open Market:  " + Open_Market.strip()
                Tender_Details = string.capwords(str(Tender_Details.strip()))
                SegFeild[18] = Tender_Details

                Bid_Number = get_htmlsource.partition("Document/Bid #:</a></div>")[2].partition("</div>")[0].strip()
                Bid_Number = Bid_Number.partition("\">")[2].rstrip(' ')
                SegFeild[13] = Bid_Number.strip()

                Purchaser = get_htmlsource.partition("Issued By:</a></div>")[2].partition("</div>")[0].strip()
                Purchaser = Purchaser.partition("\">")[2]
                Purchaser = html.unescape(str(Purchaser)).rstrip(' ')
                SegFeild[12] = Purchaser.strip().upper()

                Submission_Date = get_htmlsource.partition("Opening Date:</a></div>")[2].partition("</div>")[0].strip()
                Submission_Date = Submission_Date.partition("\">")[2].strip()
                datetime_object = datetime.strptime(Submission_Date, "%m/%d/%Y")
                mydate = datetime_object.strftime("%Y-%m-%d").strip()
                SegFeild[24] = mydate

                SegFeild[7] = "US"

                # notice type
                SegFeild[14] = "2"

                SegFeild[22] = "0"

                SegFeild[26] = "0.0"

                SegFeild[27] = "0"  # Financier

                SegFeild[28] = str(href)

                # Source Name
                SegFeild[31] = 'procure.ohio.gov'
                SegFeild[20] = ''
                SegFeild[21] = ''
                SegFeild[42] = SegFeild[7]
                SegFeild[43] = ''

                for SegIndex in range(len(SegFeild)):
                    print(SegIndex, end=' ')
                    print(SegFeild[SegIndex])
                    SegFeild[SegIndex] = html.unescape(str(SegFeild[SegIndex]))
                    SegFeild[SegIndex] = str(SegFeild[SegIndex]).replace("'", "''")
                if len(SegFeild[19]) >= 200:
                    SegFeild[19] = str(SegFeild[19])[:200]+'...'

                if len(SegFeild[18]) >= 1500:
                    SegFeild[18] = str(SegFeild[18])[:1500]+'...'

                check_date(get_htmlsource, SegFeild)
                print(" Total: " + str(global_var.Total) + " Duplicate: " + str(global_var.duplicate) + " Expired: " + str(global_var.expired) + " Inserted: " + str(global_var.inserted) + " Skipped: " + str(global_var.skipped) + " Deadline Not given: " + str(global_var.deadline_Not_given) + " QC Tenders: " + str(global_var.QC_Tenders),'\n')
                a = False
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print("Error ON : ", sys._getframe().f_code.co_name + "--> " + str(e), "\n", exc_type, "\n", fname, "\n",
                  exc_tb.tb_lineno)
            a = True
    ctypes.windll.user32.MessageBoxW(0, "Total: " + str(global_var.Total) + "\n""Duplicate: " + str(
            global_var.duplicate) + "\n""Expired: " + str(global_var.expired) + "\n""Inserted: " + str(
            global_var.inserted) + "\n""Skipped: " + str(
            global_var.skipped) + "\n""Deadline Not given: " + str(
            global_var.deadline_Not_given) + "\n""QC Tenders: " + str(global_var.QC_Tenders) + "",
                                         "procure.ohio.gov", 1)
    global_var.Process_End()
    browser.close()
    sys.exit()


def check_date(get_htmlSource, SegFeild):
    deadline = str(SegFeild[24])
    curdate = datetime.now()
    curdate_str = curdate.strftime("%Y-%m-%d")
    try:
        if deadline != '':
            datetime_object_deadline = datetime.strptime(deadline, '%Y-%m-%d')
            datetime_object_curdate = datetime.strptime(curdate_str, '%Y-%m-%d')
            timedelta_obj = datetime_object_deadline - datetime_object_curdate
            day = timedelta_obj.days
            if day > 0:
                insert_in_Local(get_htmlSource, SegFeild)
            else:
                print("Expired Tender")
                global_var.expired += 1
        else:
            print("Deadline Not Given")
            global_var.deadline_Not_given += 1
    except Exception as e:
        exc_type , exc_obj , exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print("Error ON : " , sys._getframe().f_code.co_name + "--> " + str(e) , "\n" , exc_type , "\n" , fname , "\n" ,exc_tb.tb_lineno)

choromedriver()