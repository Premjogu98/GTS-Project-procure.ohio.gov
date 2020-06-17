duplicate = 0
inserted = 0
expired = 0
skipped = 0
deadline_Not_given = 0
Fromdate = ''
Todate = ''
On_Error = 0
TenderDetails2 =""
Total = 0
QC_Tenders = 0


def Process_End():
    # print("Total Links "+str(Total)+"")
    print("Total: ", Total-1)
    print('Duplicate: ' , duplicate)
    print('Expired: ' , expired)
    print('Inserted: ' , inserted)
    print('Deadline Not given: ', deadline_Not_given)
    print('QC Tenders: ', QC_Tenders)

