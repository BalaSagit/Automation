import os.path
import datetime

now = str(datetime.datetime.now())[0:10]
string1="Hi Team,\n\tPlease find below, the status report of the failed folders.\n\n"


def find_count(path,name):
    global string1
    num_files = len([f for f in os.listdir(path)if os.path.isfile(os.path.join(path, f))])
    string1+="\t\t"+name+" :   "+str(num_files)+" Failed Files\n"


def sendEmail(string1):
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    # email = ";".join(emailList)
    mail.To = 'balaji.s@mrcooper.com'
    # mail.CC = 'Bharathwaj.Vasudevan@mrcooper.com'
    mail.Subject = 'Failed Folder Monitoring - '+ now
    mail.Body = '{}'.format(string1)
    try:
        mail.Send()
        print("Successfully sent email")
    except Exception:
        print("Error: unable to send email")

print("Calculating the count of failed folders...")
find_count("\\\\Chec.local\\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\WALZ\\Failed","WALZ               ")
find_count("\\\\Chec.local\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\Venture\Failed","Venture NSM   ")
find_count("\\\\Chec.local\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\VentureUSAA\\Failed","Venture USAA  ")
find_count("\\\\Chec.local\\dfs\\vrapscont_archive\\FTP\\LoanDocuments\\BlitzDocs\\Failed","Blitz Docs         ")
find_count("\\\\Vrisi01\\shared$\\Rshare\CHECIT\\FTPLive\\LPSImport\\lpsimport_failed_NSM","LPS NSM           ")
find_count("\\\\Vrisi01\\shared$\\Rshare\CHECIT\\FTPLive\\LPSImport\\lpsimport_failed_USAA","LPS USAA          ")
find_count("\\\\vrdmzsftp02\\lps$\\FileNET\\Failed","Lending Space  ")

string1+="\nThanks and Regards\nImaging Support"
sendEmail(string1)
