
import datetime
import xlrd
import sys
import csv
import os
import json
import time
import requests
from sets import Set
from xlsxwriter.workbook import Workbook
from requests.auth import HTTPBasicAuth
from TableIndex import UserIndex,PerPersonIndex,PerPersonalIndex,EmpEmploymentIndex,EmpJobIndex,PrefixMap,MaritalStatusMap,EmploymentTypeMap
import httplib2
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
from django.conf import settings
import cloudstorage as gcs

from google.appengine.api import app_identity
from google.appengine.api import mail
sender_email_id="raj.jha@flipkart.com"
############################################################################################################################################
ProductionTemplateFilePath = "ProductionTemplate.tsv"
api_url = "https://api10preview.sapsf.com/odata/v2/"
userName = "SFADMIN"
companyID = "C0017935023D"
password = "SFADMIN123"
datemode = ''
DuplicateCheckIndex = ["Date of Birth*","Date of Joining*","Fathers Name*","FirstName*","LastName*"]
############################################################################################################################################

SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'sfValidator'


#ProductionTemplateFilePath = settings.ProductionTemplateFilePath
#api_url = settings.api_url
#userName = settings.userName
#companyID =settings.companyID
#password = settings.password

#DuplicateCheckIndex = settings.DuplicateCheckIndex

#SCOPES = settings.SCOPES
#CLIENT_SECRET_FILE = settings.CLIENT_SECRET_FILE
#APPLICATION_NAME = settings.APPLICATION_NAME



datemode = ''
class Node:
	def __init__(self):
		self.child = {}
		self.code = "Invalid"


def XlsxToTsv(FilePath):
	reload(sys)
	sys.setdefaultencoding('utf-8')
	wb = xlrd.open_workbook(file_contents = gcs.open(FilePath).read())
	sh = wb.sheet_by_index(0)
	global datemode
	datemode = wb.datemode
	FilePathTsv= FilePath.split(".xlsx")[0]
	csvFile = gcs.open(FilePathTsv+'.tsv', 'w')
	wr = csv.writer(csvFile,delimiter='\t')

	for rownum in xrange(sh.nrows):
        	wr.writerow(sh.row_values(rownum))
	csvFile.close()
	gcs.delete(FilePath)
	return FilePathTsv+".tsv"

def MasterFileTuples(FilePath):
	FileName = XlsxToTsv(FilePath)
	fileHandle = open(FileName,"ru")
	LevelTitle = fileHandle.readline().split("\t")
	TupleArray = Set()
	Lines = fileHandle.read()
	for line in Lines.splitlines():
		level = line.split("\t")
		Temp = {}
		for i in xrange(len(LevelTitle)):
			Temp[LevelTitle[i]]= level[i]
		Tup = []
		for i in DuplicateCheckIndex:
			Tup.append(Temp[i])
		TupleArray.add(tuple(Tup))
	return TupleArray
	
def EmployeeData(FilePath):
	TsvFileName = XlsxToTsv(FilePath)
	fileHandle = gcs.open(TsvFileName,"r")
	ErrorReport = []
	LevelTitle =  fileHandle.readline().split("\t")
	LevelTitle = [level for level in LevelTitle]
	Lines = fileHandle.read()
	TotalRecord = 0
	Output = {}
	for Level in LevelTitle:
		Output[Level] = []
	DuplicateChecker = Set()
	idx=0
#	MasterFile = MasterFileTuples(MasterFileName)
	for line in Lines.splitlines():
		idx+=1
		level = line .split("\t")
		level = [x.split(".0")[0] for x in level]
		Temp = {}
		for i in xrange(0,len(LevelTitle)):
			if level[i] == "":
				Temp[LevelTitle[i]] = None
				if LevelTitle[i][-1] == "*":
					ErrorReport.append("Required Field "+LevelTitle[i]+" Empty at " +str(idx+1)+" Row Number")
			else:
				if LevelTitle[i]=='Prefix*':
					level[i] = PrefixMap[ level[i] ]
				if LevelTitle[i]=='Marital Status*':
                                        level[i] = MaritalStatusMap[ level[i] ]
				if LevelTitle[i]=='Employment Type*':
                                        level[i] = EmploymentTypeMap[ level[i] ]

				
				Temp[LevelTitle[i]] = level[i]
					
		DuplicateCheckerTuple = []
                for i in DuplicateCheckIndex:
                        DuplicateCheckerTuple.append(Temp[i])
                DuplicateCheckerTuple = tuple(DuplicateCheckerTuple)
                if DuplicateCheckerTuple in DuplicateChecker:
                        continue
#		if DuplicateCheckerTuple in MasterFile:
#			ErrorReport.append("Employee already present in the Master File.Error at " +str(idx+1)+ "Row Number ")			
		for i in xrange(0,len(LevelTitle)):
                        Output[LevelTitle[i]].append(Temp[LevelTitle[i]])
		TotalRecord+=1
                DuplicateChecker.add(DuplicateCheckerTuple)
	gcs.delete(TsvFileName)
	return (LevelTitle,Output,TotalRecord,ErrorReport)
def GFDTstructure(filePath):
	fileHandle = open(filePath,"r")
        LevelTitle = fileHandle.readline().split("\t")
	LevelTitle = [LevelTitle[i] for i in xrange(0,len(LevelTitle),2)]

	Lines = fileHandle.read()

	root = Node()
	for row in Lines.splitlines():
		row = row.split("\t")
		curNode = root
		for i in xrange(0,len(row),2):
			if row[i].lower() in curNode.child.keys():
				curNode = curNode.child[row[i].lower()]
			else:
				curNode.child[row[i].lower()] = Node()
				curNode.child[row[i].lower()].code = row[i+1]
				curNode = curNode.child[row[i].lower()]
	return LevelTitle,root

def GFDTStructureVerifier(LevelTitle,root,Fieldid,Employees):
	ErrorReport = []
	for idx in xrange(len(Employees)):
		curNode = root
		for level in LevelTitle:
			if Employees[level+"*"][idx].lower() not in curNode.child.keys():
				ErrorReport.append(level.title() + " Error at " + str(idx+2) + " Row Number")
				break
			else:
				curNode = curNode.child[  Employees[level+"*"][idx].lower()  ]
				
	return ErrorReport
	
def XlsxErrorReport(Errors,ReportName):
	tsv_file = "/sfbulkupload.appspot.com/"+ReportName+"Error" + ".tsv"
	xlsx_file = ReportName+"Error" + ".xlsx"


	tsvHandle = gcs.open(tsv_file,"w")
	tsvHandle.write("Error\n")
	for error in Errors:
		tsvHandle.write(error+"\n")
	tsvHandle.close()
	return tsv_file

	
	workbook = Workbook(xlsx_file)
	worksheet = workbook.add_worksheet()
	
	tsv_reader = csv.reader(open(tsv_file, 'rb'), delimiter='\t')

	for row, data in enumerate(tsv_reader):
		worksheet.write_row(row, 0, data)
	workbook.close()
	os.remove(tsv_file)
	return xlsx_file


def send_mail(Subject,Message,FileName,To_email):
	try:
		To_email = [To_email]
		email = EmailMessage(Subject,Message, To_email)
		email.attach_file(FileName)
		email.send(fail_silently=False)
	except:
		print "Mail was Not Send"

	
def getNewUserID():
	r= requests.post(api_url+"generateNextPersonID?$format=json",auth=HTTPBasicAuth(userName+'@'+companyID,password))
	response = json.loads(r.text)
	return response["d"]["GenerateNextPersonIDResponse"]["personID"].decode("utf-8").encode("ascii","ignore")

	
def sfInsert(FieldId,Employees,TotalEmployee):

	for idx in xrange(TotalEmployee):
		employee = {}
		for x in FieldId:
			employee[x] = Employees[x][idx]
		UserPayload={}
		userID = getNewUserID()
		UserEntityInsert(userID,employee)
		PerPersonInsert(userID,employee)
#		EmpEmploymentInsert(userID,employee)
		EmpJobInsert(userID,employee)
		PerPersonalInsert(userID,employee)


def getSFDate(date):
	year,month,day,hour,minute,second =  xlrd.xldate_as_tuple(int(date),datemode)
	py_date = datetime.datetime(year, month, day, hour, minute, second)
	date = int((py_date-datetime.datetime(1970,1,1)).total_seconds())*1000
	return "/Date("+str(date)+")/"


def getDate(date):
	year,month,day,hour,minute,second =  xlrd.xldate_as_tuple(int(date),datemode)
        py_date = datetime.datetime(year, month, day, hour, minute, second)
	return py_date

	
def UserEntityInsert(userID,employee):
    UserPayload = {}	
    for key,value in UserIndex.iteritems():
	if "date" not in key.lower():
	     UserPayload[value] = employee[key]
	else:
	     UserPayload[value] = getSFDate(employee[key])
    UserPayload["status"] = "active"
    UserPayload["userId"] = userID
    UserPayload["username"] = userID
    UserPayload["hr"] = "User('"+UserPayload["hr"]+"')"
    UserPayload["manager"] = "User('"+UserPayload["manager"]+"')"
    Userauth = HTTPBasicAuth(userName+"@"+companyID,password)
    Meta = {}
    Meta["uri"] ="User(userId='"+userID+"')"
    UserPayload["__metadata"]=Meta
    

    UserHeader = {}
    UserHeader["content-type"] ="application/json; charset=utf-8"
    UserHeader["accept"] = "application/json"
    r = requests.post(api_url+"upsert",auth = Userauth,headers = UserHeader,  data = json.dumps(UserPayload))
 #   print r.text


def PerPersonInsert(userID,employee):
    PerPersonPayload = {}	
    for key,value in PerPersonIndex.iteritems():
	if "date" not in key.lower():
	      PerPersonPayload[value] = employee[key]
	else:
	      PerPersonPayload[value] = getSFDate(employee[key])
    PerPersonPayload["personIdExternal"] = userID
    PerPersonPayload["userId"] = userID
    Userauth = HTTPBasicAuth(userName+"@"+companyID,password)
    Meta = {}
    Meta["uri"] ="PerPerson('"+userID+"')"
    PerPersonPayload["__metadata"]=Meta
    PerPersonHeader = {}
    PerPersonHeader["content-type"] ="application/json; charset=utf-8"
    PerPersonHeader["accept"] = "application/json"
    r = requests.post(api_url+"upsert",auth = Userauth,headers = PerPersonHeader,  data = json.dumps(PerPersonPayload))
#    print r.text
#    sys.exit()        

def EmpEmploymentInsert(userID,employee):
    EmpEmploymentPayload = {}	
    for key,value in EmpEmploymentIndex.iteritems():
	if "date" not in key.lower():
	      EmpEmploymentPayload[value] = employee[key]
	else:
	      EmpEmploymentPayload[value] = getSFDate(employee[key])
#    EmpEmploymentPayload["status"] = "active"
    EmpEmploymentPayload["userId"] = userID
#    EmpEmploymentPayload["username"] = userID
    Userauth = HTTPBasicAuth(userName+"@"+companyID,password)
    Meta = {}
    Meta["uri"] ="EmpEmployment(personIdExternal='"+userID+"',userId='"+userID+"')"
    EmpEmploymentPayload["__metadata"]=Meta
    EmpEmploymentHeader = {}
    EmpEmploymentHeader["content-type"] ="application/json; charset=utf-8"
    EmpEmploymentHeader["accept"] = "application/json"
    r = requests.post(api_url+"upsert",auth = Userauth,headers = EmpEmploymentHeader,  data = json.dumps(EmpEmploymentPayload))
#    print r.text        			
 #   sys.exit()		
def EmpJobInsert(userID,employee):
    EmpJobPayload = {}	
    for key,value in EmpJobIndex.iteritems():
	if "date" not in key.lower():
	      EmpJobPayload[value] = employee[key]
	else:
	      EmpJobPayload[value] = getSFDate(employee[key])
    EmpJobPayload["eventReason"] = "direct"
    EmpJobPayload["jobTitle"] = "contract"
    EmpJobPayload["payGrade"] = "contract"
#    EmpJobPayload["countryOfCompany"]="1776"
    EmpJobPayload["userId"] = userID
    Userauth = HTTPBasicAuth(userName+"@"+companyID,password)
    Meta = {}
    Meta["uri"] ="EmpJob"
    EmpJobPayload["__metadata"]=Meta
    EmpJobHeader = {}
    EmpJobHeader["content-type"] ="application/json; charset=utf-8"
    EmpJobHeader["accept"] = "application/json"
    r = requests.post(api_url+"upsert",auth = Userauth,headers = EmpJobHeader,  data = json.dumps(EmpJobPayload))
#    print r.text
#    sys.exit()


def PerPersonalInsert(userID,employee):
    PerPersonalPayload = {}	
    for key,value in PerPersonalIndex.iteritems():
	if "date" not in key.lower():
	      PerPersonalPayload[value] = employee[key]
	else:
	      PerPersonalPayload[value] = getSFDate(employee[key])
    PerPersonalPayload["personIdExternal"] = userID
    Userauth = HTTPBasicAuth(userName+"@"+companyID,password)
    Meta = {}
    Meta["uri"] ="PerPersonal(personIdExternal='"+userID+"',startDate=datetime'"+str(getDate(employee["Date of Joining*"]).isoformat())+"')"
    PerPersonalPayload["__metadata"]=Meta
    PerPersonalHeader = {}
    PerPersonalHeader["content-type"] ="application/json; charset=utf-8"
    PerPersonalHeader["accept"] = "application/json"
    r = requests.post(api_url+"upsert",auth = Userauth,headers = PerPersonalHeader,  data = json.dumps(PerPersonalPayload))
 #   print r.text
 #   sys.exit()        			
		
def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials	
def CreateMessage(Subject,Message,FileName,To_email):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEText(Message)
  message['to'] = To_email
  message['from'] = "me"
  message['subject'] = Subject
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def CreateMessageWithAttachment(Subject,Message,FileName,emailID):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file_dir: The directory containing the file to be attached.
    filename: The name of the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = emailID
  message['from'] = "me"
  message['subject'] = Subject

  msg = MIMEText(Message)
  message.attach(msg)

  path = FileName
  content_type, encoding = mimetypes.guess_type(path)

  if content_type is None or encoding is not None:
    content_type = 'application/octet-stream'
  main_type, sub_type = content_type.split('/', 1)
  if main_type == 'text':
    fp = open(path, 'rb')
    msg = MIMEText(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'image':
    fp = open(path, 'rb')
    msg = MIMEImage(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'audio':
    fp = open(path, 'rb')
    msg = MIMEAudio(fp.read(), _subtype=sub_type)
    fp.close()
  else:
    fp = open(path, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()

  msg.add_header('Content-Disposition', 'attachment', filename=FileName)
  message.attach(msg)

  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def SendMessage(Subject,Message,FileName,To_email):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.

  """
  credentials = get_credentials()
  http = credentials.authorize(httplib2.Http())
  service = discovery.build('gmail', 'v1', http=http)
  message = CreateMessageWithAttachment(Subject,Message,FileName,To_email)
  
  message = (service.users().messages().send(userId="me", body=message).execute())
  print ('Message Id: %s' % message['id'])
  return message


def Check(filePath,emailID):
	FieldId,Employees,TotalEmployee,ErrorReport = EmployeeData(filePath)
	LevelTitle,root = GFDTstructure(ProductionTemplateFilePath)
	ErrorReport += GFDTStructureVerifier(LevelTitle,root,FieldId,Employees)
	if len(ErrorReport) == 0:
		sfInsert(FieldId,Employees,TotalEmployee)
		gcs.delete(filePath)
		return 1

	else:
		FileName = XlsxErrorReport(ErrorReport,filePath)
		Message = open("ErrorEmailBody.txt").read()
		Subject = "Success Factor Upload File Error"
		mail.send_mail(sender=sender_email_id.format(
                app_identity.get_application_id()),
                to=emailID,
                subject=Subject,
                body=Message,attachments=[(FileName, gcs.open(FileName).read())])
		gcs.delete(FileName)
		return 0
		
	
	
	

	
