import openpyxl
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
ProductionTemplateFilePath = "ProductionTemplate.tsv"
api_url = "https://api10preview.sapsf.com/odata/v2/"
userName = "SFADMIN"
companyID = "C0017935023D"
password = "SFADMIN123"
datemode = ''
DuplicateCheckIndex = ["Date of Birth*","Date of Joining*","Fathers Name*","FirstName*","LastName*"]

class Node:
	def __init__(self):
		self.child = {}
		self.code = "Invalid"


def XlsxToTsv(FilePath):
	reload(sys)
	sys.setdefaultencoding('utf-8')
	wb = xlrd.open_workbook(FilePath)
	sh = wb.sheet_by_index(0)
	global datemode
	datemode = wb.datemode
	FilePathTsv= FilePath.split(".xlsx")[0]
	csvFile = open(FilePathTsv+'.tsv', 'wu')
	wr = csv.writer(csvFile,delimiter='\t')

	for rownum in xrange(sh.nrows):
        	wr.writerow(sh.row_values(rownum))
	csvFile.close()
#	os.remove(FilePath)
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
	fileHandle = open(XlsxToTsv(FilePath),"rU")
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
	return (LevelTitle,Output,TotalRecord,ErrorReport)
def GFDTstructure(filePath):
	fileHandle = open(filePath,"ru")
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
	tsv_file = ReportName+"Error" + ".tsv"
	xlsx_file = ReportName+"Error" + ".xlsx"


	tsvHandle = open(tsv_file,"w")
	tsvHandle.write("Error\n")
	for error in Errors:
		tsvHandle.write(error+"\n")
	tsvHandle.close()

	
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
		sys.exit()


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
#    UserPayload["username"] = userID
#    UserPayload["hr"] = "User('9090')"
#    UserPayload["manager"] = "User('9090')"
    Userauth = HTTPBasicAuth(userName+"@"+companyID,password)
    Meta = {}
    Meta["uri"] ="User(userId='"+userID+"')"
    UserPayload["__metadata"]=Meta
    UserHeader = {}
    UserHeader["content-type"] ="application/json; charset=utf-8"
    UserHeader["accept"] = "application/json"
    r = requests.post(api_url+"upsert",auth = Userauth,headers = UserHeader,  data = json.dumps(UserPayload))
    print r.text


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
    print r.text
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
    print r.text        			
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
    EmpJobPayload["userId"] = userID
    Userauth = HTTPBasicAuth(userName+"@"+companyID,password)
    Meta = {}
    Meta["uri"] ="EmpJob"
    EmpJobPayload["__metadata"]=Meta
    EmpJobHeader = {}
    EmpJobHeader["content-type"] ="application/json; charset=utf-8"
    EmpJobHeader["accept"] = "application/json"
    r = requests.post(api_url+"upsert",auth = Userauth,headers = EmpJobHeader,  data = json.dumps(EmpJobPayload))
    print r.text
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
    print r.text
 #   sys.exit()        			
		


def Check(filePath,emailId):
	FieldId,Employees,TotalEmployee,ErrorReport = EmployeeData(filePath)
	LevelTitle,root = GFDTstructure(ProductionTemplateFilePath)
	ErrorReport += GFDTStructureVerifier(LevelTitle,root,FieldId,Employees)
	if len(ErrorReport) != 0:
		sfInsert(FieldId,Employees,TotalEmployee)

	else:
		FileName = XlsxErrorReport(ErrorReport,filePath)
		Message = open("ErrorEmailBody.txt").read()
		Subject = "Success Factor Upload File Error"
		send_mail(Subject,Message,FileName,emailId)
#		os.remove(filePath+".xlsx")
	
	
	

	
#Fieldid,Employees,TotalEmployee,ErrorReport = EmployeeData("testData.xlsx")
#LevelTitle,root = GFDTstructure("ProductionTemplate.tsv")
Check("testData.xlsx","raj.jha@gmail.com")
#MasterFileTuples("testData.xlsx")
#print len(GFDTStructureVerifier(LevelTitle,root,Fieldid,Employees)),len(ErrorReport)
