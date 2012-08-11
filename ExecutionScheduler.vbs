'Execute Test Set-'Training_PNT', 

'QC URL

strQCURL= "http://ec2-23-21-149-213.compute-1.amazonaws.com:8080/qcbin"

'QC Domain & Project

strQCDomain="FIDO"

strQCProject="WindowsClient"

'QC User Credentials

strQCUser="svejendla"

strQCPassword="satyam123$"

'QC Api- class name

'OTA-QC-"TDApiOle80.TDConnection"

Dim QCObj
Set QCObj=CreateObject("TDApiOle80.TDConnection")

QCObj.InitConnectionEx strQCURL

QCObj.Login strQCUser,strQCPassword

QCObj.Connect strQCDomain,strQCProject

'Select Test Set to Run

Set objQCTreeMgr=QCObj.TestSetTreeManager
'QC Tree-objQCTreeMgr
'QC Test Set Folder-objQCTestSetFolder
'Test Set-objQCTestSet
strTestSetFolder="Root\FIDO_Windows_Client"
strTestSetname="Sample"
Set objQCTestSetFolder=objQCTreeMgr.NodeByPath(strTestSetFolder)
Set objQCTestSet=objQCTestSetFolder.FindTestSets(strTestSetname)

'Loop Statement 

intCnt=1

While intCnt <= objQCTestSet.Count' Value 2-We have Test Cases in Test Set
	'Select the Test Case in Test Set
	Set objTestCase=objQCTestSet.Item(intCnt)
	If objTestCase.Name=strTestSetname  Then
		intCnt=objQCTestSet.Count+1
	End If
	intCnt=intCnt+1
Wend



'Run Test Set and Assign Execution Results

'Run Test Case- Locally or Remote Host name

'Assign the Test Set Run and Scehdule Test Set

Set objQCScheduler=objTestCase.StartExecution("")
'Option Parameters in start Execution
objQCScheduler.RunAllLocally=True
objQCScheduler.Run

'Execution Results
Set objQCExecutionStatus=objQCScheduler.ExecutionStatus

'Time taken to execute the test case

While objQCExecutionStatus.Finished=False
	objQCExecutionStatus.RefreshExecStatusInfo "all",True
	If objQCExecutionStatus.Finished=False Then
		WScript.sleep 5
	End If	
Wend