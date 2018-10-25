'To terminate all the processes in the machine
Call KillProcess("UFT.exe")
Call KillProcess("QtpAutomationAgent.exe")
Call KillProcess("iexplore.exe")
Call KillProcess("chrome.exe")
Call KillProcess("firefox.exe")
Call KillProcess("werfault.exe")
 
'Create QTP object
Set QTP = CreateObject("QuickTest.Application")
ConsoleOutput("Launching QTP Application")
QTP.Launch
QTP.Visible = TRUE
 
'Open QTP Test
ConsoleOutput("Opening Test....")
QTP.Open "C:\GITHUB\tests\GUITest2", TRUE 'Set the QTP test path
 
'Set Result location
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtpResultsOpt.ResultsLocation = "C:\GITHUB\tests\GUITest2\Res1" 'Set the results location
 
'Run QTP test
ConsoleOutput("Starting to run....")
QTP.Test.Run qtpResultsOpt
 
'Close QTP
ConsoleOutput("Execution Completed Successfully!!!!!!!!!!")
QTP.Test.Close
ConsoleOutput("Terminating QTP....")
QTP.Quit
 

Sub KillProcess(ByVal ProcessName)
	
	On Error Resume Next
	
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Dim colProcesses : Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" &  ProcessName & "'")

	ConsoleOutput("Terminating Process : " & ProcessName)
	
	For Each objProcess in colProcesses
		intTermProc = objProcess.Terminate
	Next
	
	On Error GoTo 0
	
End Sub


Sub ConsoleOutput(ByVal MessageToBeDisplayed)
	WScript.StdOut.WriteLine Time() & " :: " & MessageToBeDisplayed
End Sub