'display built-in environment variables, can be viewed in File/Settings.../Environment
'msgbox environment("ProductName")
'msgbox environment("ProductVer")
'can define user-defined variables to use with UFT: for example URL = www.google.com.vn
'systemutil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe", environment.Value("URL")

'use datatable method importsheet to retrieve data from excel file: test.xlsx sheet number 2 to Global UFT Data sheets
datatable.ImportSheet "C:\Users\Administrator\UFTtest\test.xlsx",2,"Action1"
'Note: top row will always be used for column Titles
'GetRowCount method is used to get the last data row except top row
n=datatable.GetSheet("Action1").GetRowCount

'for loop is used to test iteration
'Record & run settings should be set to run on any opened Win app
For i  = 1 To n Step 1
	datatable.SetCurrentRow(i)
'-----Note that for this iteration test, systemutil.Run is used to recall the GUI program after it is closed at the end-----
systemutil.Run "C:\Program Files (x86)\Micro Focus\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe" @@ hightlight id_;_1116312_;_script infofile_;_ZIP::ssf3.xml_;_
'datatable object sheet 2 = Action1 --> contains Username & Password
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set datatable("Username",2)
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").Set datatable("Password",2) @@ hightlight id_;_1930939960_;_script infofile_;_ZIP::ssf6.xml_;_
wait(1)
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_5506142_;_script infofile_;_ZIP::ssf8.xml_;_
'add checkpoint to compare
'WpfWindow("Micro Focus MyFlight Sample").Dialog("Login Failed").Static("Username must be at least").Check CheckPoint("Username must be at least 4 characters long") @@ hightlight id_;_2033816_;_script infofile_;_ZIP::ssf24.xml_;_
'wait(1)
WpfWindow("Micro Focus MyFlight Sample").Close
'next command instructs the loop to continue
Next
