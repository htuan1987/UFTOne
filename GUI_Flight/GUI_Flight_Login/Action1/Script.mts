datatable.ImportSheet "C:\Users\Administrator\UFTtest\test.xlsx",2,"Global"
n=datatable.GetSheet("Global").GetRowCount

For i  = 1 To n Step 1
	datatable.SetCurrentRow(i)

systemutil.Run "C:\Program Files (x86)\Micro Focus\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe" @@ hightlight id_;_1116312_;_script infofile_;_ZIP::ssf3.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set datatable("Username")
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").Set datatable("Password") @@ hightlight id_;_1930939960_;_script infofile_;_ZIP::ssf6.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_5506142_;_script infofile_;_ZIP::ssf8.xml_;_
WpfWindow("Micro Focus MyFlight Sample").Close
Next
