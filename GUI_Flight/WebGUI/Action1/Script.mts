'Browser("Directory listing for").Page("Directory listing for").Link("/0").Click

'msgbox Browser("Directory listing for").Page("Directory listing for").Object.getElementsByTagName("A").length
''Link("/10").Click @@ script infofile_;_ZIP::ssf2.xml_;_
''Browser("Directory listing for").Page("Index of /10/").Link("lost+found/").Click @@ script infofile_;_ZIP::ssf3.xml_;_
''Browser("Directory listing for").Page("Item not available").Link("Home").Click @@ script infofile_;_ZIP::ssf4.xml_;_
'Function EnumerateApp(ParentObj, Desc, OperationMethod, PostOperationMethod, RestoreMethod)
'	dim ObjCol, CurrentObj, idx
'	idx = 0
'	' Retrieve a collection of all the objects of the given description
'	Set ObjCol = ParentObj.ChildObjects(Desc)
'
'	Do While (idx < ObjCol.Count)
'		' Get the current object
'		set CurrentObj = ObjCol.item(idx)
'
'		' Perform the desired operation on the object
'		eval("CurrentObj." & OperationMethod)
'
'		' Perform the post operations (after the object operation)
'		eval(PostOperationMethod & "(ParentObj, CurrentObj)")
'
'		' Return the application to the original state
'		eval(RestoreMethod & "(ParentObj, CurrentObj)")
'
'		idx = idx + 1
'		' Retrieve the collection of objects
'		' (Since the application might have changed)
'		Set ObjCol = ParentObj.ChildObjects(Desc)
'	Loop
'End Function
'
'' ********************************** An Example of usage **********************
'' Report all the pages referred to by the current page
'' ***********************************************************************************
'
'Function ReportPage(ParentObj, CurrentObj)
'	dim FuncFilter, PageTitle
'
'	PageTitle = ParentObj.GetROProperty("title")
'	FuncFilter = Reporter.Filter
'	Reporter.Filter = 0
'	Reporter.ReportEvent 0, "Page Information", "page title " & PageTitle
'	Reporter.Filter = FuncFilter
'End Function
'
'Function BrowserBack(ParentObj, CurrentObj)
'	BrowserObj.Back
'End Function
'
'' Save the Report Filter mode
'OldFilter = Reporter.Filter
'Reporter.Filter = 2 ' Enables Errors Only
'
' Create the description of the Link object
Function returnLastLink(ParentObj, Desc)
	Set ObjCol = ParentObj.ChildObjects(Desc)
	msgbox ObjCol(ObjCol.Count-1).GetROProperty("type")
End Function
Set Desc = Description.Create()
'Desc("html tag").Value = "A"
Desc("text").Value = "India"
'Desc("Class Name").Value = "Link"

Set BrowserObj = Browser("Directory listing for")
Set PageObj = BrowserObj.Page("Directory listing for")
Print 
'Set ObjCol = PageObj.ChildObjects(Desc)
'msgbox ObjCol.Count
'
''' Start the enumeration
''call EnumerateApp(PageObj, Desc, "Click", "ReportPage", "BrowserBack")
''
''Reporter.Filter = OldFilter ' Returns the original filter
'For i = 0 To ObjCol.Count-1 Step 1
'	msgbox ObjCol(i).GetROProperty("text")
'Next


call returnLastLink(PageObj,Desc)

'Browser("Micro Focus UFT Agent").Page("Micro Focus UFT Agent").Sync
'Browser("Extensions - Micro Focus").Page("Extensions - Micro Focus").Sync
'Browser("Gmail").Page("Gmail").WebEdit("Email or phone").Set
'Browser("Directory listing for").Page("Index of /35/").Link("lost+found/").Click

