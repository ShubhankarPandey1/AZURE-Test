''******************************************
''If Parameter("Card_Number")<>"" Then
''	
''
''	Else
'		'create exel object
'		Set myxl=CreateObject("excel.application")
'		'load exel
'		Set myWB = myxl.Workbooks.Open ("C:\"&Parameter("FileName"))
'		'set visibility
'		myxl.Application.Visible=true
'		'set worksheet
'		SheetName = Parameter("SheetName")
'		Set objworksheet=myxl.Application.Worksheets(SheetName)
'		objworksheet.Activate
'		
'		
'		
'		Row=objworksheet.UsedRange.Rows.count
'		Col=objworksheet.UsedRange.Columns.count
'		
'		If objworksheet.Range("N1").Value="" Then
'			objworksheet.Columns("N:N").Insert
'			objworksheet.Range("N1").Value = "Used Card"
'		End If
'		
'		
'		
'		For i = 2 To Row
'			
'			If objworksheet.Cells(i,14).Value = "" Then
'			
'				If objworksheet.Cells(i,13).Value <> "" Then
'					CardNumberToUse = objworksheet.Cells(i,2)
'					CardNameToUse =  objworksheet.Cells(i,1)
'					objworksheet.Cells(i,14).Value = "YES" 
'					Parameter("Card_Number") = CardNumberToUse
'					Parameter("O_Card_Name") = CardNameToUse
'					Exit For
'				End If
'		
'			End If
'			
'		Next
'		
'		myxl.ActiveWorkbook.Save
'		myxl.ActiveWorkbook.Close
'		myxl.Application.Quit
''End If
''*******************************************
'Dim myxl1,objWorkbook1,objworksheet1
'
'Set myxl1=CreateObject("excel.application")
''load exel
'Set objWorkbook1 = myxl1.Workbooks.Open ("C:\Users\NHIDCL\Documents\FlightLogin.xlsx")
''set visibility
'myxl1.Application.Visible=true
''set worksheet
'Set objworksheet1=myxl1.Application.Worksheets("Sheet1")
'objworksheet1.Activate

'username = objworksheet1.Cells(2,1)
'varPassword =  objworksheet1.Cells(2,2)

'Row=objworksheet1.UsedRange.Rows.count
'Col=objworksheet1.UsedRange.Columns.count

DataTable.ImportSheet "C:\Users\NHIDCL\Documents\FlightLogin.xlsx", 1, "Global"
row = DataTable.getsheet(1).GetRowCount

username = DataTable("Username", dtGlobalSheet)'"john"
varPassword = DataTable("Password", dtGlobalSheet)'"HP"
EncryptedPWD = Crypt.Encrypt(varPassword)


LoadFunctionLibrary "C:\Users\NHIDCL\Documents\Unified Functional Testing\Library1.qfl"

Call Login(username,EncryptedPWD)


 @@ hightlight id_;_2056644480_;_script infofile_;_ZIP::ssf4.xml_;_
'WpfWindow("devname:= Micro Focus MyFlight Sample Application").WpfEdit("devname:= agentName").Set "john"
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select "Frankfurt" @@ hightlight id_;_2056645056_;_script infofile_;_ZIP::ssf6.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfImage("WpfImage").Click 13,8 @@ hightlight id_;_2056645104_;_script infofile_;_ZIP::ssf7.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfCalendar("Su").SetDate "17-Feb-2020" @@ hightlight id_;_2056719848_;_script infofile_;_ZIP::ssf8.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_1903545024_;_script infofile_;_ZIP::ssf9.xml_;_
WpfWindow("Micro Focus MyFlight Sample").Dialog("Error").WinButton("OK").Click @@ hightlight id_;_395438_;_script infofile_;_ZIP::ssf10.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select "Business" @@ hightlight id_;_2056721912_;_script infofile_;_ZIP::ssf14.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfImage("WpfImage").Click 11,11 @@ hightlight id_;_2056645104_;_script infofile_;_ZIP::ssf15.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfCalendar("Su").SetDate "25-Feb-2020" @@ hightlight id_;_2056719848_;_script infofile_;_ZIP::ssf16.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_1903545024_;_script infofile_;_ZIP::ssf17.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,2 @@ hightlight id_;_2056720760_;_script infofile_;_ZIP::ssf18.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 1,2 @@ hightlight id_;_2056720760_;_script infofile_;_ZIP::ssf19.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 6,0 @@ hightlight id_;_2056720760_;_script infofile_;_ZIP::ssf20.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click @@ hightlight id_;_1903550400_;_script infofile_;_ZIP::ssf21.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfObject("$330.80").Click 22,14 @@ hightlight id_;_1922373528_;_script infofile_;_ZIP::ssf22.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfObject("$330.80").Check CheckPoint("$330.80") @@ hightlight id_;_1922374336_;_script infofile_;_ZIP::ssf23.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfObject("$330.80").Output CheckPoint("$330.80_2") @@ hightlight id_;_1922379184_;_script infofile_;_ZIP::ssf25.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("NEW SEARCH").Click @@ hightlight id_;_2056720568_;_script infofile_;_ZIP::ssf28.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_1903541136_;_script infofile_;_ZIP::ssf29.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 3,3 @@ hightlight id_;_1903551792_;_script infofile_;_ZIP::ssf30.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click
