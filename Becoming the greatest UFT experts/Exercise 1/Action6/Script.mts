
If UserIsLoggedIn() Then
	BookingNumber = DataTable("BookingNumber", dtLocalSheet)
	ExpectedResult = DataTable("ExpectedResult", dtLocalSheet)
	TicketPrice = DataTable("TicketPrice", dtLocalSheet)
	Set ExpectedError = CreateObject("Scripting.Dictionary")
	ExpectedError.Add "Title", ""
	ExpectedError.Add "PassTitle", ""
	ExpectedError.Add "FailTitle", ""
	ExpectedError.Add "PassDesc", ""
	ExpectedError.Add "FailDesc", ""
	ExpectedError.Add "DialogName" , ""
	
	ExecuteFlightUpdating()
Else
	Reporter.ReportEvent micGeneral, "User Not Logged in", "User Not Logged in, booking is not avaialable"
	ExitActionIteration
End If

Function ExecuteFlightUpdating()

	WpfWindow("Micro Focus MyFlight Sample").WpfTabStrip("WpfTabStrip").Select "SEARCH ORDER" @@ hightlight id_;_1902083024_;_script infofile_;_ZIP::ssf1.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfRadioButton("byNumberRadio").Set
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("byNumberWatermark").Set CStr(BookingNumber)
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("SEARCH").Click
	
	Select Case ExpectedResult
		Case "Not a Number"
			ExpectedError.Item("Title") = "Not a Number"
			ExpectedError.Item ("PassTitle") = "Not a number Dialog displayed"
			ExpectedError.Item ("FailTitle") = "Not a number Dialog was not displayed"
			ExpectedError.Item ("PassDesc") = "Not a number Dialog displayed"
			ExpectedError.Item ("FailDesc") = "Not a number Dialog was not displayed"
			ExpectedError.Item ("DialogName") = "Error"
			CheckErrorDialog(ExpectedError)
			Exit Function
			
   		Case "Negative Number"
   			ExpectedError.Item("Title") = "Negative Number"
			ExpectedError.Item ("PassTitle") = "Negative number Dialog displayed"
			ExpectedError.Item ("FailTitle") = "Negative number Dialog was not displayed"
			ExpectedError.Item ("PassDesc") = "Negative number Dialog displayed"
			ExpectedError.Item ("FailDesc") = "Negative number Dialog was not displayed"
			ExpectedError.Item ("DialogName") = "Error"
			CheckErrorDialog(ExpectedError)
			Exit Function
			
		Case "Does not Exist"
			ExpectedError.Item("Title") = "Does not Exist"
			ExpectedError.Item ("PassTitle") =  "Order does not exist Dialog displayed"
			ExpectedError.Item ("FailTitle") = "Order does not exist Dialog was not displayed"
			ExpectedError.Item ("PassDesc") = "Order does not exist Dialog displayed"
			ExpectedError.Item ("FailDesc") = "Order does not exist Dialogwas not displayed"
			ExpectedError.Item ("DialogName") = "Error"
			CheckErrorDialog(ExpectedError)
			Exit Function
    		
	End Select
	
	SelectedNumberOfTickets = WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTicketsCombo").GetSelection
	If SelectedNumberOfTickets*2 > 99 Then
		Reporter.ReportEvent micGeneral, "Cannot double number of ticket ", "Number of tickets cannot be more than 99"
		WpfWindow("Micro Focus MyFlight Sample").WpfButton("NEW SEARCH").Click
		Exit Function
	End If
 @@ hightlight id_;_1916831048_;_script infofile_;_ZIP::ssf10.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTicketsCombo").Select CStr(SelectedNumberOfTickets * 2) @@ hightlight id_;_1916831048_;_script infofile_;_ZIP::ssf15.xml_;_
	UpdatedTicketsNbr = WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTicketsCombo").GetSelection
	ExpectedTotal = UpdatedTicketsNbr * TicketPrice
	ActualTotal  = WpfWindow("Micro Focus MyFlight Sample_2").WpfObject("TotalPrice").GetROProperty("text") @@ hightlight id_;_2061514232_;_script infofile_;_ZIP::ssf34.xml_;_
	
	Print ("Actual: "& CDbl(ActualTotal) &" Expected: " & CDbl(ExpectedTotal))
	Print ("Without Converting, Actual: "& ActualTotal &" Expected: " & ExpectedTotal)
	
	If CDbl(ActualTotal ) =  CDbl(ExpectedTotal) Then
		Reporter.ReportEvent micPass, "Correct price Displayed", "Correct price Displayed"
	Else
		Reporter.ReportEvent micFail, "Incorrect price Displayed", "Incorrect price Displayed"	
	End If
	
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("updateBtn").Click
	
	If WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order updated").Exist(1000) Then
		Reporter.ReportEvent micPass, "Flight Updated Correctly", "Flight Updated Correctly"
	Else
		Reporter.ReportEvent micFail, "Failed to update flight", "Failed to update flight"
	End If
	
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("NEW SEARCH").Click @@ hightlight id_;_1916831528_;_script infofile_;_ZIP::ssf22.xml_;_

End Function

Function CheckErrorDialog (ExpectedError)
	If WpfWindow("Micro Focus MyFlight Sample").Dialog(ExpectedError("DialogName")).Exist(1000) Then
		Reporter.ReportEvent micPass, ExpectedError("PassTitle") , ExpectedError("PassDesc")
	Else
		Reporter.ReportEvent micFail, ExpectedError("FailTitle"), ExpectedError ("FailDesc") 
		ExitActionIteration
	End If @@ hightlight id_;_1442642_;_script infofile_;_ZIP::ssf29.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").Dialog(ExpectedError("DialogName")).WinButton("OK").Click
        WpfWindow("Micro Focus MyFlight Sample_2").WpfTabStrip("WpfTabStrip").Select "BOOK FLIGHT"
End Function

Function UserIsLoggedIn()
	UserIsLoggedIn = Environment.Value("LoggedIn")
End Function
