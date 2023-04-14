
If UserIsLoggedIn() Then
  	TableUserName = DataTable("Username", dtLocalSheet)
	LoggedInUser = Environment.Value("EnvUserName")
	PassengerName = DataTable("PassengerName", dtLocalSheet)
	FlightFrom = DataTable("Departure", dtLocalSheet)
	Destination = DataTable("Arrival", dtLocalSheet)
	FlightDate = DataTable("Date", dtLocalSheet)
	ClassType = DataTable("Class", dtLocalSheet)
	NbrTickets = DataTable("NrOfTickets", dtLocalSheet)
	FlightNbr = DataTable("Flight", dtLocalSheet)
	Price = DataTable("Price", dtLocalSheet)
	PricePerTicket = Price / NbrTickets
	DataGridColumnFlightNbr = 4
	DataGridColumnPrice = 0
		
    	ExecuteFlightBooking()    	
Else
	Reporter.ReportEvent micGeneral, "User Not Logged in", "User Not Logged in, booking is not avaialable"
	ExitActionIteration
End If


Function ExecuteFlightBooking()
If  StrComp(LoggedInUser, TableUserName, vbTextCompare) = 0 Then
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select FlightFrom
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select Destination @@ hightlight id_;_1954986256_;_script infofile_;_ZIP::ssf14.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfCalendar("datePicker").SetDate FlightDate
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select(ClassType) @@ hightlight id_;_1954960384_;_script infofile_;_ZIP::ssf23.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTickets").Select NbrTickets @@ hightlight id_;_1954988896_;_script infofile_;_ZIP::ssf27.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_1897725744_;_script infofile_;_ZIP::ssf79.xml_;_
	FromDeparture = WpfWindow("Micro Focus MyFlight Sample").WpfObject("From").GetROProperty("text") @@ hightlight id_;_1912645848_;_script infofile_;_ZIP::ssf80.xml_;_
	ToDestination = WpfWindow("Micro Focus MyFlight Sample").WpfObject("To").GetROProperty("text")
	
	If InStr(FromDeparture, FlightFrom) > 0 Then
		Reporter.ReportEvent micPass, "Correct departure displayed", "Correct departure displayed"
	Else
		Reporter.ReportEvent micFail, "Incorrect Departure Displayed", "Incorrect Departure Displayed"
	End If
	
	If InStr(ToDestination, Destination) > 0 Then
		Reporter.ReportEvent micPass, "Correct Destination displayed", "Correct Destination displayed"
	Else
		Reporter.ReportEvent micFail, "Incorrect Destination Displayed", "Incorrect Destination Displayed"
	End If
	
	If FlightIsNotAvailable (FlightNbr) Then
		Reporter.ReportEvent micGeneral, "Flight not available", "Flight not available"
		WpfWindow("Micro Focus MyFlight Sample").WpfButton("BACK").Click @@ hightlight id_;_2076107008_;_script infofile_;_ZIP::ssf70.xml_;_
		Exit Function
	End If
	
	WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell GetFlightRowFor(FlightNbr), DataGridColumnFlightNbr
	
	If GetDisplayedPrice() =  PricePerTicket Then
		Reporter.ReportEvent micPass, "Correct Price Displayed", "Correct Price Displayed"
	Else
		Reporter.ReportEvent micFail, "Incorrect Price Displayed", "Incorrect Price Displayed"
	End If
	
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set PassengerName @@ hightlight id_;_2100444184_;_script infofile_;_ZIP::ssf62.xml_;_
	DisplayedTotalPrice = WpfWindow("Micro Focus MyFlight Sample").WpfObject("totalPrice").GetROProperty("text")
	Print (DisplayedTotalPrice &" "& TypeName(DisplayedTotalPrice))
	print (Price &" "& TypeName(Price))
	If  CDbl(DisplayedTotalPrice) = CDbl(Price) Then
		Reporter.ReportEvent micPass, "Displayed price is correct", "Displayed price is correct"
		else
		Reporter.ReportEvent micFail, "Incorrect price is Displayed", "Incorrect price is Displayed"
	End If
	
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click

	If WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order completed").Exist(1000) Then
		Reporter.ReportEvent micPass, "Flight Booked Correctly", "Flight Booked Correctly"
	Else
		Reporter.ReportEvent micFail, "Failed to book flight", "Failed to book flight"
	End If

	OrderComfirmationMessage = WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order completed").GetROProperty("text")
	FlightBookingNumber  = ExtractNumberFromString(OrderComfirmationMessage)
	SaveBookingNbr(FlightBookingNumber)
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("NEW SEARCH").Click	
Else  
	Reporter.ReportEvent micGeneral, "User Not allowed to perform this action", "User trying to use other username to book a flight"
End If
End Function

Function SaveBookingNbr(FlightBookingNumber)
	DataTable.LocalSheet.SetCurrentRow(DataTable.LocalSheet.GetCurrentRow) 
	DataTable.Value("BookedFlightReference", dtLocalSheet) = FlightBookingNumber
End Function

Function GetDisplayedPrice()
	CellPrice = WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").GetCellData(GetFlightRowFor(FlightNbr),DataGridColumnPrice)
	GetDisplayedPrice = CDbl(Mid(CellPrice, 4))
End Function

Function FlightIsNotAvailable(FlightNbr)
	If GetFlightRowFor(FlightNbr) = -1 Then
		FlightIsNotAvailable = true
		Exit function
		else
		FlightIsNotAvailable = false
	End If
End Function

Function GetFlightRowFor(FlightNbr)
RowCount = WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").RowCount
For i = 0 To RowCount -1
	CellValue = WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").GetCellData(i,4)
	   If CellValue = FlightNbr Then  
	      GetFlightRowFor = i
	   Exit For
	   End If
	   GetFlightRowFor = -1
Next
End Function

Function UserIsLoggedIn()
	UserIsLoggedIn = Environment.Value("LoggedIn")
End Function

Function GetLoggedInUser()
	GetLoggedInUser = Environment.Value("EnvUserName")
End Function

Function ExtractNumberFromString(text)
    Dim regex, matches, match
    Set regex = New RegExp
    regex.Pattern = "\d+"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(text)
    If matches.Count > 0 Then
        Set match = matches.Item(0)
        ExtractNumberFromString = match.Value
    Else
        ExtractNumberFromString = ""
    End If
End Function
