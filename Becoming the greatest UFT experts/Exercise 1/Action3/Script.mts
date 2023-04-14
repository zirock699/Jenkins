DatatablePath = "C:\BookedFlights.xls"
 
SaveDatatableChanges() 
Logout()

Function  Logout()
	Environment.Value("LoggedIn") = false
	Environment.Value("EnvUserName") = nil
	WpfWindow("Micro Focus MyFlight Sample").Close
End Function

Function SaveDatatableChanges()
	datatable.Export(DatatablePath)
	Print ("Datatable saved to: " & DatatablePath)
End Function
