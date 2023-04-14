UserName = DataTable("Username", dtGlobalSheet)
Password = DataTable("Password", dtGlobalSheet)
ExpectedResult = DataTable("Expected", dtGlobalSheet)
PasswordMaskChar = "●"

' type UserName and Password
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set(UserName) @@ hightlight id_;_1928075008_;_script infofile_;_ZIP::ssf2.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").Set(Password)
VerifyPasswordIsMasked()
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click
VerifyAuthenticationWindows()


 @@ hightlight id_;_394506_;_script infofile_;_ZIP::ssf11.xml_;_
WpfWindow("Micro Focus MyFlight Sample").Close()

Function VerifyAuthenticationWindows()
If  StrComp(ExpectedResult, "Login not ok", vbTextCompare) = 0 Then

	IF WpfWindow("Micro Focus MyFlight Sample").Dialog("Login Failed").Exist(2) THEN
	Reporter.ReportEvent micPass, "Login Failed Dialog Displayed", "Login Failed Dialog Displayed"
	WpfWindow("Micro Focus MyFlight Sample").Dialog("Login Failed").Close()
	Else
 	Reporter.ReportEvent micFail, "Login Failed Dialog Is Not Displayed", "Login Failed Dialog Is Not  Displayed"
	End  If
	
ElseIf StrComp(ExpectedResult, "Login ok", vbTextCompare) = 0 Then @@ hightlight id_;_2084940960_;_script infofile_;_ZIP::ssf21.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfObject("Hello").Check CheckPoint("Hello")
	
End If


 @@ hightlight id_;_2084940960_;_script infofile_;_ZIP::ssf21.xml_;_
	
End Function

Function VerifyPasswordIsMasked()
	VisiblePasswordFieldValue = WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").GetROProperty("text")
	If IsPasswordMasked(VisiblePasswordFieldValue) Then
	Reporter.ReportEvent micPass, "Password is Masked", "Password is Masked"
	Else
 	Reporter.ReportEvent micFail, "Password is Not Masked", "Password is Not Masked"
	End If
End Function

Function IsPasswordMasked(VisiblePasswordFieldValue)
	If AllCharsSame (VisiblePasswordFieldValue, PasswordMaskChar) AND Len(VisiblePasswordFieldValue) = Len (Password) Then
		  IsPasswordMasked = true
	End If
End Function

Function AllCharsSame(str, char)
    For i = 1 To Len(str)
        If Mid(str, i, 1) <> char Then
            AllCharsSame = False
            Exit Function
        End If
    Next
    AllCharsSame = True
End Function


