'
' This script will check all fields and aspects for the create new website dialog
' This means:	- checkPort
'			- checkHostname
' 			- checkDescription
'			- checkForConflictWithExistingWebsite
'
' And for the use existing website dialog
'			- setIISWebsiteData

Function checkPort( ) 
    Dim oRegularExpression, iPort, iMinPort, iMaxPort, iReturn
    
    On Error Resume Next 
      
    ' Get iPort from WIX
    ' Port has to Lng
    If Not IsEmpty(Session.Property("WEBSITEPORT")) Then
      iPort = CLng(Session.Property("WEBSITEPORT"))
    End If
    
    ' Set default result
    Session.Property("checkNewWebsitePort") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Set min and max values
    iMinPort = 1
    iMaxPort = 65535
    
    ' Prepare regex validation for integer value
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = "^[0-9]+$"
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate if iPort is a number and between min and max values
  	If Not ((oRegularExpression.Test(iPort)) And (iPort >= iMinPort) And (iPort <= iMaxPort)) Then

    	' Port is either not an integer or not in a valid range
    	MsgBox "The specified website port is invalid - please ensure that the value is an integer and between "  & _ 
    			   iMinPort & " and " & iMaxPort & "!", vbCritical, "Error" 
      
      Session.Property("checkNewWebsitePort") = "0"
      iReturn = ERROR_INSTALL_FAILURE 
    	
    End If
     
    checkPort = iReturn
End Function

Function checkHostname() 
    Dim oRegularExpression, iReturn, sHostheader
    
    ' Get sHostheader
    sHostheader = Session.Property("WEBSITEHOSTHEADER")
    
    ' Set default result
    Session.Property("checkNewWebsiteHostname") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Hostheader can be empty -> IP:Port binding
    If Not IsEmpty(sHostheader) Then
    	
      ' Prepare regex validation for hostheader
  	  Set oRegularExpression = New RegExp
  	  oRegularExpression.Pattern = "^(([A-Z0-9]+[A-Z0-9_-]*)(\.[A-Z0-9][A-Z0-9-]*)+)*$"
  	  oRegularExpression.IgnoreCase = True
  	
  	  ' Validate hostheader
  	  If Not oRegularExpression.Test(sHostheader) Then

    	  ' Port is either not an integer or not in a valid range
    	  MsgBox "The specified website hostheader value is invalid - please ensure that the value "  & _ 
    			     " is a valid hostname (e.g. www.mysite.com) and - if necessary - convert IDN domains!", vbCritical, "Error" 
      
        Session.Property("checkNewWebsiteHostname") = "0"
        iReturn = ERROR_INSTALL_FAILURE 
    	
      End If
    End If
    
    checkHostname = iReturn
End Function

Function checkDescription() 
    Dim oRegularExpression, iReturn, sDescription
    
    ' Get sDescription
    sDescription = Session.Property("WEBSITEDESCRIPTION")
    
    ' Set default result
    Session.Property("checkNewWebsiteDescription") = "1"
    iResult = ERROR_SUCCESS 
    
    
    ' Prepare regex validation for description
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = "^[a-zA-Z0-9,\.-_\s\(\)\[\]]+$"
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate website description
  	If Not oRegularExpression.Test(sDescription) Then

    	' Msgbox if regex failed
    	MsgBox "The specified website description is invalid - please ensure that the value only contains the " & vbCrLf & _
    	       "characters from A to Z,.-_()[] and the Numbers from 0 to 9! ", vbCritical, "Error" 
      
      Session.Property("checkNewWebsiteDescription") = "0"
      iReturn = ERROR_INSTALL_FAILURE 
    	
    End If
     
    checkDescription = iReturn
End Function

Function checkForConflictWithExistingWebsite ()
    Dim iPort, sIP, sHostHeader, iReturn
    Dim oThisServer, oWebsite, sBindings
    Dim bConflict
    
     On Error Resume Next 
    
    ' Get iPort, sIP, sHostHeader from WIX
    iPort 		= Session.Property("WEBSITEPORT")
    sIP 		= Session.Property("WEBSITEIP")
    sHostheader = Session.Property("WEBSITEHOSTHEADER")

    ' Clear IP if sIP equals All
    If sIP = "AllUnassigned" Then
       sIP = ""
    End If

    ' Set default result
    Session.Property("checkNewWebsiteConflict") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Set Conflict to false
    bConflict = False
    
   ' Get already configured Websites to check if a combination of ip and port is already configured
   Set oThisServer = GetObject( "IIS://localhost/W3SVC" ) 
      
   ' In case IIS cannot be accessed
   If(Err.Number <> 0) Then 
      MsgBox "Unable to retrive websites from IIS - please verify that IIS is running! " & vbCrLf & _ 
             "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
             
      Session.Property("checkNewWebsiteConflict") = "0"
      iReturn = ERROR_INSTALL_FAILURE  
   Else
      	
      ' Circle through defined Websites
      For Each oWebsite In oThisServer 
        
        ' Check if object is a website
        If oWebsite.Class = "IIsWebServer" Then
          	 
		  ' Enumerate the HTTP bindings (ServerBindings) and SSL bindings (SecureBindings)
          bConflict = loopTroughBindings(oWebsite.ServerBindings, iPort, sIP, sHostheader)
          If Not bConflict Then bConflict = loopTroughBindings(oWebsite.SecureBindings, iPort, sIP, sHostheader)

				  ' Exit on Conflict
					If bConflict Then
						MsgBox "Another Website (" & oWebsite.ServerComment &") with the same configuration (IP, Port, Hostheader) is already in use. If you want to use this " & _
						 			 "Website please choose >Existing Website< in the Windows Installer Dialog otherwise please change the " & _
						 			 "provided values!", vbCritical, "Error" 
						 
						Session.Property("checkNewWebsiteConflict") = "0"
						iReturn = ERROR_INSTALL_FAILURE
						Exit For
					End If
					  
        End If 
      Next 
   End If
   
   checkForConflictWithExistingWebsite = iReturn
End Function

Function setIISWebsiteData () 
  Dim sWebsiteComment, oThisServer, oWebsite, oWebsiteRoot, aBindingTypes, aBindingType, aBinding, sBinding, oRegularExpression, oMatches, oMatch
  
  On Error Resume Next
  
  ' Get existing website from WIX
  sWebsiteComment = Session.Property("ExistingWebsite")

  Set oThisServer = GetObject( "IIS://localhost/W3SVC" ) 
      
  If( Err.Number <> 0 ) Then 
     MsgBox "Unable to retrive websites from IIS - please verify that IIS is running" & vbCrLf & _ 
            "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
     setIISWebsiteData = ERROR_INSTALL_FAILURE
     Exit Function 
  Else  
    For Each oWebsite In oThisServer 
        If oWebsite.Class = "IIsWebServer" Then 
            
          If oWebsite.ServerComment & " (ID: " & oWebsite.Name & ")" = sWebsiteComment Then
          	Set oWebsiteRoot = GetObject( "IIS://localhost/W3SVC/" & oWebsite.Name & "/ROOT")
            Session.Property("WEBSITEUNIQUEID") = oWebsite.Name

						' Set website description and path
            Session.Property("WEBSITEDESCRIPTION") = oWebsite.ServerComment
						Session.Property("WEBSITEDIR") = oWebsiteRoot.Path

            ' Define bindings array
						aBindingTypes = Array(oWebsite.ServerBindings, oWebsite.SecureBindings)
						
						' Loop through bindings
						For each aBindingType In aBindingTypes 
					    If IsArray(oWebsite.ServerBindings) Then
					  	   aBinding = oWebsite.ServerBindings(LBound(oWebsite.ServerBindings))
					  	   sBinding = aBinding(LBound(aBinding))
					    Else
					  	   sBinding = oWebsite.ServerBindings
					    End If

              ' Create regular expression
   					  Set oRegularExpression = NEW RegExp
   					  oRegularExpression.Pattern = "([^:]*):([^:]*):(.*)"
   
   					  'sBinding is a string like IP:Port:Host
   
   					  Set oMatches = oRegularExpression.Execute(sBinding)
   					  For Each oMatch In oMatches
          	    Session.Property("WEBSITEIP") = oMatch.SubMatches(0)
				If Session.Property("WEBSITEIP") = "" Then
				  Session.Property("WEBSITEIP") = "*"
				End if
          	    Session.Property("WEBSITEPORT") = oMatch.SubMatches(1)
          	    Session.Property("WEBSITEHOSTHEADER") = oMatch.SubMatches(2)
              Next
            
  					  setIISWebsiteData = ERROR_SUCCESS
          	  Exit Function
          	Next
          End If
        End If 
    Next 
  End If 

End Function

Function loopTroughBindings (oBindingList, iPort, sIP, sHostheader)
  Dim iCount
  
  If IsEmpty(oBindingList) Then
  	 loopTroughBindings = False
  	 Exit Function
  ElseIf IsArray(oBindingList) Then
  	For iCount = LBound(oBindingList) To UBound(oBindingList)
  	  If compareBindingsToValues(oBindingList(iCount), iPort, sIP, sHostheader) Then
  	  	loopTroughBindings = True
  	  	Exit Function
  	  End If
    Next
  Else 
  	If compareBindingsToValues(oBindingList, iPort, sIP, sHostheader) Then
  	  loopTroughBindings = True
  	  Exit Function
  	End If
  End If
  
	loopThroughBindings = False
End Function

Function compareBindingsToValues (sBinding, iPort, sIP, sHostheader)
   Dim oRegularExpression, oMatches, oMatch
   
   ' Create regular expression
   Set oRegularExpression = NEW RegExp
   oRegularExpression.Pattern = "([^:]*):([^:]*):(.*)"
   
   'sBinding is a string looking IP:Port:Host
   
   Set oMatches = oRegularExpression.Execute(sBinding)
   For Each oMatch In oMatches
   				
   		' Matches: 0 -> IP, 1 -> Port, 2 -> Host
   		' If IP, Port and Hostheader equals, exit
   		
			If ((oMatch.SubMatches(0) = sIP) And (CInt(oMatch.SubMatches(1)) = CInt(iPort)) And (oMatch.SubMatches(2) = sHostheader)) Then
				compareBindingsToValues = True
				Exit Function
			End If
					
	  Next
   
    compareBindinsToValues = False
End Function