' 
' This script retrieves all configured ip addresses and
' populates them in an drop down box -> CreateNewWebsiteDlg
'

Function getIPs ()
  Dim oNetworkCards, oNetworkCard, sIP, iCount, iIPCount
  
  On Error Resume Next

  'Creage a NetAdapterConfiguration object
  Set oNetworkCards = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

  'Add default to WIX
  Session.Property("WEBSITEIP") = "AllUnassigned"
  addToWixDropDown 0, "AllUnassigned"

  'Initilize overall iIPCount
  iIPCount = 1

  ' Loop through network cards
  For Each oNetworkCard in oNetworkCards
  
     ' Only interesting if ipenabled is true
     If oNetworkCard.IPEnabled THEN
	   
		  	' Loop trough addresses
        For iCount = LBound(oNetworkCard.IPAddress)  to UBound(oNetworkCard.IPAddress) 
          
          ' Get ip and add it to the drop down box in WIX
	        sIP = oNetworkCard.IPAddress(iCount)
          If sIP <> "" Then
             If iIPCount = 1 Then
               Session.Property("FIRSTIP") = sIP
             End if
          
		 			   addToWixDropDown iIPCount, sIP
             iIPCount = iIPCount + 1
          End If
          
        Next
        
     End If
  Next
  
  getIPs = ERROR_SUCCESS 
End Function

Sub addToWixDropDown(iCount, sName)
   Dim oView, sQuery

	 ' Add IPs to drop down box
	 ' Note: We have two drop down boxes in the template:
	 '       - one for the websites
	 '       - one for the local ip addresses
	 ' Both are empty by design and filled through this or the other getWebsites function...
	 ' to get this working there is a invisible dummy drop down box on the first dialog


   ' Define query
   sQuery = "INSERT INTO `ComboBox` (`Property`, `Order`, `Value`, `Text`) " & _
            "VALUES ('WEBSITEIP', " & iCount & ", '" & sName & "', '" & sName & "') TEMPORARY"

   ' Execute query
   Set oView = Session.Database.OpenView(sQuery)
   oView.Execute

End Sub