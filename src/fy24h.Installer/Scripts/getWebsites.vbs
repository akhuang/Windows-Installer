  ' 
  ' This script retrieves all configured IIS websites and
  ' populates them in an drop down box -> UseExistingWebsiteDlg
  '
  
  Function getWebsites( ) 
    Dim oThisServer, oWebsite , iCount, sQuery, oView
    iCount = 0
    
    ' Get websites currently configured in IIS and adds them to a drop down box

    On Error Resume Next 
    Set oThisServer = GetObject( "IIS://localhost/W3SVC" ) 
      
    If( Err.Number <> 0 ) Then 
        MsgBox "Unable to retrieve IIS Websites - Please verify that the IIS is running! " & vbCrLf & _ 
               "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        GetWebSites = ERROR_INSTALL_FAILURE 
        Exit Function 
    Else	 
      
      ' Loop trough websites
      For Each oWebsite In oThisServer 
        If oWebsite.Class = "IIsWebServer" Then 
        
            ' Set first found website
            If iCount = 0 Then 
                Session.Property("ExistingWebsite") = oWebsite.ServerComment & " (ID: " & oWebsite.Name & ")"
            End If 

						' Add websites to drop down box
						' Note: We have two drop down boxes in the template:
						'       - one for the websites
						'       - one for the local ip addresses
						' Both are empty be design and filled trough this or the other getIPs function...
						' to get this working there is a invisible dummy drop down box on the first dialog
						
            sQuery = "INSERT INTO `ComboBox` (`Property`, `Order`, `Value`, `Text`) " &_
                     "VALUES ('ExistingWebsite', " & iCount & ", '" & oWebsite.ServerComment & " (ID: " & oWebsite.Name & ")" & "', '" & oWebsite.ServerComment & " (ID: " & oWebsite.Name & ")" & "') TEMPORARY"
            
            wscript.echo sQuery
            Set oView = Session.Database.OpenView(sQuery)
            oView.Execute
            Set oView = Nothing
            
            iCount = iCount + 1 
        End If 
      Next 
      
      getWebsites = ERROR_SUCCESS 
    End If 

    ' Clean up 
    Set oThisServer = Nothing 
    Set oWebSite = Nothing 
End Function