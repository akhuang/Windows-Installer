'
' This script verifies path variables for the BrowseDlg 
' E.g.: - getIISDefaultWebsitePath
'				- getIISWebsitePath
'				- checkInstallDirSet
'				- checkVirtualDirectoryName
'				- checkFolderStructure
  
Function getIISDefaultWebsitePath ()
  Dim sDefaultFolder
  
  sDefaultFolder = "C:\Inetpub"

  ' Only set default folder if install dir hasn't been set yet
  If Session.Property("InstallDirSet") = 0 Then
    Session.Property("InstallDir") = sDefaultFolder
  End If

  getIISDefaultWebsitePath = sDefaultFolder
End Function

Function getIISWebsitePath () 
  Dim sWebsiteComment, oThisServer, oWebsite, oWebsiteRoot

  On Error Resume Next
  
  ' Get choosen Website from WIX
  sWebsiteComment = Session.Property("ExistingWebsite")
  
  ' Create IIS Object
  Set oThisServer = GetObject( "IIS://localhost/W3SVC" ) 
      
  If( Err.Number <> 0 ) Then 
     MsgBox "Unable to retrive websites from IIS - please verify that IIS is running" & vbCrLf & _ 
            "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
     getIISWebsitePath = ""
     Exit Function 
  Else
  	
  	' Loop through websites
    For Each oWebsite In oThisServer 
        If oWebsite.Class = "IIsWebServer" Then 
            
          If oWebsite.ServerComment & " (ID: " & oWebsite.Name & ")" = sWebsiteComment Then
          	
          	' Create website root
          	Set oWebsiteRoot = GetObject( "IIS://localhost/W3SVC/" & oWebsite.Name & "/ROOT")
          	
          	' Set website root as path if install dir hasn't been set yet
          	If Session.Property("InstallDirSet") = 0 Then
          	  Session.Property("InstallDir") = oWebsiteRoot.Path
            End If
            
          	getIISWebsitePath = oWebsiteRoot.Path
          	Exit Function
          End If
          
        End If 
    Next 
  End If 

End Function

Function checkInstallDirSet ()
  Dim iResult
  
  ' Default return
  iResult = ERROR_SUCCESS
  
  'If something fails, move on
  On Error Resume Next
  
  
  ' Check if a folder was chosen
  If Session.Property("InstallDirSet") = "0" Then
  	
  	' Not chosen
  	MsgBox "Please choose an installation folder!", vbCritical, "Error" 
  	
  	Session.Property("checkFolderStructure") = "0"
  	checkFolderStructure = ERROR_INSTALL_FAILURE
  	Exit Function
  	
  End If
  
  checkInstallDirSet = ERROR_SUCCESS
End Function

Function checkVirtualDirectoryName() 
    Dim oRegularExpression, iReturn, sVirtualDirName, iMaxCharacters, sRegExPattern
    
    ' Set Defaults
    iMaxCharacters = 32
    
    ' Get sUsername
    sVirtualDirName = Session.Property("VIRTUALDIRECTORYNAME")

    ' Set default result
    Session.Property("checkVirtualDirName") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Set regex params
    sRegExPattern = "^[A-Za-z0-9]{1," & iMaxCharacters & "}$"
    sMsgBoxText = "Your virtual directory name is invalid! Please note that only the characters: A-Z; 0-9 are allowed. The maximum length is " & iMaxCharacters & " characters!"
    
    ' Prepare regex for virtural directory name
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = sRegExPattern
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate virtual directory name
  	If Not oRegularExpression.Test(sVirtualDirName) Then

    	' Port is either not an integer or not in a valid range
    	MsgBox sMsgBoxText, vbCritical, "Error" 
      
      Session.Property("checkVirtualDirName") = "0"
      iReturn = ERROR_INSTALL_FAILURE 
    	
    End If
     
    checkVirtualDirectoryName = iReturn
End Function

Function checkFolderStructure()
  Dim sInstallDir, sRootPath, iAnswer, oFileSystem, oFolder, oFiles, oSubfolders, iResult
  
  ' Default return
  iResult = ERROR_SUCCESS

  On Error Resume Next
  
  ' Get values from WIX
  sInstallDir =  Session.Property("InstallDir")
  
  ' Check if chosen Folder is beneeth original structure
  If Session.Property("TARGETMODE") = "NewWebsite" Then
    sRootPath = getIISDefaultWebsitePath()
  Else 
  	sRootPath = getIISWebsitePath()
  End If
  
  ' Check if folder is beneath root folder and ask for confirmation if necessary
  If InStr(1, UCase(sInstallDir), UCase(sRootPath)) = 0 Then
  	iAnswer = MsgBox("The installation location has been expected beneath " & sRootPath & _
  			   					 " but currently is " & sInstallDir &". Do you want to use the chosen directory?", 36, "Please confirm installation location..." )
  	
  	If iAnswer = 7 Then
  	  Session.Property("checkFolderStructure") = "0"
  	  checkFolderStructure = ERROR_SUCCESS
  	  Exit Function
  	End If
  End If
  
  ' Check if folder exists
	Set oFileSystem = CreateObject("Scripting.FileSystemObject")
  If Not oFileSystem.FolderExists(sInstallDir) Then
  	iAnswer = MsgBox("The chosen installation folder does not exist - " &_
  			   					 "do you want to create it? ", 36, "Create folder..." )
  	
  	If iAnswer = 7 Then
  	  Session.Property("checkFolderStructure") = "0"
  	  checkFolderStructure = ERROR_SUCCESS
  	  Exit Function	
  	End If
	Else
		
		' Check if folder is empty
	  Set oFolder = oFileSystem.GetFolder(sInstallDir)
	  Set oFiles = oFolder.Files
	  Set oSubfolders = oFolder.SubFolders
	  
	  If (oFiles.count + oSubfolders.count) > 0 Then
	  	iAnswer = MsgBox("The chosen Folder contains Files and/or Directories. Please confirm folder usage? ", 36, "Please confirm folder usage..." )
  	
  	  If iAnswer = 7 Then
  	    Session.Property("checkFolderStructure") = "0"
  	    checkFolderStructure = ERROR_SUCCESS
  	    Exit Function
  	  End If
	  End If
	End If

  Session.Property("checkFolderStructure") = "1"
  checkFolderStructure = iResult
End Function