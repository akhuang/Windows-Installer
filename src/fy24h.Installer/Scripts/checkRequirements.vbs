'
' This script handels all checks for the requirement page
' (Administrator check is handled through 'Privlieged' in WIX)
'
' Checks:
' checkIIS 				-> checks for IIS, at least version 6 and above
'checkIIS7AndFastCgi		-> checks if IIS7 is installed with the FastCGI global module (checks if web server role -> CGI feature is installed)
' checkASPNet 		-> checks for an ASP.NET 2.0 installation
' checkAJAX				-> checks for the ASP.NET Ajax Extension
' checkIUSR				-> checks if the IUSR or IUSR_Computername / the 
'									   internet guest account exists and determines it's name
' checkSQLSERVER	-> SoftCheck for SQLSERVER and tries to determine installed instances 
'

Function checkIIs ()
  Dim iisVersion
  
  ' Check for IIS version 6 and above
  
  On Error Resume Next
  	Const msiMessageTypeInfo = &H04000000
	Const msiMessageTypeError = &H01000000
	Const msiMessageTypeWarning = &H02000000
	Set record = Session.Installer.CreateRecord(0)
  
	iisVersion = getRegistryKey("SYSTEM\CurrentControlSet\Services\W3SVC\Parameters", "MajorVersion", True)
  
	If Not (iisVersion = "") Then
		If CInt(iisVersion) >= 6 Then
	  		Session.Property("IISCHECK") = "OK"
			Session.Property("IISVERSION") = cstr(iisVersion)
	  		checkIIs = ERROR_SUCCESS
			record.StringData(0) = "checkRequirements.vbs -> checkIIs() -> IISCHECK=OK -> IISVersion = " & iisVersion
			Session.Message msiMessageTypeInfo, record
	   		Exit Function
	  	End If
	End If
	
	record.StringData(0) = "checkRequirements.vbs -> checkIIs() -> IISCHECK=BAD -> IISVersion could not be determined"
	Session.Message msiMessageTypeError, record
	Session.Property("IISVERSION") = 0
	Session.Property("IISCHECK") = "BAD"
	checkIIs = ERROR_INSTALL_FAILURE
End Function

Function checkIIS7Modules()
	' Check for IIS7 required modules
  
	On Error Resume Next
	Err.Clear

	Const msiMessageTypeInfo = &H04000000
	Const msiMessageTypeError = &H01000000
	Const msiMessageTypeWarning = &H02000000
	Set record = Session.Installer.CreateRecord(0)
  
	iis7moduleshive = "SOFTWARE\Microsoft\InetStp\Components"

	'WMICompatibility, ASPNET, BasicAuthentication, WindowsAuthentication	

	result = true
	Modules = split(Session.Property("IIS7REQUIREDMODULES"), ",")
	for each entry in Modules
		value = getRegistryKey(iis7moduleshive, Trim(entry), True)
		if (Err.Number <> 0) Then 
			value = "0"
			Err.Clear
		end if
		if (CBool(value)<>true) then
			MissingModules = MissingModules & Trim(entry) & ", "
		end if
		result = result and CBool(value)
	next
	
	If (result = true) Then
		Session.Property("IIS7MODULESCHECK") = "OK"
  		checkIIS7Modules = ERROR_SUCCESS
		record.StringData(0) = "checkRequirements.vbs -> checkIIS7Modules() -> IIS7MODULESCHECK=OK"
		Session.Message msiMessageTypeInfo, record
		Session.Property("IIS7MODULESRESULT") =Session.Property("IIS7REQUIREDMODULES")
   		Exit Function
	End If
	
	Session.Property("IIS7MODULESRESULT") = MissingModules
	
	ErrorString = "In order to run this on Vista or Windows Server 2008 you need the following IIS7 modules:" & vbCrLf
	BadModules = split(MissingModules,",")
	for each missingModule in BadModules
		ErrorString = ErrorString & missingModule & vbCrLf
	next
	ErrorString = ErrorString & "Use ServerManager or Control Panel to turn these Windows features on" & vbCrLf & "see: http://technet.microsoft.com/en-us/library/cc753473.aspx"
	record.StringData(0) = ErrorString
	Session.Message msiMessageTypeInfo, record
	Session.Property("IIS7MODULESCHECK") = "BAD"
	checkIIS7Modules = ERROR_INSTALL_FAILURE

	' Explain what is needed
	iAnswer= MsgBox(ErrorString & vbCrLf & "Open URL?", 36, "IIS7 Required Modules Missing...")
	If Not iAnswer = 7 Then
		Set oShell = CreateObject("Shell.Application")
		oShell.ShellExecute Session.Property("IEXPLOREEXE"), "http://technet.microsoft.com/en-us/library/cc753473.aspx", "", "open", 1
	End if
	
	Err.Clear
End Function

Function checkIIS7AndFastCgi ()
    ' Check for IIS7 with FastCgi installed
	On Error Resume Next

	Err.Clear
	Set oWebAdmin = GetObject("winmgmts:root\WebAdministration")   'IIS7 WMI Namespace - requires that IIS managment tools are installed
	Set oFastCgi = oWebAdmin.Get("GlobalModulesSection.Location='',Path='MACHINE/WEBROOT/APPHOST'")
	Const msiMessageTypeInfo = &H04000000
	Const msiMessageTypeError = &H01000000
	Const msiMessageTypeWarning = &H02000000
	Set record = Session.Installer.CreateRecord(0)
	
	oFastCgi.Get "GlobalModules", "Name='FastCgiModule'", oModule
	if (Err.Number = 0) Then 
		Session.Property("IIS7ANDFASTCGI") = "true"
		record.StringData(0) = "checkRequirements.vbs -> checkIIS7AndFastCgi() -> returned IIS7ANDFASTCGI = true"
		Session.Message msiMessageTypeInfo, record
		else 
		Session.Property("IIS7ANDFASTCGI") = "false"
		record.StringData(0) = "checkRequirements.vbs -> checkIIS7AndFastCgi() -> returned IIS7ANDFASTCGI = false (check that IIS managment tools are installed)"
		Session.Message msiMessageTypeInfo, record
	end if
	Err.Clear
End Function


Function checkASPNet ()
  Dim sList, iAnswer, sMessage, oThisServer, sPHPExecutable, aExtension, iExtensionCount, sExtension, iCount, oFileSystem
  
  On Error Resume Next

  ' Set label
  Session.Property("SCRIPTLabel") = "Checking for ASP.NET (Version: " & Session.Property("SCRIPTVERSIONMIN") & " to " & Session.Property("SCRIPTVERSIONMAX") & ")"

  ' Defaults
  iExtensionCount = 0
  Session.Property("SCRIPTCHECK") = "BAD"
  
  ' Create IIS and file system object
  Set oFileSystem = CreateObject("Scripting.FileSystemObject")
  Set oThisServer = GetObject( "IIS://localhost/W3SVC" ) 
      
  ' Error if IIS is nocht accessible
  If( Err.Number <> 0 ) Then 
      MsgBox "Unable to access IIS API - " & _ 
             "please ensure that IIS is running" & vbCrLf & _ 
             "    - Error Details: " & Err.Description & " [Number:" & _ 
             Hex(Err.Number) & "]", vbCritical, "Error" 
             checkASPNet = ERROR_INSTALL_FAILURE 
      Exit Function 
  Else
  
      ' Loop trough IIS extensions
      aExtension = oThisServer.ListExtensionFiles 
      For iCount = LBound(aExtension)  To UBound(aExtension) 
        sExtension = aExtension(iCount)

        ' Check if extension contains the word php
        if InStr(1, sExtension, "aspnet") <> 0 then

           ' Check if file exists
  		     If oFileSystem.FileExists(sExtension) Then
  		     
  		        ' Get version
  		        sVersion = oFileSystem.GetFileVersion(sExtension)

  		        ' Check version
  		        If Session.Property("SCRIPTVERSIONMIN") <= sVersion AND Session.Property("SCRIPTVERSIONMAX")  >= sVersion Then
                 Session.Property("ASPNETISAPIDLL") = sExtension
                 Session.Property("SCRIPTCHECK") = "OK"
                 checkASPNet = ERROR_SUCCESS
                 Exit Function
  	          End if        
           End if
        End if
      Next 
  End If 

  Session.Property("SCRIPTCHECK") = "BAD"
  checkASPNet = ERROR_INSTALL_FAILURE
End Function

Function checkAJAX ()
  Dim ajaxCheck
  
  On Error Resume Next
  
  ' Checks for ASP.NET AJAX Extension
  ajaxCheck = getRegistryKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{082BDF7B-4810-4599-BF0D-E3AC44EC8524}", "DisplayName", False)

  If Not (ajaxCheck = "") Then
  	 Session.Property("AJAXCheck") = "OK"
  	 checkAJAX = ERROR_SUCCESS
  	 Exit Function
  End If

  Session.Property("AJAXCheck") = "BAD"
  checkAJAX = ERROR_INSTALL_FAILURE
End Function

Function checkIUSR ()
    Dim oThisServer, oWebsite , iCount, sQuery, oView
    iCount = 0
    
    ' IIS internet guest account
    On Error Resume Next 
    Set oThisServer = GetObject( "IIS://localhost/W3SVC" ) 
      
    ' Set default
    Session.Property("IUSRCheck") = "BAD"
    checkIUSR = ERROR_INSTALL_FAILURE

    If( Err.Number <> 0 ) Then 
        MsgBox "Unable to retrieve IIS Websites - Please verify that the IIS is running! " & vbCrLf & _ 
               "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        GetWebSites = ERROR_INSTALL_FAILURE 
        Exit Function 
    Else	
    	  Session.Property("IUSRCheck") = "OK"
        Session.Property("INTERNETGUESTACCOUNT") = oThisServer.AnonymousUserName
        If Session.Property("INTERNETGUESTACCOUNT") = "IUSR" Then
          Session.Property("INTERNETGUESTDOMAIN") = "NT AUTHORITY"
        Else
          Session.Property("INTERNETGUESTDOMAIN") = Session.Property("ComputerName")
        End if
        checkIUSR = ERROR_SUCCESS
    End If
End Function

Function checkSQLSERVER () 'for SQL Server only
  Dim oWMI, sFoundSQLServer, oSQLSERVER
  On Error Resume Next

  ' This functions tries to find all local databases

  sFoundSQLServer = ""
  Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")
  For each oSQLSERVER in oWMI.ExecQuery("SELECT Name FROM Win32_Service WHERE Caption like '%SQL Server%'")
	sFoundSQLServer = sFoundSQLServer & oSQLSERVER.Name & ", "
  Next
  
  If Not (sFoundSQLServer = "") Then
  	Session.Property("SQLSERVERCHECK") = "RedX"
  	
  	' Some cosmetics and save all found sql servers / instances
  	sFoundSQLServer = Mid(sFoundSQLServer, 1, (Len(sFoundSQLServer) - 2))
  	sFoundSQLServer = Replace(sFoundSQLServer, "#", " ")
  	sFoundSQLServer = Replace(sFoundSQLServer, "$", " ")
  	Session.Property("SQLSERVERFOUND") = "(" & sFoundSQLServer & ")"
  	checkADMIN = ERROR_SUCCESS
  	Exit Function
  End If
  
'  Session.Property("SQLSERVERCheck") = "BAD"
  checkADMIN = ERROR_INSTALL_FAILURE
End Function

Function checkMySQLODBC ()
  Dim iODBCVersion35121, iODBCVersion515, oShell, wiInstaller
  
  On Error Resume Next
  
  'we ask windows installer database whether a MySQL Connector/ODBC is installed
  Set wiInstaller = CreateObject("WindowsInstaller.Installer")
  for each product in wiInstaller.Products
	productName = wiInstaller.ProductInfo(product,"ProductName")
	if (InStr(1, productName,"MySQL Connector/ODBC") <> 0) then
		select case productName
			case "MySQL Connector/ODBC 3.51"
				Session.Property("MYSQL_ODBC_DRIVER") = "MySQL ODBC 3.51 Driver"
			case "MySQL Connector/ODBC 5.1"
				Session.Property("MYSQL_ODBC_DRIVER") = "MySQL ODBC 5.1 Driver"
			case else 'best guess
				VersionMajor = wiInstaller.ProductInfo(product,"VersionMajor")
				VersionMinor = wiInstaller.ProductInfo(product,"VersionMinor")
				Session.Property("MYSQL_ODBC_DRIVER") = "MySQL ODBC " & VersionMajor & "." & VersionMinor & " Driver"
		end select
		Session.Property("SQLSERVERCheck") = "OK"
		checkMySQLODBC = ERROR_SUCCESS
		Exit Function	
  	end if
  next
  set wiInstaller = nothing
  
  ' Explain why odbc drive is need
  iAnswer = MsgBox("In order to enable this installer to setup a MySQL Database you need to install the MySQL ODBC Connector. " &_ 
                   "These drivers can be obtained from: " & vbCrLf & vbCrLf &_
                   "|- http://dev.mysql.com/downloads/connector/odbc" & vbCrLf &_
                   "|-- Windows downloads" & vbCrLf &_
                   "|--- Windows MSI Installer (x86)" & vbCrLf & vbCrLf &_
                   "Do you want to open this URL?", 36, "MySQL ODBC Connector missing...")
      
  If Not iAnswer = 7 Then
     Set oShell = CreateObject("Shell.Application")
     oShell.ShellExecute Session.Property("IEXPLOREEXE"), "http://dev.mysql.com/downloads/connector/odbc", "", "open", 1
  End if

  Session.Property("SQLSERVERCheck") = "BAD"
  checkMySQLODBC = ERROR_INSTALL_FAILURE
End Function

Function getRegistryKey (sKeyPath, sValueName, bDword) 
  Dim sComputer, oReg, sValue
  
  ' A helper function to retrieve registry values
  
	Const HKEY_LOCAL_MACHINE = &H80000002

	On Error Resume Next
	Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
 
  ' Act on specified value type
  If bDword Then
	  oReg.GetDWORDValue HKEY_LOCAL_MACHINE,sKeyPath,sValueName,sValue
	Else
	  oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,sKeyPath, sValueName, sValue
	End If
	
	getRegistryKey = sValue 
End Function