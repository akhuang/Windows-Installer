'
' This script handels all database checks and interaction
' before the installation process takes place. that means
' it checks if a connect is possible, if the user has the:
' - right to create databases
' - right to create new users
'
' And if the database only supports windows authentification
' or sql authentification as well (is a must).
'
' This script also contains functions to propose a username,
' a database and a sql user password. There are also checks
' if the chosen values for username, database or password
' are valid.
'

' Global database object
Dim oDB

Function getConnectionString()
   Dim sDBHost, sDBPort, sDBUsername, sDBPassword, sSecurity, sConnectionString, sDatabaseEngine, sPortAddon

    Const msiMessageTypeInfo = &H04000000
	Const msiMessageTypeError = &H01000000
	Const msiMessageTypeWarning = &H02000000
	Set record = Session.Installer.CreateRecord(0)
	record.StringData(0) = "checkDatabase.vbs -> getConnectionString() -> begin " & iisVersion
	Session.Message msiMessageTypeInfo, record
   
   
   ' Get Values from Windows Installer
   sDBHost         = Session.Property("DBHOST")
   sDBPort         = Session.Property("DBPORT")
   sDBUsername     = Session.Property("SQLADMINUSERNAME")
   sDBPassword     = Session.Property("SQLADMINPASSWORD")
   sSecurity       = Session.Property("DBSECURITYMETHOD")
   sDatabaseEngine = Session.Property("DATABASEENGINE")
   sDataBaseName   = Session.Property("SQLDATABASE")
   
   'sPortAddon = ":" & sDBPort
   'If sDBHost = "localhost" Then
   '  sPortAddon = ""
   'End if
   sConnectionString = "Data Source=" & sDBHost & ";Initial Category=" & sDataBaseName & ";"
   If sDatabaseEngine = "MSSQL" Then
     If sSecurity = "integrated" Then
       ' integrated security
       sConnectionString = sConnectionString & ";Integrated Security=SSPI;Persist Security Info=False;"
     Else
       ' user auth / sql authentification
       sConnectionString = sConnectionString & ";User ID=" & sDBUsername & ";Password=" & sDBPassword 
     End If
   Elseif sDatabaseEngine = "MYSQL" Then
     'sConnectionString = "driver=MySQL ODBC 3.51 Driver;server=" & sDBHost & sPortAddon & ";uid=" & sDBUsername & ";pwd=" & sDBPassword & ";database=mysql;option=NUM"
	sConnectionString = "driver=" & Session.Property("MYSQL_ODBC_DRIVER") & ";Server=" & sDBHost & ";Port=" & sDBPort  & ";uid=" & sDBUsername & ";pwd=" & sDBPassword & ";database=mysql"
	 
   Else
     sConnectionString = null
   End if
   
   ' Return connection string
   getConnectionString = sConnectionString
End Function

Function createDatabaseObject()
    Dim sConnectionString

    On Error Resume Next
	Const msiMessageTypeInfo = &H04000000
	Const msiMessageTypeError = &H01000000
	Const msiMessageTypeWarning = &H02000000
	Set record = Session.Installer.CreateRecord(0)
	record.StringData(0) = "checkDatabase.vbs -> createDatabaseObject() -> begin " & iisVersion
	Session.Message msiMessageTypeInfo, record
		
    ' Get connection string
		sConnectionString = getConnectionString()
		
    ' Crate ADODB Object
    Set oDB = CreateObject("ADODB.Connection")
      
    ' Check if object could be created
    If( Err.Number <> 0 ) Then 
        MsgBox "Unable to create ADODB Object - please verify that ASP.NET 2.0 is installed!" & vbCrLf & _ 
               "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        createDatabaseObject = False
        Exit Function 
    Else 

    	' Open connection
    	oDB.Open(sConnectionString)
    	
    	' Check if connection worked
    	If( Err.Number <> 0 ) Then 
    	  MsgBox "Unable to connect to Database - please provide accurate login data!" & vbCrLf & _ 
               "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        createDatabaseObject = False 
        Exit Function 
      End If
    End If
    
    createDatabaseObject = True
End Function

Function proposeDatabaseName ()
  Dim sDatabase, sDBPrefix, sSQL, sColumn, oResultSet, aDatabases(), iCount, iNumber, bFound, sDatabaseEngine
 
  On Error Resume Next
 
  ' Get set database and preconfigured database prefix property from WIX
  sDatabase       = Session.Property("SQLDATABASE")
  sDBPrefix       = Session.Property("DBPrefix")
  sDatabaseEngine = Session.Property("DATABASEENGINE")
  
  ' Only act if the user hasn't changed the database name yet -> sDatabase = 0
  If sDatabase = "" And createDatabaseObject() Then
     If sDatabaseEngine = "MSSQL" Then
  	   sSQL =  "sp_databases"
  	   sColumn = "DATABASE_NAME"
  	 Else
  	   sSQL = "show databases;"
  	   sColumn = "Database"
  	 End if
  	 
     iCount = 0
    
     ' Get all existing databases
     ReDim aDatabases(0) 
     Set oResultSet = oDB.Execute(sSQL)
     
     If( Err.Number <> 0 ) Then 
    	  MsgBox "Unable to get all existing databases - Please choose another User for Database setup! [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        proposeDatabaseName = False 
        Exit Function 
     End If
     
     ' Add all existing databases to aDatabases array
     Do While Not oResultSet.EOF
     	  ReDim preserve aDatabases((iCount + 1))
        aDatabases(iCount) = oResultSet(sColumn)
        iCount = iCount + 1
        oResultSet.MoveNext
     Loop
  	
  	 ' Find the next free database according to preconfigured prefix
  	 ' E.g.: MyDatabase_1, MyDatabase_2... and so on
  	 
  	 iNumber = 1
  	 While sDatabase = ""
 
  	 	 bFound = False
  	 	 For iCount = LBound(aDatabases) To UBound(aDatabases)
  	 	   If UCase(aDatabases(iCount)) = UCase(sDBPrefix & iNumber) Then
  	 	   	  bFound = True
  	 	   End If
  	 	 Next
  	 	 
  	 	 If Not bFound Then
  	 	 	 sDatabase = sDBPrefix & iNumber 
  	 	 End If
  	 	 
  	 	 iNumber = iNumber + 1
  	 Wend
  	
  	 Session.Property("SQLDATABASE") = sDatabase
  End If
End Function

Function checkDatabaseName ()
    Dim oRegularExpression, iReturn, sDatabaseName, iMinLength, iMaxLength, sDBExists, sSQL, oResultSet, bFound, sDatabaseEngine

    On Error Resume Next
    
    ' Set defaults
    iMinLength = 4
    iMaxLength = 16
    bFound = false
    
    ' Get sDatabaseName
    sDatabaseName = Session.Property("SQLDATABASE")
    sDatabaseEngine = Session.Property("DATABASEENGINE")
    
    ' Set default result
    Session.Property("checkDatabaseName") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Prepare regex validation for integer value
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = "^[a-z0-9A-Z\_]{" & iMinLength & "," & iMaxLength & "}$"
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate if the database name machtes naming policy
  	If Not oRegularExpression.Test(sDatabaseName) Then

    	' Port is either not an integer or not in a valid range
    	MsgBox "The specified database name is invalid. The name has to be at least " & iMinLength & " Characters (Max: " & iMaxLength & ") allowed are A-Z, a-z, underscore and the Numbers from 0 to 9!", vbCritical, "Error" 
      
      Session.Property("checkDatabaseName") = "0"
      iReturn = ERROR_INSTALL_FAILURE 
    	Exit Function
    End If
    
    ' Validate against database... (sql injection impossible because of regex)
    If createDatabaseObject() Then
      sDBExists = "no"
      sSQL = "IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = '" & sDatabaseName & "') " &_
             "   SELECT 'yes' as dbExists  " &_
             "ELSE " &_
             "   SELECT 'no' as usedbExists "

     If sDatabaseEngine = "MSSQL" Then
  	   sSQL = "SELECT name FROM master.dbo.sysdatabases WHERE name = '" & sDatabaseName & "'"
  	 Else
  	   sSQL = "SHOW databases LIKE '" & sDatabaseName & "';"
  	 End if

      Set oResultSet = oDB.Execute(sSQL)
      
     If( Err.Number <> 0 ) Then 
    	  MsgBox "Unable to check if database already exists - Please choose another User for Database setup! [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        checkDatabaseName = False 
        Exit Function 
     End If
      
      Do While Not oResultSet.EOF
        bFound = true
        oResultSet.MoveNext
      Loop

      If bFound Then
    	   MsgBox "The specified database already exists please choose another name!", vbCritical, "Error" 
         checkDatabaseName = ERROR_INSTALL_FAILURE
         Session.Property("checkDatabaseName") = "0"
         Exit Function 
      End If    
   End If
   
   checkDatabaseName = iReturn
End Function

Function proposeUsername ()
  Dim sUsername, sUSRPrefix, sSQL, oResultSet, aUsers(), iCount, iNumber, bFound, sDatabaseEngine
 
  On Error Resume Next
 
  ' Get username and user prefix from WIX
  sUsername = Session.Property("SQLUSERUSERNAME")
  sUSRPrefix = Session.Property("USRPrefix")
  sDatabaseEngine = Session.Property("DATABASEENGINE")
  
  ' Propose username only if the user hasn't set it yet
  If sUsername = "" And createDatabaseObject() Then
  	
  	 ' Get all database users
     If sDatabaseEngine = "MSSQL" Then
  	   sSQL =  "SELECT name as SQLUsername FROM sys.sql_logins"
  	 Else
  	   sSQL =  "SELECT user AS SQLUsername FROM mysql.user GROUP BY user;"
  	 End if
     iCount = 0
    
     ReDim aUsers(0) 
     Set oResultSet = oDB.Execute(sSQL)
     
     If( Err.Number <> 0 ) Then 
    	  MsgBox "Unable to get existing sql logins (sys.sql_logins) - Please choose another User for Database setup! [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        proposeUsername = False 
        Exit Function 
     End If
     
     ' Loop through sql result
     Do While Not oResultSet.EOF
     	  ReDim preserve aUsers((iCount + 1))
        aUsers(iCount) = oResultSet("SQLUsername")
        iCount = iCount + 1
        oResultSet.MoveNext
     Loop
  	
  	 ' Find next free user name
  	 iNumber = 1
  	 While sUsername = ""
 
  	 	 bFound = False
  	 	 For iCount = LBound(aUsers) To UBound(aUsers)
  	 	   If UCase(aUsers(iCount)) = UCase(sUSRPrefix & iNumber) Then
  	 	   	  bFound = True
  	 	   End If
  	 	 Next
  	 	 
  	 	 If Not bFound Then
  	 	 	 sUsername = sUSRPrefix & iNumber 
  	 	 End If
  	 	 
  	 	 iNumber = iNumber + 1
  	 Wend

  	 Session.Property("SQLUSERUSERNAME") = sUsername
  End If
End Function

Function checkSQLPassword() 
    Dim oRegularExpression, iReturn, sPassword, iMinPasswordLength, iMaxPasswordLength
    
    ' Set defaults
    iMinPasswordLength = 8
    iMaxPasswordLength = 32
    
    ' Get sPassword
    sPassword = Session.Property("SQLUSERPASSWORD")
    
    ' Set default result
    Session.Property("checkSQLPassword") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Prepare regex validation for integer value
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = "^[a-z0-9A-Z@]{" & iMinPasswordLength & "," & iMaxPasswordLength & "}$"
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate if iPort is integer and between min and max values
  	If Not oRegularExpression.Test(sPassword) Then

    	' Port is either not an integer or not in a valid range
    	MsgBox "The specified password is invalid. The Password has to be at least " & iMinPasswordLength & " Characters (Max: " & iMaxPasswordLength & ") allowed are A-Z, a-z, @ and the Numbers from 0 to 9!", vbCritical, "Error" 
      
      Session.Property("checkSQLPassword") = "0"
      iReturn = ERROR_INSTALL_FAILURE 
    	
    End If
     
    checkSQLPassword = iReturn
End Function

Function checkUsernameName ()
    Dim oRegularExpression, iReturn, sUsername, iMinLength, iMaxLength, sUSRExists, sSQL, oResultSet, sDatabaseEngine, bFound
    
    On Error Resume Next
    
    ' Set defaults
    iMinLength = 4
    iMaxLength = 16
    bFound = false
    
    ' Get sDatabaseName
    sUsername = Session.Property("SQLUSERUSERNAME")
    sDatabaseEngine = Session.Property("DATABASEENGINE")
    
    ' Set default result
    Session.Property("checkSQLUsernameName") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Prepare regex validation for integer value
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = "^[a-z0-9A-Z\_]{" & iMinLength & "," & iMaxLength & "}$"
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate if the database name machtes naming policy
  	If Not oRegularExpression.Test(sUsername) Then

    	' Port is either not an integer or not in a valid range
    	MsgBox "The specified sql username is invalid. The name has to be at least " & iMinLength & " Characters (Max: " & iMaxLength & ") allowed are A-Z, a-z, underscore and the Numbers from 0 to 9!", vbCritical, "Error" 
      
      Session.Property("checkSQLUsernameName") = "0"
      checkUsernameName = ERROR_INSTALL_FAILURE 
    	Exit Function
    End If
    
    ' Validate against database... (sql injection impossible because of regex)
    If createDatabaseObject() Then
      sUSRExists = "no"
      
     If sDatabaseEngine = "MSSQL" Then
  	   sSQL =  "SELECT name FROM sys.sql_logins WHERE name = '" & sUsername & "'"
  	 Else
  	   sSQL =  "SELECT user FROM mysql.user WHERE user = '" & sUsername & "'"
  	 End if

     Set oResultSet = oDB.Execute(sSQL)
      
     If( Err.Number <> 0 ) Then 
    	  MsgBox "Unable to get existing logins - Please choose another user for database setup! [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        proposeUsername = False 
        Exit Function 
     End If
      
      ' Loop rough sql resultset
      Do While Not oResultSet.EOF
        bFound = true
        oResultSet.MoveNext
      Loop

      If bFound Then
    	   MsgBox "The specified sql username already exists please choose another one!", vbCritical, "Error" 
         checkUsernameName = ERROR_INSTALL_FAILURE
         Session.Property("checkSQLUsernameName") = "0"
         Exit Function 
      End If    
   End If
   
   checkUsernameName = iReturn
End Function

Function proposePassword()
  Dim aChars(52), sPassword, iPasswordLength
  Dim aSpecialChars(11)
  On Error Resume Next
  
  ' Propose password...
  ' I guess that once we have stored the sql user password in the web.config file
  ' nobody will use it again... so we'll make it save
  
  ' Password length
  iPasswordLength = "16"
  
  ' Only set if the user hasn't chosen a password yet
  If Session.Property("SQLUSERPASSWORD") = "" Then
  
  	
  	 ' Characters and digits to choose from
     aChars(0) = "A"
     aChars(1) = "B"
     aChars(2) = "C"
     aChars(3) = "D"
     aChars(4) = "E"
     aChars(5) = "F"
     aChars(6) = "G"
     aChars(7) = "H"
     aChars(8) = "I"
     aChars(9) = "J"
     aChars(10) = "K"
     aChars(11) = "L"
     aChars(12) = "M"
     aChars(13) = "N"
     aChars(14) = "O"
     aChars(15) = "P"
     aChars(16) = "Q"
     aChars(17) = "R"
     aChars(18) = "S"
     aChars(19) = "T"
     aChars(20) = "U"
     aChars(21) = "V"
     aChars(22) = "W"
     aChars(23) = "X"
     aChars(24) = "Y"
     aChars(25) = "Z"
     aChars(26) = "a"
     aChars(27) = "b"
     aChars(28) = "c"
     aChars(29) = "d"
     aChars(30) = "e"
     aChars(31) = "f"
     aChars(32) = "g"
     aChars(33) = "h"
     aChars(34) = "i"
     aChars(35) = "j"
     aChars(36) = "k"
     aChars(37) = "l"
     aChars(38) = "m"
     aChars(39) = "n"
     aChars(40) = "o"
     aChars(41) = "p"
     aChars(42) = "q"
     aChars(43) = "r"
     aChars(44) = "s"
     aChars(45) = "t"
     aChars(46) = "u"
     aChars(47) = "v"
     aChars(48) = "w"
     aChars(49) = "x"
     aChars(50) = "y"
     aChars(51) = "z"

	 aSpecialChars(0) = "0"
	 aSpecialChars(1) = "1"
     aSpecialChars(2) = "2"
     aSpecialChars(3) = "3"
     aSpecialChars(4) = "4"
     aSpecialChars(5) = "5"
     aSpecialChars(6) = "6"
     aSpecialChars(7) = "7"
     aSpecialChars(8) = "8"
	 aSpecialChars(9) = "9"
	 aSpecialChars(10) = "@"

     ' Generate a nice and secure password
     sPassword = ""
     Randomize
     Do Until Len(sPassword) = Int(iPasswordLength)
		sPassword = sPassword & aChars(int(rnd()*52))
     Loop
	 
	 'Put some special chars into the password string at a random position
	 counter=0
	 Randomize
	 Do Until counter=6
		position=int(rnd()*Len(sPassword))
		sPassword=Left(sPassword,position) + aSpecialChars(int(rnd()*11)) + Right(sPassword,Len(sPassword)-(position+1))
		counter=counter+1
	 Loop

     Session.Property("SQLUSERPASSWORD") = sPassword
     
  End If
  proposePassword = ERROR_SUCCESS
End Function


Function checkDatabaseLogin () 
    Dim sSQL, oResultSet, sDatabaseEngine

    On Error Resume Next
  	Const msiMessageTypeInfo = &H04000000
	Const msiMessageTypeError = &H01000000
	Const msiMessageTypeWarning = &H02000000
	Set record = Session.Installer.CreateRecord(0)
	record.StringData(0) = "checkDatabase.vbs -> checkDatabaseLogin() -> begin"
	Session.Message msiMessageTypeInfo, record
	
    ' Check if a dabase login is possible and if the provided
    ' username is privileged enough to do a few things like:
    ' - create databases
    ' - create users
    ' - and if the database authentification mode is mixed

    ' Set Default
    Session.Property("checkDatabaseConnection") = "0"
    sDatabaseEngine = Session.Property("DATABASEENGINE")
    
    If createDatabaseObject() Then
    	If sDatabaseEngine = "MSSQL" Then
    	  ' Check for create database, login privleges
    	  Dim sCreateDatabaseAndLogin
    	  sCreateDatabaseAndLogin = "no"
		  sSQL =  "if (HAS_PERMS_BY_NAME(NULL, NULL, 'CREATE ANY DATABASE') =1) AND  (HAS_PERMS_BY_NAME(NULL,NULL,'ALTER ANY LOGIN')=1) Select 'yes' as createDatabase ELSE   SELECT 'no'  as createDatabase"
    	
    	  Set oResultSet = oDB.Execute(sSQL)
    	  Do While Not oResultSet.EOF
          sCreateDatabaseAndLogin = oResultSet("createDatabase")
          oResultSet.MoveNext
        Loop

    	  If sCreateDatabaseAndLogin = "no" Then
    	    MsgBox "The provided user account has insuffient rights - it does not have CREATE ANY DATABASE and ALTER ANY LOGIN permissions!", vbCritical, "Error" 
          checkDatabaseLogin = ERROR_INSTALL_FAILURE
          Exit Function 
    	  End If
    	
    	 ' Check for supported authentication methods
    	  Dim sSecurityMode
    	  sSecurityMode = "integrated"
    	  sSQL = "IF serverproperty('IsIntegratedSecurityOnly') = 1 " &_
    	         "   SELECT 'integrated' as securityMode " &_
    	         "ELSE " &_
    	         "   SELECT 'user' as securityMode"

    	  Set oResultSet = oDB.Execute(sSQL)
    	  Do While Not oResultSet.EOF
          sSecurityMode = oResultSet("securityMode")
          oResultSet.MoveNext
        Loop

    	  If sSecurityMode = "integrated" Then
    	    MsgBox "This database server only supports integrated windows authentification - Please enable SQL Logins (See http://www.microsoft.com/technet/prodtechnol/sql/2005/mgsqlexpwssmse.mspx for further instructions)!", vbCritical, "Error" 
          checkDatabaseLogin = ERROR_INSTALL_FAILURE
          Exit Function 
    	  End If
		  
		  'Check for full-text search if  checkForFullTextSearch is 1
		  If Session.Property("checkForFullTextSearch")=1 Then
			Dim sFulltextSearch
	    	sFulltextSearch = 0
	    	sSQL =  "SELECT fulltextserviceproperty('IsFulltextInstalled') As fullTextSearch"
	    	
	    	Set oResultSet = oDB.Execute(sSQL)
	    	Do While Not oResultSet.EOF
	          sFulltextSearch = oResultSet("fullTextSearch")
	          oResultSet.MoveNext
	        Loop

			If sFulltextSearch = 0 Then
				MsgBox "Full-text search is not installed for this SQL Server instance.", vbCritical, "Error" 
				checkDatabaseLogin = ERROR_INSTALL_FAILURE
				Exit Function 
			End If    	
		  End If
		  
     Else
       Dim sUser, sGrants
       record.StringData(0) = "checkDatabase.vbs -> checkDatabaseLogin() -> show grants for user"
	   Session.Message msiMessageTypeInfo, record
       ' Get user/host name
       sSQL = "SELECT user();"
       Set oResultSet = oDB.Execute(sSQL)
    	 Do While Not oResultSet.EOF
          sUser = oResultSet("user()")
          oResultSet.MoveNext
       Loop
       
       ' Get priviliges
       sSQL = "show grants;"
         Set oResultSet = oDB.Execute(sSQL)
    	 Do While Not oResultSet.EOF
          sGrants = oResultSet("Grants for " & sUser)
          oResultSet.MoveNext
       Loop
       
       If instr(1, sGrants, "GRANT ALL PRIVILEGES ON *.*") = 0 Then
          MsgBox "Please specify a user which have been granted all privileges (e.g. root)!", vbCritical, "Error" 
          checkDatabaseLogin = ERROR_INSTALL_FAILURE
          Exit Function
       End if
       
     End if
     ' Return values
     Session.Property("checkDatabaseConnection") = "1"
     checkDatabaseLogin = ERROR_SUCCESS
   End If
End Function

Function checkMYSQL40()
    On Error Resume Next
	Err.Clear
	
    ' Get connection string
		sConnectionString = getConnectionString()
		
    ' Crate ADODB Object
    Set oDB = CreateObject("ADODB.Connection")
      
    ' Check if object could be created
    If( Err.Number <> 0 ) Then 
        MsgBox "Unable to create ADODB Object" & vbCrLf & _ 
               "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
        createDatabaseObject = False
		writeWindowsInstallerLogEntry "checkDatabase.vbs -> checkMYSQL40 -> ADODB.Connection failed -> Err.Description="&Err.Description,0
        Exit Function 
    Else 
    	' Open connection
    	oDB.Open(sConnectionString)
    	
    	' Check if connection worked
    	If( Err.Number <> 0 ) Then 
    	  MsgBox "Unable to connect to Database - please provide accurate login data!" & vbCrLf & _ 
               "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
			createDatabaseObject = False 
			writeWindowsInstallerLogEntry "checkDatabase.vbs -> checkMYSQL40 -> Open DB failed -> Err.Description="&Err.Description,0
			Exit Function 
		End If
		
		Set oResultSet = oDB.Execute("select @@sql_mode;")
		for each field in oResultSet.fields
			if (InStr(1, oResultSet(field.name), "MYSQL40") <> 0) then
				isMYSQL40Mode = true
				exit for
			end if
			oResultSet.MoveNext
		next 
		
		'cleanup
		set oDB=Nothing
		set oResultSet=Nothing
    End If
    
	'for logging
	If( Err.Number <> 0 ) Then 
    	  MsgBox "An error occurred during executing mysql(select @@sql_mode) query" & vbCrLf & _ 
               "Error Details: " & Err.Description & " [Number:" & Hex(Err.Number) & "]", vbCritical, "Error" 
			writeWindowsInstallerLogEntry "checkDatabase.vbs -> checkMYSQL40 -> error during executing 'select @@sql_mode' -> Err.Description="&Err.Description,0
			Session.Property("checkDatabaseConnection") = "0"
			checkMYSQL40 = ERROR_INSTALL_FAILURE 
			Exit Function 
	Elseif (isMYSQL40Mode = true) then
		writeWindowsInstallerLogEntry "checkDatabase.vbs -> checkMYSQL40 -> sql-mode = MYSQL40",0
		Session.Property("checkDatabaseConnection") = "1"
		checkMYSQL40 = ERROR_SUCCESS
	else
		writeWindowsInstallerLogEntry "checkDatabase.vbs -> checkMYSQL40 -> sql-mode <> MYSQL40",0
	    MsgBox "This web application requires a special MySql Setting:" & vbCrLF & "sql-mode=" & chr(34) & "MYSQL40" & chr(34) & vbCrLF & "Please alter the my.ini file in your MySQL installation folder. Restart the MySQL Service. Continue this setup", vbInformation, "MySQL sql-mode is not MYSQL40" 
		checkMYSQL40 = ERROR_INSTALL_FAILURE 
		Session.Property("checkDatabaseConnection") = "0"
	End If
    
End Function

function writeWindowsInstallerLogEntry(info,infotype)
	
	On Error Resume Next
	Err.Clear
	'infotype can be:
	' 0 for info
	' 1 for warning
	' 2 for error - gives a user prompt - you maybe do not want this.

  	Const msiMessageTypeInfo = &H04000000
	Const msiMessageTypeWarning = &H02000000
	Const msiMessageTypeError = &H01000000

	Set record = Session.Installer.CreateRecord(0)
  	record.StringData(0) = info
	
	select case infotype
		case 0	'information
			Session.Message msiMessageTypeInfo, record
		case 1	'warning
			Session.Message msiMessageTypeWarning, record
		case 2	'error
			Session.Message msiMessageTypeError, record
		case else
			Session.Message msiMessageTypeInfo, record
	end select
	
end function