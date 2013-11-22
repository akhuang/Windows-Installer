'
' This script handles all checks for the first user setup dialog.
' This means: - checkUsername
' 						- checkPassword
'							- checkEmail
'

Function checkUsername() 
    Dim oRegularExpression, iReturn, sUsername, bFirstUsernameWhitespaces, sMsgBoxText, iMaxCharacters, sRegExPattern
    
    ' Set Defaults
    iMaxCharacters = 24
    
    ' Get sUsername from WIX
    sUsername = Session.Property("USERUSERNAME")
    
    ' Check for options... if a user name can contain whitespaces for example
    If Session.Property("FirstUsernameWhitespaces") = "1" Then
    	 sRegExPattern = "^[A-Za-z0-9\s]{1," & iMaxCharacters & "}$"
       bFirstUsernameWhitespaces = True
       sMsgBoxText = "Your Username is invalid! Please note that only the characters: A-Z; 0-9 and whitespaces are allowed. The maximum length is " & iMaxCharacters & " characters!"
    Else
    	 sRegExPattern = "^[A-Za-z0-9]{1," & iMaxCharacters & "}$"
       bFirstUsernameWhitespaces = False
    	 sMsgBoxText = "Your Username is invalid! Please note that only the characters: A-Z; 0-9 are allowed. The maximum length is " & iMaxCharacters & " characters!"
    End If

    ' Set default result
    Session.Property("checkFirstUserUsername") = "1"
    iResult = ERROR_SUCCESS 
    
    
    ' Prepare regex validation for integer value
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = sRegExPattern
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate if iPort is integer and between min and max values
  	If Not oRegularExpression.Test(sUsername) Then

    	' Port is either not an integer or not in a valid range
    	MsgBox sMsgBoxText, vbCritical, "Error" 
      
      Session.Property("checkFirstUserUsername") = "0"
      iReturn = ERROR_INSTALL_FAILURE 
    	
    End If
     
    checkHostname = iReturn
End Function

Function checkPassword() 
    Dim oRegularExpression, iReturn, sPassword, iMinPasswordLength, iMaxPasswordLength
    
    ' Set defaults
    iMinPasswordLength = 8
    iMaxPasswordLength = 32
    
    ' Get sPassword from WIX
    sPassword = Session.Property("USERPASSWORD")
    
    ' Set default result
    Session.Property("checkFirstUserPassword") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Prepare regex validation for integer value
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = "^[a-z0-9A-Z]{" & iMinPasswordLength & "," & iMaxPasswordLength & "}$"
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate if iPort is integer and between min and max values
  	If Not oRegularExpression.Test(sPassword) Then

    	' Port is either not an integer or not in a valid range
    	MsgBox "The specified password is invalid. The Password has to be at least " & iMinPasswordLength & " Characters (Max: " & iMaxPasswordLength & ") allowed are A-Z, a-z and the Numbers from 0 to 9!", vbCritical, "Error" 
      
      Session.Property("checkFirstUserPassword") = "0"
      iReturn = ERROR_INSTALL_FAILURE 
    	
    End If
     
    checkPassword = iReturn
End Function

Function checkEmail() 
    Dim oRegularExpression, iReturn, sEmail
    
    'If something fails, move on
  	On Error Resume Next
    
    ' Get sHostheader
    sEmail = Session.Property("USEREMAIL")
    
    ' Set default result
    Session.Property("checkFirstUserEmail") = "1"
    iResult = ERROR_SUCCESS 
    
    ' Prepare regex validation for email value
    ' This is not the complete regular expression for cheking all possible email addresses
    ' like ones containing ip addresses and so on but it will do the job
  	Set oRegularExpression = New RegExp
  	oRegularExpression.Pattern = "^[\w-\.]+\@([\da-zA-Z-]+\.)+[\da-zA-Z-]{2,3}$" 
  	oRegularExpression.IgnoreCase = True
  	
  	' Validate email address
  	If Not oRegularExpression.Test(sEmail) Then

    	' MsgBox if the email address is invalid
    	MsgBox "The specified email is invalid!", vbCritical, "Error" 
      
      Session.Property("checkFirstUserEmail") = "0"
      iReturn = ERROR_INSTALL_FAILURE 
    	
    End If
     
    checkEmail = iReturn
End Function

Function md5Password()
  Dim oShell, sExecString, oPass, sPath, oFile

  ON error resume next

  ' sPHPEXE 
  sPHPEXE = Session.Property("SCRIPTEXECUTABLE")

  ' Creating Filesystem object
  Set oFileSystem = CreateObject("Scripting.FileSystemObject")

  ' Get path
  Set oFile = oFileSystem.GetFile(sPHPEXE)
  sPath = Replace(oFile.Path, oFile.Name, "")

  ' Create batch file to process php command (in order to avoid a cmd popup in the installer)
  Set oWriteFile = oFileSystem.OpenTextFile(Session.Property("TempFolder") & "tmpMD5.bat", 2, true) 'Writing
     oWriteFile.WriteLine("@Echo off")
     oWriteFile.WriteLine(sPHPEXE & " -c " & sPath & " %temp%\tmpMD5.php >%temp%\phpmd5.txt")
  oWriteFile.Close
  
  ' Create php file
  Set oWriteFile = oFileSystem.OpenTextFile(Session.Property("TempFolder") & "tmpMD5.php", 2, true) 'Writing
  If Session.Property("MD5PREFIX")="randomSalt" Then
  	oWriteFile.WriteLine("<?php function osc_rand($min = null, $max = null) {")
    oWriteFile.WriteLine("static $seeded;")
    oWriteFile.WriteLine("if (!isset($seeded)) {mt_srand((double)microtime()*1000000); $seeded = true;}")
    oWriteFile.WriteLine("if (isset($min) && isset($max)) { if ($min >= $max) { return $min;} else {return mt_rand($min, $max);}} else {return mt_rand();}")
	oWriteFile.WriteLine("}")
	oWriteFile.WriteLine("function osc_encrypt_string($plain) {")
    oWriteFile.WriteLine(" $password = '';")
    oWriteFile.WriteLine(" for ($i=0; $i<10; $i++) {")
    oWriteFile.WriteLine("  $password .= osc_rand();")
    oWriteFile.WriteLine("}")
    oWriteFile.WriteLine("$salt = substr(md5($password), 0, 2);")
    oWriteFile.WriteLine("$password = md5($salt . $plain) . ':' . $salt;")
    oWriteFile.WriteLine("return $password;")
    oWriteFile.WriteLine("}")
    oWriteFile.WriteLine(" echo osc_encrypt_string('" & Session.Property("USERPASSWORD") & "');?>")
  Else
     oWriteFile.WriteLine("<?php echo md5('" & Session.Property("MD5PREFIX") & Session.Property("USERPASSWORD") & "');?>")
  End If
  oWriteFile.Close

  ' Execute batchfile
  Set oShell = CreateObject("WSCript.Shell")
  oShell.Run Session.Property("TempFolder") & "tmpmd5.bat" , 0, True
  
  ' Get tmp output file
  Set oReadFile = oFileSystem.OpenTextFile(Session.Property("TempFolder") & "phpmd5.txt", 1) 'Reading

  ' Put all configured php modules into an array
  iCount = 0
  Do Until oReadFile.AtEndOfStream
	  Session.Property("USERMD5PASSWORD") = oReadFile.Readline
  Loop
  oReadFile.Close()

  ' Delete tmp files
  oFileSystem.DeleteFile(Session.Property("TempFolder") & "tmpmd5.bat")
  oFileSystem.DeleteFile(Session.Property("TempFolder") & "tmpMD5.php")
  oFileSystem.DeleteFile(Session.Property("TempFolder") & "phpmd5.txt")
	
  md5Password = ERROR_SUCCESS
End Function