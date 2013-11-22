Function createScriptConfig ()
	Const ERROR_INSTALL_FAILURE = 1603 
	Const ERROR_SUCCESS = 0 
	Const msiMessageTypeInfo = &H04000000
	Const msiMessageTypeError = &H01000000
	Const msiMessageTypeWarning = &H02000000
	Const ForReading = 1
	Const ForWriting = 2
	Dim aVariablesToParse, oFileSystem, oReadFile, oWriteFile, sFileList, aFiles, iCount, aLines(), iLineCount, iFileCount
	' Variables to take from WIX and write to php files
	aVariablesToParse = array("ProductName","Date","Time","OriginalDatabase","TARGETMODE","WEBAPPLICATIONNAME","WEBAPPLICATIONPOOLNAME","SCRIPTPARSE","RequireSQLDatabase","DATABASEENGINE","WEBSITEIP", "WEBSITEPORT", "WEBSITEDESCRIPTION", "WEBSITEHOSTHEADER", "TARGETDIR", "VIRTUALDIRECTORYNAME", "SQLDATABASE", "DBHOST","DBHOSTANDPORT", "SQLUSERUSERNAME", "SQLUSERPASSWORD", "USERUSERNAME", "USERPASSWORD", "USERMD5PASSWORD", "USEREMAIL", "LOCALURL", "FIRSTIP", "DBPORT", "DBCLIENT", "URLHOST", "URLDIR")
	
	on error resume next
	Err.Clear
	
	' Creating Filesystem object
	Set oFileSystem = CreateObject("Scripting.FileSystemObject")
  
	' Get template files from WIX
	sFileList = Session.Property("SCRIPTPARSE")
  
	' Path from WIX (users choice)
	sPath = Session.Property("TARGETDIR")
  
  
	If sFileList <> "" Then
		aFiles = Split(Session.Property("SCRIPTPARSE"), ",")
	
		For iFileCount = LBound(aFiles) To UBound(aFiles)
			sFile = Trim(aFiles(iFileCount))

			if oFileSystem.FileExists(sPath & sFile) Then
        
				' Open template
				Set oReadFile = oFileSystem.OpenTextFile(sPath & sFile, ForReading) 'Reading
		 
				strFile = oReadFile.ReadAll
				oReadFile.Close
		 
				' Open target file
				Set oWriteFile = oFileSystem.OpenTextFile(sPath & sFile, ForWriting, true) 'Writing

				For iCount = LBound(aVariablesToParse) To UBound(aVariablesToParse)
					If InStr(1, strFile, "@@" & aVariablesToParse(iCount) & "@@") > 0 Then
						strFile = Replace(strFile, "@@" & aVariablesToParse(iCount) & "@@", Replace(Session.Property(aVariablesToParse(iCount)), "\", "\\"))
					End if
				Next

				oWriteFile.Write strFile
				oWriteFile.Close
			end if
		Next
	End If

	
	'Write entry to log file and set result
	Set record = Session.Installer.CreateRecord(0)
	if (Err.Number = 0) Then 
		record.StringData(0) = "createScriptConfig.vbs -> createScriptConfig() -> all successful"
		Session.Message msiMessageTypeInfo, record
		createScriptConfig = ERROR_SUCCESS
	else
		record.StringData(0) = "createScriptConfig.vbs -> createScriptConfig() -> an error has occurred: " & Err.Description & VBCrLf & " [Number:" & Hex(Err.Number) & "]"
		Session.Message msiMessageTypeInfo, record
		createScriptConfig = ERROR_INSTALL_FAILURE
	end if

End Function