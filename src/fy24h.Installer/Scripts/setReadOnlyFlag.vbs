function SetReadOnlyFlag() 
	Dim fs, f, r 
	On Error Resume Next
	Err.Clear
	
	' Creating Filesystem object
	Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	
	' Get template files from WIX
	sFileList = Session.Property("SETREADONLY")
  
	' Path from WIX (users choice)
	sPath = Session.Property("TARGETDIR")
  
  
	If sFileList <> "" Then
		aFiles = Split(Session.Property("SETREADONLY"), ",")
	
		For iFileCount = LBound(aFiles) To UBound(aFiles)
			sFile = Trim(aFiles(iFileCount))
			Set f = oFileSystem.GetFile(sPath & sFile) 
				If f.attributes and 1 Then 
					'ReadOnly Flag is set exit function
					exit function
				Else 
					f.attributes = f.attributes + 1
					writeWindowsInstallerLogEntry "SetReadOnlyFlag.vbs -> SetReadOnlyFlag() for: " & sPath & sFile,0
				End If 
		Next
	End If
  
	if (Err.Number <> 0) Then 
		writeWindowsInstallerLogEntry "SetReadOnlyFlag.vbs -> SetReadOnlyFlag() -> failed: " & Err.Description & VBCrLf & "[Number:" & Hex(Err.Number) & "]",0
		SetReadOnlyFlag = ERROR_INSTALL_FAILURE
	else
		SetReadOnlyFlag = ERROR_SUCCESS
	end if
 
end function


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