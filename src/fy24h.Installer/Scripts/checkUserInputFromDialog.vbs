
Function checkCRMAppInput( )
    MsgBox "start"
    
    Dim crmAppVirtualDirectoryName, wcfAttachmentDirectoryName, crmAppValidFg

    crmAppVirtualDirectoryName =  Session.Property("VIRTUALDIRECTORYNAME");
    
    crmAppValidFg = "0"
    Session.Property("CRMAPPVALIDFG") = crmAppValidFg
    MsgBox "CRMAPPVALIDFG end"
'    Session.Property("SQLSERVERCONNECTIONCHECK") = "0"

    'MsgBox crmAppVirtualDirectoryName
    MsgBox "crmAppVirtualDirectoryName = '' "

    If crmAppVirtualDirectoryName = "" Then

        MsgBox "安装目录不能为空", vbCritical

        Exit Function
    End If

    wcfAttachmentDirectoryName = Session.Property("DIRATTACHMENTS")

    'MsgBox wcfAttachmentDirectoryName

    If wcfAttachmentDirectoryName = "" Then
    
        MsgBox "文件存放目录不能为空", vbCritical

        Exit Function

    End If

    crmAppValidFg = "1" 

    Session.Property("CRMAPPVALIDFG") = crmAppValidFg
End Function


Function checkDatabaseInput( )
    
    Dim dbHostName, dbName, dbUserName, dbPassword, databaseValidFg

    dbHostName = Session.Property("DBHOST")

    databaseValidFg = "0"
    Session.Property("DATABASEVALIDFG") = databaseValidFg

   'MsgBox crmAppVirtualDirectoryName

    If dbHostName = "" Then
        
        MsgBox "服务器名称不能为空", vbCritical 

        Exit Function
    End If

    dbName = Session.Property("SQLDATABASE")

    'MsgBox wcfAttachmentDirectoryName

    If dbName = "" Then
    
        MsgBox "数据库名称不能为空", vbCritical 

        Exit Function

    End If

    dbUserName = Session.Property("SQLADMINUSERNAME")
    If dbUserName = "" Then

        databaseValidFg = "0"
    
        MsgBox "用户名不能为空", vbCritica

        Exit Function

    End If

    dbPassword = Session.Property("SQLADMINPASSWORD")
    If dbPassword = "" Then

        databaseValidFg = "0"
    
        MsgBox "密码不能为空", vbCritical

        Exit Function

    End If

    databaseValidFg = "1"
    Session.Property("DATABASEVALIDFG") = databaseValidFg
End Function

Function showTestFailureMsg()
    Dim dbConnectionResult 

    dbConnectionResult = Session.Property("SQLSERVERCONNECTIONCHECK")

    If dbConnectionResult = "0" Then
    
        MsgBox "连接服务器失败", vbCritical

        Exit Function
    Else
         MsgBox "连接成功", vbInformation
    End If

End Function