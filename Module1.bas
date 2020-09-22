Attribute VB_Name = "mdlMain"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Dim aPath As String, aINIFile As String
Dim aDat As String, k As Long


Sub Main()
    aPath = App.Path
    If Right$(aPath, 1) <> "\" Then aPath = aPath & "\"
    aINIFile = aPath & "Settings.ini"
    aDat = Space$(128)
    
    frmEMail.txtSubject = GetSetting("Subject", "E-Mail")
    frmEMail.txtMail = GetSetting("E-Mail", "E-Mail")
    frmEMail.txtName = GetSetting("Name", "E-Mail")
   
    WebServer.txtWebFolder = GetSetting("WebFolder", "Options")
    WebServer.txtTotalVisitors = GetSetting("Visitors", "Options")
    
    
End Sub

Function GetSetting(strSettingName, strFolder)
k = GetPrivateProfileString(strFolder, strSettingName, "", aDat, 128, aINIFile)
GetSetting = Left$(aDat, k)
End Function

Function WriteSetting(strSettingName, strFolder, strValue)
WritePrivateProfileString strFolder, strSettingName, strValue, aINIFile
End Function

Function LoadEMail()
frmEMail.txtSubject = GetSetting("Subject", "E-Mail")
frmEMail.txtMail = GetSetting("E-Mail", "E-Mail")
frmEMail.txtName = GetSetting("Name", "E-Mail")
End Function

Function SaveEMail()
WriteSetting "Subject", "E-Mail", frmEMail.txtSubject
WriteSetting "E-Mail", "E-Mail", frmEMail.txtMail
WriteSetting "Name", "E-Mail", frmEMail.txtName
End Function
