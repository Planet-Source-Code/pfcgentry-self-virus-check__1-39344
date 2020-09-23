Attribute VB_Name = "modSubMain"
'Please compile this Project before Running.
'Name the Exe file as SelfCheck.exe

Const OPEN_EXISTING = 3

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


Sub Copy_File(File, Destination)
'Can also be used to Copy a file Using API. (Not used in this Project)
retval = CopyFile(File, Destination, 1)
End Sub

Sub Move_File(File, Destination)
'Can also be used to Move a file Using API. (Not used in this Project)
retval = MoveFile(File, Destination)
End Sub

Sub Main()

'Load frmMain
frmMain.Show

End Sub

Function Virus_Check()
    
    'Case an Error Happen that is not Handled like a Renameing of the EXE File goto the Error Handler.
    On Error GoTo CheckUnknown

    'MAke a Copy of the file and place in a Temp Dir so it can be checked.
    Call Copy_File(App.Path & "\SelfCheck.exe", "C:\SizeCheck.dpg")
     
    Dim hFile As Long, nSize As Currency
    'Create the Temp file size.
    hFile = CreateFile("C:\SizeCheck.dpg", GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    'GEt the file size.
    GetFileSizeEx hFile, nSize
       
    'Place the Size of the App in Bytes here. Just Replace 2.0480 with your apps byte size.
    If nSize <> "2.0480" Then
    'IF the sizes are different from the Exe's then Display the Warning Message.
    MsgBox "***WARNING*** this files size should be 20480 bytes, and has been altered in some way, maybe a virus or a trojan has been added to it or the File Name has been Changed. The files new size is " & Str$(nSize * 10000) & " bytes, please delete this file as soon as possible!", vbCritical, "***WARNING***"
    'Tell the user you are closeing the Application.
    MsgBox "For safety Reasons this Application will now close", vbExclamation, "Closeing"
    'Unload the Main Form.
    Unload frmMain
    'Kill the Temp File
    Kill ("C:\SizeCheck.dpg")
    'End the Application
    End
    Else:
    'Do Nothing
    'KIll the Temp file
    Kill ("C:\SizeCheck.dpg")
    End If
    
    'Unknown Error Most likely the FIle has been Renamed.
CheckUnknown:
    
    'Runtime error '53 Means file was not found. SO in this case it was probally renamed from it's origional Exe Name when Compiled.
    If Err.Number = "53" Then
    'Display the Error Message
    MsgBox "Unknown Error Occured while self Virus/Trojan/EXE Name Change checking, This file has Possibly been renamed to something else, This files EXE name should be ""SelfCheck.exe"""
    'End the Application
    End
    End If
   
End Function
