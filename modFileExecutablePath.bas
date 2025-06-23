Attribute VB_Name = "modFileExecutablePath"
'https://visualbasic.happycodings.com/files-directories-drives/code28.html?i=1
Option Explicit

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal sResult As String) As Long

'Purpose     :  Returns the path of the executable associated with a specified file
'Inputs      :  sFileName                   The name of the file to find the executable for.
'               sDirectory                  The name of the path to find the executable for.
'Outputs     :  Returns the path of the executable for the specified file
'Notes       :  eg. FileExecutablePath("book1.xls","C:\") would return "C:\Program Files\Microsoft Office\Office\excel.exe"


Function FileExecutablePath(sFileName As String, sDirectory As String) As String
    Const MAX_PATH As Long = 260, ERROR_FILE_NO_ASSOCIATION As Long = 31, ERROR_FILE_NOT_FOUND As Long = 2
    Const ERROR_PATH_NOT_FOUND As Long = 3, ERROR_FILE_SUCCESS As Long = 32, ERROR_BAD_FORMAT As Long = 11
    Dim lRetVal As Long, lPos As Long
    Dim sResult As String * MAX_PATH
    
    On Error Resume Next
    lRetVal = FindExecutable(sFileName, sDirectory, sResult)
    
    Select Case lRetVal
    
    Case ERROR_FILE_NO_ASSOCIATION
        FileExecutablePath = "No association"
    
    Case ERROR_FILE_NOT_FOUND
        FileExecutablePath = "File not found"
    
    Case ERROR_PATH_NOT_FOUND
        FileExecutablePath = "Path not found"
    
    Case ERROR_BAD_FORMAT
        FileExecutablePath = "Bad format"
    
    Case Is >= ERROR_FILE_SUCCESS
        'Found path to executable
        lPos = InStr(sResult, Chr$(0))
        If lPos Then
           FileExecutablePath = Left$(sResult, lPos - 1)
        End If
    End Select
    On Error GoTo 0
End Function


