Attribute VB_Name = "modGetFilePath"
'https://www.devhut.net/vba-extract-the-path-from-a-file-name/
'---------------------------------------------------------------------------------------
' Procedure : GetFilePath
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Return the path from a path\filename input
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - string of a path and filename (ie: "c:\temp\test.xls")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2008-02-06              Initial Release
'---------------------------------------------------------------------------------------
Function GetFilePath(sFile As String)
On Error GoTo Err_Handler
    
    GetFilePath = Left(sFile, InStrRev(sFile, "\"))

Exit_Err_Handler:
    Exit Function

Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: GetFilePath" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Function GetFileName(ByVal sFile As String)
    On Error GoTo Error_Handler

    GetFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: GetFileName" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function


