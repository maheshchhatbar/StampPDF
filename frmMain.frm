VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "StampPDF Wrapper Free Ver 1.0"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4170
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textboxShellCommand 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   7215
   End
   Begin VB.TextBox textboxAppend 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   7215
   End
   Begin VB.TextBox textboxFilename 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   7215
   End
   Begin VB.TextBox textboxPrepend 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "LH - "
      Top             =   960
      Width           =   7215
   End
   Begin VB.TextBox textboxLetterhead 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "d:\StampPDF_Wrapper\Letterhead.pdf"
      Top             =   600
      Width           =   7215
   End
   Begin VB.TextBox txtboxPDFTK 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe"
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label7 
      Caption         =   "Shell command executed"
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Apend to filename"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   $"frmMain.frx":0000
      Height          =   1455
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   7215
   End
   Begin VB.Label Label4 
      Caption         =   "Filename of generated file"
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Prepend to filename"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Path and filename for Letterhead to use"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Path with Filename to pdftk.exe"
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As _
     String, ByVal lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
Dim pdftkpath As String
If CheckPath(txtboxPDFTK.Text) = True Then
    txtboxPDFTK.Enabled = False
    Else
    If txtboxPDFTK.Enabled = True Then
    pdftkpath = FileExecutablePath("pdftk.exe", "C:\")
    If pdftkpath <> "" Then
        txtboxPDFTK.Text = pdftkpath
        txtboxPDFTK.Enabled = False
    End If
    End If
    
End If
If CheckPath(textboxLetterhead.Text) = True Then
    textboxLetterhead.Enabled = False
End If

End Sub

'https://stackoverflow.com/questions/23374731/how-to-drag-and-drop-files-onto-vb6-app
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim intFile As Integer
  With Data
    For intFile = 1 To .Files.Count
      Print Data.Files.Item(intFile)
      'MsgBox Data.Files.Item(intFile)
      MergePDFs (Data.Files.Item(intFile))
    Next intFile
  End With 'Data
End Sub

Private Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function


Public Function MergePDFs(strFileName As String)
If Right(strFileName, 4) <> ".pdf" Then
    MsgBox "Please drop pdf files"
    Exit Function
End If
'MsgBox "Path is " & GetFilePath(strFileName)
'MsgBox "FileName is " & GetFileName(strFileName)
Dim strCommand As String
strCommand = """" & txtboxPDFTK.Text & """" & " " & """" & strFileName & """" & " stamp " & """" & textboxLetterhead.Text & """" & " output " & """" & GetFilePath(strFileName) & textboxPrepend.Text & GetFileName(strFileName) & """"
strCommand = Left(strCommand, Len(strCommand) - 5) & textboxAppend.Text & ".pdf" & """"
'MsgBox strCommand
Debug.Print strCommand
Shell strCommand
textboxFilename.Text = GetFilePath(strFileName) & textboxPrepend.Text & Left(GetFileName(strFileName), Len(GetFileName(strFileName)) - 4) & textboxAppend.Text & ".pdf"
textboxShellCommand = strCommand
End Function

