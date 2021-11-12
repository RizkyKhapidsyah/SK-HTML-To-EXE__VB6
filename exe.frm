VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   Icon            =   "exe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   960
   End
   Begin VB.TextBox Code 
      Height          =   2895
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
      Visible         =   0   'False
   End
   Begin SHDocVwCtl.WebBrowser w 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      ExtentX         =   5530
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
    filesize = LOF(1)
    FileData$ = Space$(LOF(1))
    Get #1, , FileData$
    For i = 1 To filesize
            If Mid(FileData$, i, 8) = "<%text%>" Then
                i = i + 8
                filechunk$ = Space$(10000)
                Get #1, i, filechunk$
                Code.Text = filechunk$
                Exit Sub
            End If
        Next i
        Close #1
        

End Sub

Private Sub Form_Resize()
On Error GoTo e
    w.Width = Me.Width - 120
    w.Height = Me.Height - 400
e:
End Sub

Private Sub Timer1_Timer()
    w.Navigate "about:" & Code.Text
    Timer1.Enabled = False
End Sub

Private Sub w_StatusTextChange(ByVal Text As String)
Me.Caption = w.LocationName
End Sub
