VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "HTML to EXE"
   ClientHeight    =   6945
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog C 
      Left            =   5400
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Code 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":0CCA
      Top             =   0
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu Open 
         Caption         =   "Open"
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
      End
      Begin VB.Menu Comp 
         Caption         =   "Compile To EXE"
      End
      Begin VB.Menu v 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Comp_Click()
C.Filter = "EXE files (*.exe)|*.exe"
C.ShowSave
If C.FileName <> "" Then
DLLFILE = App.Path & "\comp.dll"
APPFILE = C.FileName
FileCopy DLLFILE, APPFILE
    PUTINF = "<%text%>" & Code.Text
    File1$ = APPFILE
    File2$ = DLLFILE
    
    Open File1$ For Output As #1        'Open Application
    Open File2$ For Binary As #2        'Open DLL File
    Do While Not EOF(2)
        FileData = Input$(2000, #2)
        msg = FileData
        msg2 = msg2 + msg
        Print #1, msg2;
        msg2 = ""
        If Len(msg) > 2000 Then
            msg = ""
        End If
    Loop
    Print #1, PUTINF                    'Application
    Close #2                            'Close DLL File
    Close #1                            'Close Application
    
    Shell APPFILE, vbNormalFocus
    End If
End Sub

Private Sub Exit_Click()
    Unload Form2
    Unload Me
End Sub

Private Sub Form_Resize()
On Error GoTo e
    Code.Width = Me.Width - 120
    Code.Height = Me.Height - 690
e:
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload Form2
End Sub

Private Sub mnuRun_Click()
    Form2.Show
End Sub

Private Sub Open_Click()
Wrap$ = Chr$(13) + Chr$(10)
    C.Filter = "HTMLtoEXE File (*.hte)|*.hte"
    C.ShowOpen
    If C.FileName <> "" Then
        Form1.MousePointer = 11
        Open C.FileName For Input As #1
        On Error GoTo Giant:
        Do Until EOF(1)

            Line Input #1, LineOfText$
            AllText$ = AllText$ & LineOfText$ & Wrap$
        Loop
        twothings = Split(AllText$, vbCrLf & "---" & vbCrLf)
        Code.Text = twothings(0)
        Code.Enabled = True
Fixit:
        Form1.MousePointer = 0
        Close #1
    End If
    Exit Sub
Giant:
    MsgBox "Error: This file is too large to open!", vbCritical, "Error!"
    Resume Fixit:
End Sub

Private Sub Save_Click()
    C.Filter = "HTMLtoEXE File (*.hte)|*.hte"
    C.ShowSave
    If C.FileName <> "" Then
        Open C.FileName For Output As #1
        Print #1, Code.Text
        Close #1
    End If
End Sub
