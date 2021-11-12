VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form2 
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6705
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser W 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      ExtentX         =   7223
      ExtentY         =   7646
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    W.Navigate "about:" & Form1.Code.Text
End Sub

Private Sub Form_Resize()
On Error GoTo e
    W.Width = Me.Width - 120
    W.Height = Me.Height - 400
e:
End Sub

Private Sub W_StatusTextChange(ByVal Text As String)
    Me.Caption = W.LocationName
End Sub
