VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmCASTml 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customize And Save Time Script Language"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   6015
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      ExtentX         =   9340
      ExtentY         =   10610
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
   Begin VB.CommandButton cmdExe 
      Caption         =   "Execute"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmCASTml.frx":0000
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmCASTml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'HTTP1.Retrieve "http://cgi.onlies.net/counter/freedom/counter.cgi"

Private Sub HTTP1_Finished(ByVal Data As String)
Text1.Text = Data$
End Sub

Private Sub cmdExe_Click()
Open App.Path & "\castscript.html" For Output As #1
 Print #1, Execute(Text1.Text)
Close #1
 Call wb.Navigate("File://" & App.Path & "\castscript.html")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Kill(App.Path & "\castscript.html")
End Sub
