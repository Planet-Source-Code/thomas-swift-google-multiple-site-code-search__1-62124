VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Google VB Code Search"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser12 
      Height          =   525
      Left            =   7590
      TabIndex        =   15
      Top             =   1665
      Visible         =   0   'False
      Width           =   495
      ExtentX         =   873
      ExtentY         =   926
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
   Begin SHDocVwCtl.WebBrowser WebBrowser11 
      Height          =   555
      Left            =   6900
      TabIndex        =   14
      Top             =   1620
      Visible         =   0   'False
      Width           =   525
      ExtentX         =   926
      ExtentY         =   979
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   1050
      Left            =   150
      TabIndex        =   13
      Top             =   495
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   1852
      _Version        =   393216
      Tabs            =   12
      Tab             =   9
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "MSDN"
      TabPicture(0)   =   "Form1.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Planet Source Code"
      TabPicture(1)   =   "Form1.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "A1 VB Code"
      TabPicture(2)   =   "Form1.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Code Archive"
      TabPicture(3)   =   "Form1.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Code Project"
      TabPicture(4)   =   "Form1.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Experts Exchange"
      TabPicture(5)   =   "Form1.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Free VB Code"
      TabPicture(6)   =   "Form1.frx":0632
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Programmers Heaven"
      TabPicture(7)   =   "Form1.frx":064E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "VB City"
      TabPicture(8)   =   "Form1.frx":066A
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "VB Code"
      TabPicture(9)   =   "Form1.frx":0686
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "VB Users"
      TabPicture(10)  =   "Form1.frx":06A2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Dev X"
      TabPicture(11)  =   "Form1.frx":06BE
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser10 
      Height          =   600
      Left            =   6300
      TabIndex        =   12
      Top             =   1605
      Visible         =   0   'False
      Width           =   525
      ExtentX         =   926
      ExtentY         =   1058
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
   Begin SHDocVwCtl.WebBrowser WebBrowser9 
      Height          =   585
      Left            =   5580
      TabIndex        =   11
      Top             =   1620
      Visible         =   0   'False
      Width           =   615
      ExtentX         =   1085
      ExtentY         =   1032
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
   Begin SHDocVwCtl.WebBrowser WebBrowser8 
      Height          =   525
      Left            =   4845
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
      ExtentX         =   1085
      ExtentY         =   926
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
   Begin SHDocVwCtl.WebBrowser WebBrowser7 
      Height          =   495
      Left            =   4155
      TabIndex        =   9
      Top             =   1710
      Visible         =   0   'False
      Width           =   495
      ExtentX         =   873
      ExtentY         =   873
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
   Begin SHDocVwCtl.WebBrowser WebBrowser6 
      Height          =   465
      Left            =   3480
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   540
      ExtentX         =   952
      ExtentY         =   820
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
   Begin SHDocVwCtl.WebBrowser WebBrowser5 
      Height          =   510
      Left            =   2910
      TabIndex        =   7
      Top             =   1725
      Visible         =   0   'False
      Width           =   390
      ExtentX         =   688
      ExtentY         =   900
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
   Begin SHDocVwCtl.WebBrowser WebBrowser4 
      Height          =   570
      Left            =   2265
      TabIndex        =   6
      Top             =   1710
      Visible         =   0   'False
      Width           =   510
      ExtentX         =   900
      ExtentY         =   1005
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
   Begin SHDocVwCtl.WebBrowser WebBrowser3 
      Height          =   525
      Left            =   1500
      TabIndex        =   5
      Top             =   1755
      Visible         =   0   'False
      Width           =   600
      ExtentX         =   1058
      ExtentY         =   926
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
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   585
      Left            =   855
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   555
      ExtentX         =   979
      ExtentY         =   1032
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
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   525
      Left            =   195
      TabIndex        =   3
      Top             =   1695
      Width           =   570
      ExtentX         =   1005
      ExtentY         =   926
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   450
      Left            =   780
      TabIndex        =   0
      Top             =   30
      Width           =   6735
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         Height          =   285
         Left            =   900
         TabIndex        =   17
         Top             =   75
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   75
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   285
         Left            =   5640
         TabIndex        =   2
         Top             =   75
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   "Flowers"
         Top             =   75
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SSTab1.Tab = 0
Call WebBrowser1.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://msdn.microsoft.com")
Call WebBrowser2.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://planet-source-code.com")
Call WebBrowser3.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://a1vbcode.com")
Call WebBrowser4.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://codearchive.com")
Call WebBrowser5.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://codeproject.com")
Call WebBrowser6.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://experts-exchange.com")
Call WebBrowser7.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://freevbcode.com")
Call WebBrowser8.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://programmersheaven.com")
Call WebBrowser9.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://vbcity.com")
Call WebBrowser10.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://vbcode.com")
Call WebBrowser11.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://vbusers.com")
Call WebBrowser12.Navigate("www.google.com/search?&q=" & Text1.Text & " site:http://devx.com")
End Sub
Private Sub Command2_Click()
On Error Resume Next
Select Case SSTab1.Tab
        Case 0: WebBrowser1.GoBack
        Case 1: WebBrowser2.GoBack
        Case 2: WebBrowser3.GoBack
        Case 3: WebBrowser4.GoBack
        Case 4: WebBrowser5.GoBack
        Case 5: WebBrowser6.GoBack
        Case 6: WebBrowser7.GoBack
        Case 7: WebBrowser8.GoBack
        Case 8: WebBrowser9.GoBack
        Case 9: WebBrowser10.GoBack
        Case 10: WebBrowser11.GoBack
        Case 11: WebBrowser12.GoBack
End Select
End Sub
Private Sub Command3_Click()
On Error Resume Next
Select Case SSTab1.Tab
        Case 0: WebBrowser1.GoForward
        Case 1: WebBrowser2.GoForward
        Case 2: WebBrowser3.GoForward
        Case 3: WebBrowser4.GoForward
        Case 4: WebBrowser5.GoForward
        Case 5: WebBrowser6.GoForward
        Case 6: WebBrowser7.GoForward
        Case 7: WebBrowser8.GoForward
        Case 8: WebBrowser9.GoForward
        Case 9: WebBrowser10.GoForward
        Case 10: WebBrowser11.GoForward
        Case 11: WebBrowser12.GoForward
End Select
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Top = Me.ScaleTop
Frame1.Left = Me.ScaleWidth / 2 - Frame1.Width / 2

SSTab1.Top = Me.ScaleTop + Frame1.Height
SSTab1.Left = Me.ScaleLeft
SSTab1.Width = Me.ScaleWidth

WebBrowser1.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser1.Left = Me.ScaleLeft + 100
WebBrowser1.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser1.Width = Me.ScaleWidth - 200

WebBrowser2.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser2.Left = Me.ScaleLeft + 100
WebBrowser2.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser2.Width = Me.ScaleWidth - 200

WebBrowser3.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser3.Left = Me.ScaleLeft + 100
WebBrowser3.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser3.Width = Me.ScaleWidth - 200

WebBrowser4.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser4.Left = Me.ScaleLeft + 100
WebBrowser4.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser4.Width = Me.ScaleWidth - 200

WebBrowser5.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser5.Left = Me.ScaleLeft + 100
WebBrowser5.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser5.Width = Me.ScaleWidth - 200

WebBrowser6.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser6.Left = Me.ScaleLeft + 100
WebBrowser6.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser6.Width = Me.ScaleWidth - 200

WebBrowser7.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser7.Left = Me.ScaleLeft + 100
WebBrowser7.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser7.Width = Me.ScaleWidth - 200

WebBrowser8.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser8.Left = Me.ScaleLeft + 100
WebBrowser8.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser8.Width = Me.ScaleWidth - 200

WebBrowser9.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser9.Left = Me.ScaleLeft + 100
WebBrowser9.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser9.Width = Me.ScaleWidth - 200

WebBrowser10.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser10.Left = Me.ScaleLeft + 100
WebBrowser10.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser10.Width = Me.ScaleWidth - 200

WebBrowser11.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser11.Left = Me.ScaleLeft + 100
WebBrowser11.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser11.Width = Me.ScaleWidth - 200

WebBrowser12.Top = Me.ScaleTop + Frame1.Height + SSTab1.Height + 100
WebBrowser12.Left = Me.ScaleLeft + 100
WebBrowser12.Height = (Me.ScaleHeight - Frame1.Height) - (SSTab1.Height + 600)
WebBrowser12.Width = Me.ScaleWidth - 200

End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False
WebBrowser11.Visible = False
WebBrowser12.Visible = False
Select Case SSTab1.Tab
        Case 0: WebBrowser1.Visible = True
        Case 1: WebBrowser2.Visible = True
        Case 2: WebBrowser3.Visible = True
        Case 3: WebBrowser4.Visible = True
        Case 4: WebBrowser5.Visible = True
        Case 5: WebBrowser6.Visible = True
        Case 6: WebBrowser7.Visible = True
        Case 7: WebBrowser8.Visible = True
        Case 8: WebBrowser9.Visible = True
        Case 9: WebBrowser10.Visible = True
        Case 10: WebBrowser11.Visible = True
        Case 11: WebBrowser12.Visible = True
End Select
End Sub



