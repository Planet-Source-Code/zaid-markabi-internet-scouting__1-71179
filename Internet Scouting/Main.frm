VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   Caption         =   "Internet Scouting"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   10590
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9600
      Top             =   840
   End
   Begin VB.ListBox History 
      BackColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   2775
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      ExtentX         =   5106
      ExtentY         =   4895
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
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   59
      ImageHeight     =   47
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":23D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2F90
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":53A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":62A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":8402
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":90D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":9DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A98C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7560
      ScaleHeight     =   735
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   315
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Loading ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   50
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   10530
      TabIndex        =   1
      Top             =   6090
      Width           =   10590
      Begin VB.Label StatusLabel 
         AutoSize        =   -1  'True
         Caption         =   "Loading ..."
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   750
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6960
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Image Image9 
      Height          =   705
      Left            =   6720
      Picture         =   "Main.frx":B566
      Top             =   0
      Width           =   885
   End
   Begin VB.Image Image8 
      Height          =   705
      Left            =   5880
      Picture         =   "Main.frx":D6B4
      Top             =   0
      Width           =   885
   End
   Begin VB.Image Image7 
      Height          =   705
      Left            =   5040
      Picture         =   "Main.frx":E362
      Top             =   0
      Width           =   885
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   4200
      Picture         =   "Main.frx":F118
      Top             =   0
      Width           =   885
   End
   Begin VB.Image Image5 
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image6 
      Height          =   735
      Left            =   840
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   1680
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   3360
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   2520
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open Web Page"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save Web Page"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit ( Close All Tabs )"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnutab 
      Caption         =   "Tab"
      Begin VB.Menu mnunewtab 
         Caption         =   "&New Tab"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuclosetab 
         Caption         =   "&Close Tab"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclosealltabs 
         Caption         =   "Close All Tabs"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnucommand 
      Caption         =   "Command"
      Begin VB.Menu mnuback 
         Caption         =   "Back"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuforward 
         Caption         =   "Forward"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuhome 
         Caption         =   "Go to Home Page"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuhistory 
      Caption         =   "History"
      Begin VB.Menu mnushow 
         Caption         =   "Show\Hide"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclearhistury 
         Caption         =   "&Clear History"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SaveHistory()
Open App.Path + "\Data\History.zaid" For Output As #1
Write #1, History.ListCount
Dim i As Integer
For i = 0 To History.ListCount - 1
Write #1, History.List(i)
Next
Close #1
End Sub

Private Sub LoadHistory()
On Error GoTo 1
History.Clear
Open App.Path + "\Data\History.zaid" For Input As #1
Dim i2 As Integer
Dim x As String
Input #1, i2
Dim i As Integer
For i = 0 To i2 - 1
Input #1, x
History.AddItem x, History.ListCount
Next
1:
Close #1
End Sub

Private Sub Form_Load()
WebBrowser.Navigate "http:\\www.google.com\"
Image5.Picture = ImageList.ListImages.Item(1).Picture
Image6.Picture = ImageList.ListImages.Item(3).Picture
Image3.Picture = ImageList.ListImages.Item(5).Picture
Image4.Picture = ImageList.ListImages.Item(7).Picture
Image2.Picture = ImageList.ListImages.Item(9).Picture
LoadHistory
End Sub

Private Sub Form_Resize()
On Error Resume Next
If History.Visible = True Then
WebBrowser.Left = History.Width
Else
WebBrowser.Left = 0
End If
Picture2.Width = Me.Width - 100 - Picture2.Left
Option1.Left = Picture2.Width - Option1.Width
Text1.Width = Picture2.Width - 100 - (Text1.Left * 2)
WebBrowser.Width = Me.Width - 100 - WebBrowser.Left
WebBrowser.Height = Me.Height - WebBrowser.Top - 1000
History.Height = Me.Height - History.Top - 1000
End Sub

Private Sub History_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub History_Click()
On Error Resume Next
Text1.Text = Trim(History.List(History.ListIndex))
WebBrowser.Navigate Text1.Text
End Sub

Private Sub Image1_Click()
SaveHistory
Set WepPage32(WepPageNum) = New Main
WepPage32(WepPageNum).Show
End Sub

Private Sub Image2_Click()
WebBrowser.Stop
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Picture = ImageList.ListImages.Item(10).Picture
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Picture = ImageList.ListImages.Item(9).Picture
End Sub

Private Sub Image3_Click()
WebBrowser.GoHome
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image3.Picture = ImageList.ListImages.Item(6).Picture
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image3.Picture = ImageList.ListImages.Item(5).Picture
End Sub

Private Sub Image4_Click()
WebBrowser.Refresh
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image4.Picture = ImageList.ListImages.Item(8).Picture
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image4.Picture = ImageList.ListImages.Item(7).Picture
End Sub

Private Sub Image5_Click()
On Error Resume Next
WebBrowser.GoBack
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Picture = ImageList.ListImages.Item(2).Picture
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Picture = ImageList.ListImages.Item(1).Picture
End Sub

Private Sub Image6_Click()
On Error Resume Next
WebBrowser.GoForward
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image6.Picture = ImageList.ListImages.Item(4).Picture
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image6.Picture = ImageList.ListImages.Item(3).Picture
End Sub

Private Sub Image7_Click()
mnusave_Click
End Sub

Private Sub Image8_Click()
mnuopen_Click
End Sub

Private Sub Image9_Click()
History.Visible = (Not History.Visible)
Form_Resize
End Sub

Private Sub mnuback_Click()
Image5_Click
End Sub

Private Sub mnuclearhistury_Click()
On Error Resume Next
Kill App.Path + "\Data\History.zaid"
LoadHistory
End Sub

Private Sub mnuclosealltabs_Click()
End
End Sub

Private Sub mnuclosetab_Click()
Unload Me
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuforward_Click()
Image6_Click
End Sub

Private Sub mnuhome_Click()
WebBrowser.GoHome
End Sub

Private Sub mnunewtab_Click()
Image1_Click
End Sub

Private Sub mnuopen_Click()
On Error GoTo 1
CommonDialog.ShowOpen
WebBrowser.Navigate CommonDialog.FileName
1:
End Sub

Private Sub mnurefresh_Click()
WebBrowser.Refresh
End Sub

Private Sub mnusave_Click()
On Error Resume Next
WebBrowser.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub mnushow_Click()
Image9_Click
End Sub

Private Sub mnustop_Click()
WebBrowser.Stop
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 10 Then
If Not Left(Text1.Text, 11) = "http://www." Then
Text1.Text = "http://www." + Text1.Text + ".com/"
WebBrowser.Navigate Text1.Text
End If
End If

If KeyAscii = 13 Then
WebBrowser.Navigate Text1.Text
End If
End Sub

Private Sub Timer1_Timer()
If IsConnected = True Then
Option1.Value = True
Option1.Caption = "Connected ..."
Else
Option1.Value = False
Option1.Caption = "DisConnected ..."
End If
End Sub

Private Sub WebBrowser_StatusTextChange(ByVal Text As String)
StatusLabel.Caption = Text
End Sub

Private Sub WebBrowser_TitleChange(ByVal Text As String)
Text1.Text = WebBrowser.LocationURL
Me.Caption = "Internet Scouting    -    " + Text
' Add to history
LoadHistory
Dim i As Integer
For i = 0 To History.ListCount - 1
If Trim(History.List(i)) = WebBrowser.LocationURL Then
GoTo 1
End If
Next
For i = 0 To History.ListCount - 1
If History.List(i) = Left(WebBrowser.LocationURL, Len(History.List(i))) Then
History.AddItem "   " + WebBrowser.LocationURL, i + 1
GoTo 1
End If
Next
History.AddItem WebBrowser.LocationURL
1:
SaveHistory
End Sub
