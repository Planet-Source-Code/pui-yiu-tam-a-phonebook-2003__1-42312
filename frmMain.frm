VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Agenda 2003"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7320
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   1535
      ButtonWidth     =   2143
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Phone Book"
            Key             =   "phone"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calendar"
            Key             =   "calendar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alarms"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Contact"
            Key             =   "search"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit Agenda"
            Key             =   "exit"
            ImageIndex      =   20
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5670
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4895
            Text            =   "Welcome to Agenda 2003"
            TextSave        =   "Welcome to Agenda 2003"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/01/2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   4895
            TextSave        =   "19:49"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Index           =   0
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":50D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51606
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52458
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":532AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":540FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5454E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55452
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5612C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5657E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58364
            Key             =   "JOBORDERS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5951A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B39C
            Key             =   "ENGAGEMENTS"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BC78
            Key             =   "LOGIN"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C55C
            Key             =   "ACCOUNTS"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CE40
            Key             =   "COLORS"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D15C
            Key             =   "EXIT"
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Functions 
         Caption         =   "Functions"
         Begin VB.Menu mnnuPhone 
            Caption         =   "Phone Book"
         End
         Begin VB.Menu mnuCalendar 
            Caption         =   "Calendar"
         End
         Begin VB.Menu mnuAlarmas 
            Caption         =   "Alarmas"
         End
         Begin VB.Menu mnuProperty 
            Caption         =   "Property"
         End
      End
      Begin VB.Menu News 
         Caption         =   "News"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu Printer_Configuration 
         Caption         =   "Printer Configuration"
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit Agenda"
      End
   End
   Begin VB.Menu View 
      Caption         =   "&View"
      Begin VB.Menu Colors 
         Caption         =   "Background Colors"
      End
      Begin VB.Menu Tool_Bar 
         Caption         =   "Tool Bar"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Search_Contact 
         Caption         =   "Search Contact"
      End
      Begin VB.Menu Back 
         Caption         =   "Back"
      End
      Begin VB.Menu Forware 
         Caption         =   "Forware"
      End
      Begin VB.Menu Line4 
         Caption         =   "-"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu Cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu MP3_Player 
      Caption         =   "&Music MP3 Player"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'This Code was written by Javier Castellanos Pro Soft programmer and You Cannot Delete This If
'You Are Going To Use This Code...But Have Fun And I Hope You Enjoy
'Javier Castellanos Pro Soft programmer
'If you find any bugs,
'Please E-mail me a personalgjoc@hotmail.com So
'I can get them fixed...
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Option Explicit
Public db As Database
Public rs As Recordset

Private Sub Colors_Click()
CommonDialog1.Flags = &H1&
CommonDialog1.ShowColor
MDIForm1.BackColor = CommonDialog1.Color
End Sub

Private Sub Exit_Click()
 Dim response As Integer
  response = MsgBox("Esta seguro de salir de Pro Agenda?", vbYesNo + vbExclamation, "Pro Soft")
  If response = vbYes Then
  End
  End If
End Sub

Private Sub MDIForm_Load()
Set db = OpenDatabase(App.Path & "\agenda.mdb")
With db
Set rs = .OpenRecordset("agenda")
End With
Toolbar1.Buttons.Item(4).Visible = False
End Sub

Private Sub MP3_Player_Click()
Load Mp3PlayerMain
Mp3PlayerMain.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
  Case "phone"
  Load Phone_Book
  Phone_Book.Show
  Toolbar1.Buttons.Item(4).Visible = True
  Case "calendar"
  Load Calendario
  Calendario.Show
  Case "search"
  On Error GoTo handle
Set rs = db.OpenRecordset("SELECT * FROM agenda WHERE name1='" & Phone_Book.Text1.Text & "'", dbOpenDynaset)
Phone_Book.Text2.Text = rs!name2
Phone_Book.Text3.Text = rs!direccion1
Phone_Book.Text4.Text = rs!ciudad1
Phone_Book.Text5.Text = rs!pais1
Phone_Book.Text6.Text = rs!zipostal1
Phone_Book.Text7.Text = rs!zipostal2
Phone_Book.Text8.Text = rs!trabajo2
Phone_Book.Text9.Text = rs!celular1
Phone_Book.Text10.Text = rs!fax1
Phone_Book.Text11.Text = rs!personal
Phone_Book.Text12.Text = rs!trabajo1
Phone_Book.Text13.Text = rs!trabajo3
Phone_Book.Text14.Text = rs!email1
Phone_Book.Text15.Text = rs!website1
Phone_Book.Text16.Text = rs!compania
Phone_Book.Text17.Text = rs!direccion2
Phone_Book.Text18.Text = rs!ciudad2
Phone_Book.Text19.Text = rs!zipostal3
Phone_Book.Text20.Text = rs!pais2
Phone_Book.Text21.Text = rs!email2
Phone_Book.Text22.Text = rs!website2
Phone_Book.Text23.Text = rs!comentarios1
Phone_Book.Text24.Text = rs!cumpleaños
Phone_Book.Text25.Text = rs!esposa1
Phone_Book.Text26.Text = rs!esposa2
Phone_Book.Text27.Text = rs!esposa3
Phone_Book.Text28.Text = rs!esposa4
Phone_Book.Text29.Text = rs!tel1
Phone_Book.Text30.Text = rs!tel2
Phone_Book.Text31.Text = rs!tel3
Phone_Book.Text32.Text = rs!tel4
Phone_Book.Text33.Text = rs!tel5
Phone_Book.Text34.Text = rs!tel6
Phone_Book.Text35.Text = rs!tel7
Phone_Book.Text36.Text = rs!tel8
rs.Close
Call MDIForm_Load

handle:
Select Case Err.Number
Case 3021
MsgBox "There is no such a record", vbOKOnly + vbInformation, "Pro Agenda 2003"
Phone_Book.Text1.Text = ""
Phone_Book.Text2.Text = ""
Phone_Book.Text3.Text = ""
Phone_Book.Text4.Text = ""
Phone_Book.Text5.Text = ""
Phone_Book.Text6.Text = ""
Phone_Book.Text7.Text = ""
Phone_Book.Text8.Text = ""
Phone_Book.Text9.Text = ""
Phone_Book.Text10.Text = ""
Phone_Book.Text11.Text = ""
Phone_Book.Text12.Text = ""
Phone_Book.Text13.Text = ""
Phone_Book.Text14.Text = ""
Phone_Book.Text15.Text = ""
Phone_Book.Text16.Text = ""
Phone_Book.Text17.Text = ""
Phone_Book.Text18.Text = ""
Phone_Book.Text19.Text = ""
Phone_Book.Text20.Text = ""
Phone_Book.Text21.Text = ""
Phone_Book.Text22.Text = ""
Phone_Book.Text23.Text = ""
Phone_Book.Text24.Text = ""
Phone_Book.Text25.Text = ""
Phone_Book.Text26.Text = ""
Phone_Book.Text27.Text = ""
Phone_Book.Text28.Text = ""
Phone_Book.Text29.Text = ""
Phone_Book.Text30.Text = ""
Phone_Book.Text31.Text = ""
Phone_Book.Text32.Text = ""
Phone_Book.Text33.Text = ""
Phone_Book.Text34.Text = ""
Phone_Book.Text35.Text = ""
Phone_Book.Text36.Text = ""
End Select

  Case "exit"
  Dim response As Integer
  response = MsgBox("Esta seguro de salir de Pro Agenda?", vbYesNo + vbExclamation, "Pro Soft")
  If response = vbYes Then
  End
  End If
  End Select
 
End Sub
  
