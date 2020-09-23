VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Mp3PlayerMain 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agenda 2003  MP3 Player"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5370
   FillColor       =   &H000080FF&
   ForeColor       =   &H00400000&
   Icon            =   "Mp3PlayerMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5370
   Begin VB.CommandButton Command7 
      Caption         =   "Remove Mp3"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Remvoe Mp3"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00800080&
      Caption         =   "Repeat Playlist"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Remove All Mp3's"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Remove All Mp3's"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800080&
      Caption         =   "Shutdown When Time is Reached"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   360
      Picture         =   "Mp3PlayerMain.frx":0BC2
      TabIndex        =   9
      ToolTipText     =   "If this is check, computer will shutdown when the selected time is reached"
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop Mp3"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Stop Mp3"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5025
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Welcome to Agenda 2003 MP3 Player"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6240
      Top             =   240
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Text            =   "Time to Stop Mp3"
      ToolTipText     =   "Type in time to stop Mp3"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Playlist"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save Playlist"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Playlist"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Load Playlist"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play Mp3"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Play Mp3"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00800080&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      ItemData        =   "Mp3PlayerMain.frx":1FDF
      Left            =   360
      List            =   "Mp3PlayerMain.frx":1FE1
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Playlist"
      Top             =   720
      Width           =   4650
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton command1 
      Caption         =   "Add Mp3"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Add Mp3"
      Top             =   3840
      Width           =   1455
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   645
      Left            =   360
      TabIndex        =   0
      Top             =   3120
      Width           =   4650
      AudioStream     =   -1
      AutoSize        =   -1  'True
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   -1  'True
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Mp3PlayerMain"
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

Private Type NOTIFYICONDATA 'this is to add a icon to the system tray...i got this from vbcode.com
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA


Dim x As String

Private Sub Command1_Click() 'Add Mp3
    cd.DialogTitle = "Open Mp3 File"
    cd.Filter = "MP3 Files|*.MP3"
    cd.InitDir = "C:\Windows\Desktop\"
    cd.ShowOpen
    List1.AddItem cd.FileName
End Sub
Private Sub Command2_Click() 'Play Mp3
MediaPlayer1.Open List1.Text
x = "   Playing Mp3"
StatusBar1.SimpleText = Time & x
If List1.Text = "" Then
MsgBox "Please choose a Mp3 to play", vbDefaultButton1, "Select Mp3"
x = "    Not Currently Playing A Mp3"
End If
Command5.Visible = True
Command2.Visible = False
End Sub
Private Sub Command3_Click() 'Load Playlist
On Error Resume Next
cd.DialogTitle = "Load Mp3 PlayList"
cd.Filter = "Mp3 PlayList|*.lst"
cd.InitDir = "c:\windows\desktop\"
cd.ShowOpen
List1.Clear
Call LoadList(List1, "C:\Windows\Desktop\Saved.lst")
End Sub
Sub LoadList(Lst As ListBox, File As String) 'For Loading a List
'Call LoadList (List1,"C:\Windows\Desktop\Saved.lst")
On Error Resume Next
Dim a As String
Open cd.FileName For Input As #1
Do Until EOF(1)
Input #1, a
Lst.AddItem a
Loop
Close 1
Exit Sub
End Sub
Private Sub Command4_Click() 'Save Playlist
On Error Resume Next
cd.DialogTitle = "Save Mp3 PlayList"
cd.Filter = "Mp3 PlayList|*.lst"
cd.InitDir = "c:\windows\desktop\"
cd.ShowSave
Call SaveList(List1, "C:\Windows\Desktop\Saved.lst")
End Sub
Sub SaveList(Lst As ListBox, File As String) 'for saving a List
'Call SaveList (List1,"C:\Windows\System\Saved.lst")
On Error Resume Next
Dim i As Variant
Dim a As Variant
Open cd.FileName For Output As #1
For i = 0 To Lst.ListCount - 1
a = Lst.List(i)
Print #1, a
Next
Close 1
Exit Sub
End Sub
Private Sub Command5_Click() 'stop Mp3
    MediaPlayer1.Stop
    x = "    Not Currently Playing A Mp3"
    Command5.Visible = False
    Command2.Visible = True
End Sub

Private Sub Command6_Click()  'Clear playlist
    List1.Clear
End Sub

Private Sub Command7_Click() 'Remove Selected Mp3
    If List1.ListIndex = -1 Then
MsgBox "No Mp3 Selected", vbExclamation, "Select Mp3"
Else
List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub Command8_Click()
    Mp3PlayerMain.WindowState = 1
    nid.cbSize = Len(nid)
   nid.hWnd = Mp3PlayerMain.hWnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = Mp3PlayerMain.Icon
   nid.szTip = "VisualScope Mp3 Player" & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove _
   (Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single)
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim msg As Long
    Dim sFilter As String
    msg = x / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
       Mp3PlayerMain.WindowState = 0
       Shell_NotifyIcon NIM_DELETE, nid
       Case WM_RBUTTONDOWN
       MsgBox "Thank You For Using VisualScope Mp3 Player.  We Hope You Enjoy Using Our Product and Continue To Use The Excellent Software From VisualScope.", vbDefaultButton1, "Thank You"
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub Form_Load()
x = "    Not Currently Playing A Mp3"
   StatusBar1.Style = sbrSimple
   StatusBar1.SimpleText = Time & x
  
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
    Unload Me
   
End Sub

Private Sub List1_dblClick() 'Play Selected Mp3 on Double Click
   MediaPlayer1.FileName = List1.Text
   MediaPlayer1.Play
   x = "   Playing Mp3"
   Command5.Visible = True
   Command2.Visible = False
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long) 'Go on to next Mp3 when done with Current one
  On Error Resume Next
 If List1.ListIndex = List1.ListCount - 1 Then
 If Check2.Value = 1 Then
 List1.ListIndex = 0
 MediaPlayer1.FileName = List1.Text
 MediaPlayer1.Play
 Else
 List1.ListIndex = 0
 End If
 Else
 List1.ListIndex = List1.ListIndex + 1
 MediaPlayer1.FileName = List1.Text
 MediaPlayer1.Play
 End If
End Sub

Private Sub timer1_timer()
StatusBar1.SimpleText = Time & x
    If Text1.Text = Time Then
MediaPlayer1.Stop 'stops mp3
x = "    Not Currently Playing A Mp3"
Command5.Visible = False
Command2.Visible = True
If Check1.Value = 1 Then
Shell "rundll32 krnl386.exe,exitkernel" 'shuts down the computer
End If
End If
End Sub
