VERSION 5.00
Begin VB.Form Alarma 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alarma de Actividades."
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Alarma"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Frame Frame11 
         Caption         =   "Alarm Function"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   3855
         Begin VB.TextBox Alarmtimebox 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   9
            Top             =   600
            Width           =   1575
         End
         Begin VB.Timer Timer4 
            Interval        =   1
            Left            =   2760
            Top             =   4080
         End
         Begin VB.Timer Timer3 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   2400
            Top             =   4080
         End
         Begin VB.Timer Timer2 
            Interval        =   1
            Left            =   2040
            Top             =   4080
         End
         Begin VB.TextBox Datebox 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox Alarmdatebox 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   7
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton Stop_button 
            Caption         =   "Stop"
            Height          =   375
            Left            =   1920
            TabIndex        =   6
            Top             =   3600
            Width           =   1575
         End
         Begin VB.CommandButton Start_button 
            Caption         =   "Start"
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   3600
            Width           =   1575
         End
         Begin VB.TextBox Alarmmessagebox 
            Height          =   1095
            Left            =   120
            MaxLength       =   256
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   2400
            Width           =   3495
         End
         Begin VB.TextBox Alarmcaptionbox 
            Height          =   285
            Left            =   120
            MaxLength       =   32
            TabIndex        =   3
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Timer Timer5 
            Interval        =   1
            Left            =   1680
            Top             =   4080
         End
         Begin VB.TextBox Timebox 
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "Your alarm time:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label38 
            Caption         =   "Date now:"
            Height          =   255
            Left            =   2040
            TabIndex        =   14
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label39 
            Caption         =   "Your alarm date:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label40 
            Caption         =   "Your alarm message:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label41 
            Caption         =   "Your alarm caption:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label42 
            Caption         =   "Time now:"
            Height          =   255
            Left            =   2040
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4440
         Top             =   600
      End
      Begin VB.Label Onoroff 
         Caption         =   "Alarm: off"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu Add 
         Caption         =   "&Add Contact"
      End
      Begin VB.Menu Remove 
         Caption         =   "&Remove Contact"
      End
      Begin VB.Menu Search 
         Caption         =   "&Search Contact"
      End
      Begin VB.Menu Clear 
         Caption         =   "&Clear List"
      End
      Begin VB.Menu Line5 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit Phone Book"
      End
   End
End
Attribute VB_Name = "Alarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'This Code was written by Javier Castellanos Pro Soft programmer and You Cannot Delete This If
'You Are Going To Use This Code...But Have Fun And I Hope You Enjoy
'Javier Castellanos Pro Soft programmer
'If you find any bugs,
'Please E-mail me a personalgjoc@hotmail.com So
'I can get them fixed...
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

Private Sub Exit_Click()
MDIForm1.Toolbar1.Buttons.Item(6).Visible = False
Unload Me
End Sub

Private Sub Form_Load()
Alarma.Hide
Timer3.Enabled = True
Alarma.Alarmcaptionbox.Text = Calendario.Text1.Text
Alarma.Alarmdatebox.Text = Calendario.Text2.Text
Alarma.Alarmmessagebox.Text = Calendario.Text3.Text
Alarma.Alarmtimebox.Text = Calendario.Text4.Text
Unload Calendario
Alarma.WindowState = 1
End Sub

Private Sub Start_button_Click()
  Timer3.Enabled = True
End Sub

Private Sub Stop_button_Click()
  Timer3.Enabled = False
End Sub

Private Sub timer1_timer()
 Timebox.Text = Time
End Sub

Private Sub Timer2_Timer()
  Datebox.Text = Date
  ' Shows current date
End Sub

Private Sub Timer3_Timer()
  If Alarmtimebox.Text = Timebox.Text And Alarmdatebox.Text = Datebox.Text Then Call Alarm
  ' If Timer3 is on and
  ' Alarmtimebox and Alarmdatebox have same text than Timebox and Datebox have
  ' then call function Alarm
End Sub

Function Alarm()
  Timer3.Enabled = False
  Dim Caption
  Dim Message
  Caption = Alarmcaptionbox.Text
  Message = Alarmmessagebox.Text
  MsgBox Message, vbInformation + vbSystemModal, Caption
  ' Turns Timer3 off and creates MsgBox
  ' using Alarmcaptionbox text and Alarmmessagebox text
Unload Me
End Function

Private Sub Timer4_Timer()
  If Timer3.Enabled = False Then Onoroff.Caption = "Alarm: off"
  If Timer3.Enabled = True Then Onoroff.Caption = "Alarm: on"
  ' If Timer3 is off then Onoroff caption is "Alarm: off"
  ' and if Timer3 is on then Onoroff caption is "Alarm: on"
End Sub

