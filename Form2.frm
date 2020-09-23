VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Phone_Book 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Phone Book 2003."
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9315
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   6360
      Top             =   360
   End
   Begin VB.Frame Frame1 
      Caption         =   "Phone Book!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6600
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4935
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8705
         _Version        =   393216
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Informacion Principal"
         TabPicture(0)   =   "Form2.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Informacion Secundaria"
         TabPicture(1)   =   "Form2.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Informacion Personal"
         TabPicture(2)   =   "Form2.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame8"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame8 
            Height          =   4455
            Left            =   -74880
            TabIndex        =   71
            Top             =   360
            Width           =   8655
            Begin VB.Frame Frame11 
               Caption         =   "Informacion Familiar"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   240
               TabIndex        =   82
               Top             =   120
               Width           =   8295
               Begin VB.TextBox Text25 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   24
                  Top             =   360
                  Width           =   5175
               End
               Begin VB.TextBox Text26 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   25
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.TextBox Text27 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   26
                  Top             =   1080
                  Width           =   2895
               End
               Begin VB.TextBox Text28 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   27
                  Top             =   1440
                  Width           =   2895
               End
               Begin VB.Label Label39 
                  AutoSize        =   -1  'True
                  Caption         =   "Wife Name"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1080
                  TabIndex        =   86
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label Label38 
                  AutoSize        =   -1  'True
                  Caption         =   "Birthday Wife"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   840
                  TabIndex        =   85
                  Top             =   720
                  Width           =   1035
               End
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "Cell Wife"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1200
                  TabIndex        =   84
                  Top             =   1080
                  Width           =   705
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone work Wife"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   600
                  TabIndex        =   83
                  Top             =   1440
                  Width           =   1305
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "Numeros de Telefono de Emergencia"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   240
               TabIndex        =   77
               Top             =   1920
               Width           =   8295
               Begin VB.TextBox Text34 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   29
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.TextBox Text33 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   28
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.TextBox Text32 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   30
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.TextBox Text30 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   31
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone 2"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4080
                  TabIndex        =   81
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label33 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone 1"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   360
                  TabIndex        =   80
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label32 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone 3"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   360
                  TabIndex        =   79
                  Top             =   720
                  Width           =   615
               End
               Begin VB.Label Label30 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone 4"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4080
                  TabIndex        =   78
                  Top             =   720
                  Width           =   615
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "Telefono Familiares"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   240
               TabIndex        =   72
               Top             =   3120
               Width           =   8295
               Begin VB.TextBox Text29 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   35
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.TextBox Text31 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   34
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.TextBox Text35 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   32
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.TextBox Text36 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   33
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.Label Label29 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone 4"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4080
                  TabIndex        =   76
                  Top             =   720
                  Width           =   615
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone 3"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   360
                  TabIndex        =   75
                  Top             =   720
                  Width           =   615
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone 1"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   360
                  TabIndex        =   74
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label36 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone 2"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4080
                  TabIndex        =   73
                  Top             =   360
                  Width           =   615
               End
            End
         End
         Begin VB.Frame Frame4 
            Height          =   4455
            Left            =   -74880
            TabIndex        =   54
            Top             =   360
            Width           =   8535
            Begin VB.Frame Frame5 
               Caption         =   "Direcciones Electronicas"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   120
               TabIndex        =   65
               Top             =   120
               Width           =   8295
               Begin VB.TextBox Text14 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   15
                  Top             =   240
                  Width           =   6015
               End
               Begin VB.TextBox Text15 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   16
                  Top             =   600
                  Width           =   6015
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "E-Mail"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   960
                  TabIndex        =   67
                  Top             =   240
                  Width           =   510
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Web Site"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   840
                  TabIndex        =   66
                  Top             =   600
                  Width           =   690
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Informacion del Trabajo"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3255
               Left            =   120
               TabIndex        =   55
               Top             =   1080
               Width           =   8295
               Begin VB.TextBox Text16 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   17
                  Top             =   360
                  Width           =   4215
               End
               Begin VB.TextBox Text17 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   56
                  Top             =   720
                  Width           =   6015
               End
               Begin VB.TextBox Text18 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   18
                  Top             =   1080
                  Width           =   2895
               End
               Begin VB.TextBox Text19 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   19
                  Top             =   1440
                  Width           =   2895
               End
               Begin VB.TextBox Text20 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   22
                  Top             =   2520
                  Width           =   2895
               End
               Begin VB.TextBox Text21 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   20
                  Top             =   1800
                  Width           =   2895
               End
               Begin VB.TextBox Text22 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   21
                  Top             =   2160
                  Width           =   2895
               End
               Begin VB.TextBox Text23 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   1725
                  Left            =   4800
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   23
                  Top             =   1320
                  Width           =   2895
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Company"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   840
                  TabIndex        =   64
                  Top             =   360
                  Width           =   690
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Adreess Company"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   120
                  TabIndex        =   63
                  Top             =   720
                  Width           =   1350
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  Caption         =   "State"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1080
                  TabIndex        =   62
                  Top             =   1080
                  Width           =   405
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  Caption         =   "Zip Code"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   840
                  TabIndex        =   61
                  Top             =   1440
                  Width           =   675
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  Caption         =   "Country"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   960
                  TabIndex        =   60
                  Top             =   2520
                  Width           =   585
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  Caption         =   "E-Mail"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1080
                  TabIndex        =   59
                  Top             =   1800
                  Width           =   510
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  Caption         =   "Web Site"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   840
                  TabIndex        =   58
                  Top             =   2160
                  Width           =   690
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  Caption         =   "Comentarios"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4800
                  TabIndex        =   57
                  Top             =   1080
                  Width           =   945
               End
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4335
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   8655
            Begin VB.Frame Frame3 
               Caption         =   "Civic Address"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1935
               Left            =   120
               TabIndex        =   46
               Top             =   120
               Width           =   8415
               Begin VB.TextBox Text24 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   7
                  Top             =   1440
                  Width           =   2895
               End
               Begin VB.TextBox Text6 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   6
                  Top             =   1440
                  Width           =   2895
               End
               Begin VB.TextBox Text5 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   5
                  Top             =   1080
                  Width           =   2895
               End
               Begin VB.TextBox Text4 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   4
                  Top             =   1080
                  Width           =   2895
               End
               Begin VB.TextBox Text3 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   3
                  Top             =   720
                  Width           =   6735
               End
               Begin VB.TextBox Text1 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   1
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.TextBox Text2 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   2
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  Caption         =   "Birthday"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4200
                  TabIndex        =   53
                  Top             =   1440
                  Width           =   630
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Zip Code"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   120
                  TabIndex        =   52
                  Top             =   1440
                  Width           =   675
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Country"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4200
                  TabIndex        =   51
                  Top             =   1080
                  Width           =   585
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "State"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   360
                  TabIndex        =   50
                  Top             =   1080
                  Width           =   405
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Adreess"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   240
                  TabIndex        =   49
                  Top             =   720
                  Width           =   615
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Last Name"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   120
                  TabIndex        =   48
                  Top             =   360
                  Width           =   825
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "First Name"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4080
                  TabIndex        =   47
                  Top             =   360
                  Width           =   825
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "Numeros de Telefono"
               BeginProperty Font 
                  Name            =   "Palatino Linotype"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1935
               Left            =   120
               TabIndex        =   38
               Top             =   2160
               Width           =   8415
               Begin VB.TextBox Text13 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   13
                  Top             =   1080
                  Width           =   2895
               End
               Begin VB.TextBox Text12 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   9
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.TextBox Text11 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   8
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.TextBox Text10 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   10
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.TextBox Text9 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   12
                  Top             =   1080
                  Width           =   2895
               End
               Begin VB.TextBox Text8 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4920
                  TabIndex        =   11
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.TextBox Text7 
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   14
                  Top             =   1440
                  Width           =   2895
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Work 3"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4200
                  TabIndex        =   45
                  Top             =   1080
                  Width           =   555
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Work 1"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4200
                  TabIndex        =   44
                  Top             =   360
                  Width           =   555
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   360
                  TabIndex        =   43
                  Top             =   360
                  Width           =   480
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Fax"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   600
                  TabIndex        =   42
                  Top             =   720
                  Width           =   255
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Celular"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   360
                  TabIndex        =   41
                  Top             =   1080
                  Width           =   540
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Work 2"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   4200
                  TabIndex        =   40
                  Top             =   720
                  Width           =   555
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Zip Code"
                  BeginProperty Font 
                     Name            =   "Palatino Linotype"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   240
                  TabIndex        =   39
                  Top             =   1440
                  Width           =   675
               End
            End
         End
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Print Phone Book"
         Height          =   195
         Left            =   4800
         TabIndex        =   70
         Top             =   5520
         Width           =   1245
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Form2.frx":0054
         Top             =   5400
         Width           =   480
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Send by Email"
         Height          =   195
         Left            =   840
         TabIndex        =   69
         Top             =   5520
         Width           =   1005
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2160
         Picture         =   "Form2.frx":035E
         Top             =   5400
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Refresh Contact"
         Height          =   195
         Left            =   2760
         TabIndex        =   68
         Top             =   5520
         Width           =   1155
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   4200
         Picture         =   "Form2.frx":0C28
         Top             =   5400
         Width           =   480
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Add 
         Caption         =   "&Add Contact"
         Shortcut        =   ^A
      End
      Begin VB.Menu Remove 
         Caption         =   "&Remove Contact"
         Shortcut        =   ^X
      End
      Begin VB.Menu Search 
         Caption         =   "&Search Contact"
         Shortcut        =   ^F
      End
      Begin VB.Menu Clear 
         Caption         =   "Clear List"
      End
      Begin VB.Menu Imprimir_Phone 
         Caption         =   "Print Phone Book"
      End
      Begin VB.Menu Line6 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit Phone Book"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Font_Colors 
         Caption         =   "Font Colors"
      End
   End
End
Attribute VB_Name = "Phone_Book"
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

Private Sub Add_Click()
With rs
.AddNew
!Name1 = Text1.Text
!name2 = Text2.Text
!direccion1 = Text3.Text
!ciudad1 = Text4.Text
!pais1 = Text5.Text
!zipostal1 = Text6.Text
!cumpleaños = Text24.Text
!personal = Text11.Text
!fax1 = Text10.Text
!celular1 = Text9.Text
!zipostal2 = Text7.Text
!trabajo1 = Text12.Text
!trabajo2 = Text8.Text
!trabajo3 = Text13.Text
!email1 = Text14.Text
!website1 = Text15.Text
!compania = Text16.Text
!direccion2 = Text17.Text
!ciudad2 = Text18.Text
!zipostal3 = Text19.Text
!email2 = Text21.Text
!website2 = Text22.Text
!comentarios1 = Text23.Text
!pais2 = Text20.Text
!esposa1 = Text25.Text
!esposa2 = Text26.Text
!esposa3 = Text27.Text
!esposa4 = Text28.Text
!tel1 = Text29.Text
!tel2 = Text30.Text
!tel3 = Text31.Text
!tel4 = Text32.Text
!tel5 = Text33.Text
!tel6 = Text34.Text
!tel7 = Text35.Text
!tel8 = Text36.Text
.Update
.Close
MsgBox "Your Record has been added", vbOKOnly + vbInformation, "Pro Agenda 2003"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text25.Text = ""
Text26.Text = ""
Text27.Text = ""
Text28.Text = ""
Text29.Text = ""
Text30.Text = ""
Text31.Text = ""
Text32.Text = ""
Text33.Text = ""
Text34.Text = ""
Text35.Text = ""
Text36.Text = ""
End With
Call Form_Load
End Sub

Private Sub Clear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text25.Text = ""
Text26.Text = ""
Text27.Text = ""
Text28.Text = ""
Text29.Text = ""
Text30.Text = ""
Text31.Text = ""
Text32.Text = ""
Text33.Text = ""
Text34.Text = ""
Text35.Text = ""
Text36.Text = ""
End Sub

Private Sub Exit_Click()
MDIForm1.Toolbar1.Buttons.Item(4).Visible = False
Unload Me
End Sub

Private Sub Font_Colors_Click()
CommonDialog1.Flags = &H1&
CommonDialog1.ShowColor
Text1.ForeColor = CommonDialog1.Color
Text2.ForeColor = CommonDialog1.Color
Text3.ForeColor = CommonDialog1.Color
Text4.ForeColor = CommonDialog1.Color
Text5.ForeColor = CommonDialog1.Color
Text6.ForeColor = CommonDialog1.Color
Text7.ForeColor = CommonDialog1.Color
Text8.ForeColor = CommonDialog1.Color
Text9.ForeColor = CommonDialog1.Color
Text10.ForeColor = CommonDialog1.Color
Text11.ForeColor = CommonDialog1.Color
Text12.ForeColor = CommonDialog1.Color
Text13.ForeColor = CommonDialog1.Color
Text14.ForeColor = CommonDialog1.Color
Text15.ForeColor = CommonDialog1.Color
Text16.ForeColor = CommonDialog1.Color
Text17.ForeColor = CommonDialog1.Color
Text18.ForeColor = CommonDialog1.Color
Text19.ForeColor = CommonDialog1.Color
Text20.ForeColor = CommonDialog1.Color
Text21.ForeColor = CommonDialog1.Color
Text22.ForeColor = CommonDialog1.Color
Text23.ForeColor = CommonDialog1.Color
Text24.ForeColor = CommonDialog1.Color
Text25.ForeColor = CommonDialog1.Color
Text26.ForeColor = CommonDialog1.Color
Text27.ForeColor = CommonDialog1.Color
Text28.ForeColor = CommonDialog1.Color
Text29.ForeColor = CommonDialog1.Color
Text30.ForeColor = CommonDialog1.Color
Text31.ForeColor = CommonDialog1.Color
Text32.ForeColor = CommonDialog1.Color
Text33.ForeColor = CommonDialog1.Color
Text34.ForeColor = CommonDialog1.Color
Text35.ForeColor = CommonDialog1.Color
Text36.ForeColor = CommonDialog1.Color
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\agenda.mdb")
With db
Set rs = .OpenRecordset("agenda")
End With
End Sub



Private Sub Image3_Click()
Beep
On Error Resume Next
Dim Answer As String
Answer = MsgBox("Confirm printing on  " & _
Printer.DeviceName, vbYesNo, "print ... ?")
If Answer = vbNo Then Exit Sub
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print Text1.Text
Printer.Print Text2.Text
Printer.Print Text3.Text
Printer.Print Text4.Text
Printer.Print Text5.Text
Printer.Print Text6.Text
Printer.Print Text7.Text
Printer.Print Text8.Text
Printer.Print Text9.Text
Printer.Print Text10.Text
Printer.Print Text11.Text
Printer.Print Text12.Text
Printer.Print Text13.Text
Printer.Print Text14.Text
Printer.Print Text15.Text
Printer.Print Text16.Text
Printer.Print Text17.Text
Printer.Print Text18.Text
Printer.Print Text19.Text
Printer.Print Text20.Text
Printer.Print Text21.Text
Printer.Print Text22.Text
Printer.Print Text23.Text
Printer.Print Text24.Text
Printer.Print Text25.Text
Printer.Print Text26.Text
Printer.Print Text27.Text
Printer.Print Text28.Text
Printer.Print Text29.Text
Printer.Print Text30.Text
Printer.Print Text31.Text
Printer.Print Text32.Text
Printer.Print Text33.Text
Printer.Print Text34.Text
Printer.Print Text35.Text
Printer.Print Text36.Text
Printer.EndDoc
End Sub



Private Sub Remove_Click()
On Error GoTo handle
Set rs = db.OpenRecordset("SELECT * FROM agenda WHERE Name1='" & Text1.Text & "'", dbOpenDynaset)
With rs
.Delete
.Close
End With
Call Form_Load

handle:
Select Case Err.Number
Case 3021
MsgBox "There is no such a record", vbOKOnly + vbInformation, "Pro Agenda 2003"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text25.Text = ""
Text26.Text = ""
Text27.Text = ""
Text28.Text = ""
Text29.Text = ""
Text30.Text = ""
Text31.Text = ""
Text32.Text = ""
Text33.Text = ""
Text34.Text = ""
Text35.Text = ""
Text36.Text = ""
End Select
End Sub

Private Sub Search_Click()
On Error GoTo handle
Set rs = db.OpenRecordset("SELECT * FROM agenda WHERE name1='" & Text1.Text & "'", dbOpenDynaset)
Text2.Text = rs!name2
Text3.Text = rs!direccion1
Text4.Text = rs!ciudad1
Text5.Text = rs!pais1
Text6.Text = rs!zipostal1
Text7.Text = rs!zipostal2
Text8.Text = rs!trabajo2
Text9.Text = rs!celular1
Text10.Text = rs!fax1
Text11.Text = rs!personal
Text12.Text = rs!trabajo1
Text13.Text = rs!trabajo3
Text14.Text = rs!email1
Text15.Text = rs!website1
Text16.Text = rs!compania
Text17.Text = rs!direccion2
Text18.Text = rs!ciudad2
Text19.Text = rs!zipostal3
Text20.Text = rs!pais2
Text21.Text = rs!email2
Text22.Text = rs!website2
Text23.Text = rs!comentarios1
Text24.Text = rs!cumpleaños
Text25.Text = rs!esposa1
Text26.Text = rs!esposa2
Text27.Text = rs!esposa3
Text28.Text = rs!esposa4
Text29.Text = rs!tel1
Text30.Text = rs!tel2
Text31.Text = rs!tel3
Text32.Text = rs!tel4
Text33.Text = rs!tel5
Text34.Text = rs!tel6
Text35.Text = rs!tel7
Text36.Text = rs!tel8
rs.Close
Call Form_Load

handle:
Select Case Err.Number
Case 3021
MsgBox "There is no such a record", vbOKOnly + vbInformation, "Pro Agenda 2003"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text25.Text = ""
Text26.Text = ""
Text27.Text = ""
Text28.Text = ""
Text29.Text = ""
Text30.Text = ""
Text31.Text = ""
Text32.Text = ""
Text33.Text = ""
Text34.Text = ""
Text35.Text = ""
Text36.Text = ""
End Select

End Sub



