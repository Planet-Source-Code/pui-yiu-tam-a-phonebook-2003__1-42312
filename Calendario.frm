VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Calendario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendario de Actividades."
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Calendar Activities"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.Frame Frame4 
         Caption         =   "Calendario"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4335
         Begin MSACAL.Calendar Calendar1 
            Height          =   2295
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   3975
            _Version        =   524288
            _ExtentX        =   7011
            _ExtentY        =   4048
            _StockProps     =   1
            BackColor       =   -2147483635
            Year            =   2003
            Month           =   1
            Day             =   3
            DayLength       =   1
            MonthLength     =   2
            DayFontColor    =   0
            FirstDay        =   2
            GridCellEffect  =   1
            GridFontColor   =   10485760
            GridLinesColor  =   -2147483632
            ShowDateSelectors=   -1  'True
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   -1  'True
            ShowTitle       =   -1  'True
            ShowVerticalGrid=   -1  'True
            TitleFontColor  =   10485760
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   0
            Top             =   3360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Listado de Actividades Pendientes"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   9735
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2535
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4471
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Actividad"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   5295
         Begin VB.CommandButton Command1 
            Caption         =   "Asignar"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   14
            Top             =   2280
            Width           =   1215
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
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   3375
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
            Height          =   525
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   1680
            Width           =   4935
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   720
            Width           =   3375
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
            Left            =   1680
            TabIndex        =   1
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hora a Realizar"
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
            Left            =   480
            TabIndex        =   13
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Texto de Recordatorio"
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
            TabIndex        =   12
            Top             =   1440
            Width           =   1665
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha  a Realizar"
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
            TabIndex        =   11
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Actividad a Realizar"
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
            TabIndex        =   10
            Top             =   360
            Width           =   1515
         End
      End
   End
End
Attribute VB_Name = "Calendario"
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

Private Sub Calendar1_DblClick()
Calendario.Text2 = Calendar1.Value
End Sub

Private Sub Command1_Click()
Load Alarma
Alarma.Show
Unload Me
End Sub

Private Sub Form_Load()
Calendar1 = Date
Text4.Text = Time
End Sub


