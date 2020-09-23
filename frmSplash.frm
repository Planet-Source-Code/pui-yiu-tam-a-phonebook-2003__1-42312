VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer 
      Left            =   4080
      Top             =   1080
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   720
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   1200
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   1680
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   2160
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   2640
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   3120
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   3600
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape shpArr 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   8
      Left            =   4080
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait a moment..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pro Soft"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1155
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmSplash"
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

Dim intIndex, intCounter As Integer
Dim boolValue As Boolean

Private Sub Form_Load()
    
frmSplash.Top = (Screen.Height - frmSplash.ScaleHeight) / 2
frmSplash.Left = (Screen.Width - frmSplash.ScaleWidth) / 2
shpArr(0).FillColor = vbRed
    
intIndex = 0
intCounter = 1
boolValue = False
tmrTimer.Interval = 100
    
End Sub

Private Sub tmrTimer_Timer()

Call mChangeColor(intIndex, &H8000000B)
   
If (boolValue = False) Then
intIndex = intIndex + 1
If (intIndex >= 8) Then
boolValue = True
End If
ElseIf (boolValue = True) Then
intIndex = intIndex - 1
If (intIndex <= 0) Then
boolValue = False
End If
End If
    
Call mChangeColor(intIndex, vbRed)
    
intCounter = intCounter + 1
If intCounter = 50 Then
Load MDIForm1
MDIForm1.Show
Unload Me
End If
    
End Sub

Private Sub mChangeColor(ByVal indx As Integer, ByVal Color As Long)
shpArr(indx).FillColor = Color
End Sub


