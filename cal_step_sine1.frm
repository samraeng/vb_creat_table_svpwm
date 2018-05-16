VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form table 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   4980
   ClientLeft      =   8250
   ClientTop       =   2325
   ClientWidth     =   9045
   LinkTopic       =   "Form2"
   ScaleHeight     =   4980
   ScaleWidth      =   9045
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox num_of_stepdown 
      Height          =   495
      Left            =   7800
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox num_of_stepup 
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox stepdown_stop 
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox stepdown_start 
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT STEP DOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox stepup_stop 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox max_value 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT STEP UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TABLE NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NUM OF STEP"
      Height          =   255
      Left            =   7680
      TabIndex        =   15
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STOP"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "START"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NUM OF STEP"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STOP"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amplitude"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   6360
      X2              =   6600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   1200
      X2              =   1440
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "step down"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "step up"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim k As Long
Dim angle As Double
Dim revers(55) As Double
Dim angle2 As Double
Dim x As Double
Dim y
'Dim start_num As Long
'Dim stop_num As Long
Dim step As Double
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|data_inject (*.c)|*.c"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    CommonDialog1.ShowSave
    ' Display name of selected file

    '
    'MsgBox CommonDialog1.filename
    Open CommonDialog1.FileName For Output As #1
    
    angle = 0
    k = 0
    x = 0
    y = 0
   ' step = Val(num_of_stepup.Text)
    step = 90 / Val(num_of_stepup.Text)
   
    angle2 = (6.28 / 360) * angle
    k = Val(max_value.Text) * Sin(angle2)
    revers(y) = k
    Print #1, "int16 const time1[]={"
     Do While x <= Val(num_of_stepup.Text)
     
      Print #1, k & ","
      
      angle = angle + step
      angle2 = (6.28 / 360) * angle
      x = x + 1
      y = y + 1
      k = Val(max_value.Text) * Sin(angle2)
      revers(y) = k
      Loop
     
     Do While y >= 1
     
     Print #1, revers(y) & ","
     y = y - 1
     Loop
     
          Print #1, "}"
     Print #1, ";"
    Close #1
   
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub


End Sub

Private Sub Command2_Click()
Dim k As Long

Dim start_num As Long
Dim stop_num As Long
Dim step As Long
Dim n, m
If Text1.Text = "" Then
MsgBox ("please enter table number ")
Exit Sub
End If

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|data_inject (*.c)|*.c"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    CommonDialog1.ShowSave
    ' Display name of selected file

    '
    'MsgBox CommonDialog1.filename
    Open CommonDialog1.FileName For Output As #1
    
    k = Val(stepdown_start.Text)
    step = k - Val(stepdown_stop.Text)
    step = step / Val(num_of_stepdown.Text)
    n = Val(num_of_stepdown.Text)
    
    Print #1, "int16 const time" & Text1.Text & "[]={"
     Do While k >= Val(stepdown_stop.Text)
     
      Print #1, k & ","
      k = k - step
      m = k
      m = m + n
      n = n - 1
     Loop
    'Print #1, k & ","
     Print #1, "}"
     Print #1, ";"
    Close #1
   
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub


End Sub


