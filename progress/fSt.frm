VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "StatusBArs"
   ClientHeight    =   6075
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7170
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   5400
      Width           =   1455
   End
   Begin Project1.drawfield drawfield1 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   2325
      _extentx        =   0
      _extenty        =   0
      begincolor      =   16777215
      endcolor        =   0
      value           =   1
      boxcount        =   40
      boxspace        =   0
   End
   Begin Project1.drawfield drawfield2 
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   5475
      _extentx        =   9657
      _extenty        =   344
      begincolor      =   12632319
      endcolor        =   33023
      value           =   1
      boxcount        =   40
   End
   Begin Project1.drawfield drawfield3 
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   6075
      _extentx        =   10716
      _extenty        =   344
      begincolor      =   8454143
      endcolor        =   49152
      value           =   1
      boxcount        =   40
      boxspace        =   3
   End
   Begin Project1.drawfield drawfield4 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   4575
      _extentx        =   8070
      _extenty        =   344
      begincolor      =   16776960
      endcolor        =   128
      value           =   1
   End
   Begin Project1.drawfield drawfield5 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   5925
      _extentx        =   10451
      _extenty        =   344
      begincolor      =   12582912
      endcolor        =   16761087
      value           =   1
      boxspace        =   4
   End
   Begin Project1.drawfield drawfield6 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   5925
      _extentx        =   10451
      _extenty        =   344
      begincolor      =   8438015
      endcolor        =   33023
      value           =   1
   End
   Begin Project1.drawfield drawfield7 
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   6285
      _extentx        =   0
      _extenty        =   0
      endcolor        =   16711935
      value           =   1
      boxspace        =   0
   End
   Begin Project1.drawfield drawfield8 
      Height          =   1695
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   5475
      _extentx        =   9657
      _extenty        =   344
      begincolor      =   0
      endcolor        =   8421504
      value           =   1
      boxspace        =   5
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   5520
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Kozari Laszlo in 2004 aug.7
'VOTE on PSC if u like it, tenx
Dim b As Boolean
Dim i As Integer
Private Sub Command1_Click()

Timer1.Interval = 5

End Sub

Private Sub Form_Load()
b = True
i = 0
End Sub

Private Sub Timer1_Timer()

If b Then i = i + 1
If Not b Then i = i - 1

If i > 100 Then i = 100: b = False
If i < 0 Then i = 0: b = True

drawfield1.Value = i
drawfield2.Value = i
drawfield3.Value = i
drawfield4.Value = i
drawfield5.Value = i
drawfield6.Value = i
drawfield7.Value = i
drawfield8.Value = i

Label1.Caption = CStr(i) + " %"

End Sub

