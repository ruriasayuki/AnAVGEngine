VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "胡泽豪的AVG引擎 test 简陋的欢迎界面！！"
   ClientHeight    =   5400
   ClientLeft      =   5490
   ClientTop       =   3645
   ClientWidth     =   8100
   LinkTopic       =   "Form2"
   ScaleHeight     =   5400
   ScaleWidth      =   8100
   Begin VB.CommandButton Command2 
      Caption         =   "结束游戏"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始游戏"
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Hide
    Form1.Form_Load
    Form1.Show
End Sub

Private Sub Command2_Click()
 End
End Sub

Private Sub Form_Load()
    Form2.Caption = "胡泽豪的AVG引擎"
    Command1.Caption = "开始游戏"
    Command2.Caption = "结束游戏"
End Sub
