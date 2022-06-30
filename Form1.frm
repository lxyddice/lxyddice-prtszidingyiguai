VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   6660
   ClientTop       =   5940
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   13470
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   375
      Left            =   4800
      TabIndex        =   35
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "插入"
      Height          =   375
      Left            =   3600
      TabIndex        =   34
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   1200
      TabIndex        =   33
      Top             =   4440
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   5880
      TabIndex        =   31
      Top             =   5640
      Width           =   6975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "直接插入"
      Height          =   615
      Left            =   3600
      TabIndex        =   30
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   2640
      TabIndex        =   29
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   390
      Left            =   1200
      TabIndex        =   28
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   390
      Left            =   2640
      TabIndex        =   27
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   390
      Left            =   1200
      TabIndex        =   26
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "插入"
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "插入"
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "插入"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "插入"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "提交"
      Height          =   375
      Left            =   11760
      TabIndex        =   9
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   5040
      Width           =   10575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "插入"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "插入"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加入头{{敌人信息/common"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label12 
      Caption         =   "能力"
      Height          =   375
      Left            =   720
      TabIndex        =   32
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "法抗"
      Height          =   375
      Left            =   2280
      TabIndex        =   25
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "防御"
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "攻击"
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "耐久"
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "攻击方式"
      Height          =   375
      Left            =   720
      TabIndex        =   19
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "描述"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "地位级别"
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "序号"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "最终编辑框"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "种类"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "名称"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sfgcg As Long
Dim bczzmc As Long
Private Sub Command1_Click()
List1.AddItem "{{敌人信息/common"
End Sub

Private Sub Command10_Click()
Text3.Text = "|能力=" & Text12.Text
End Sub

Private Sub Command11_Click()
If sfbcg > 0 Then

Kill App.Path & "\" & bczzmc & ".txt"
End If
sfbcg = sfbcg + 1
Dim fn As Integer, i As Integer
For i = 0 To List1.ListCount - 1
Open App.Path & "\" & bczzmc & ".txt" For Append As #1
Print #1, List1.List(i)
Close #1
Next
Close fn
End Sub

Private Sub Command2_Click()
Text3.Text = "|名称=" & Text1.Text
End Sub

Private Sub Command3_Click()
Text3.Text = "|种类=" & Text2.Text
End Sub

Private Sub Command4_Click()
List1.AddItem Text3.Text
End Sub

Private Sub Command5_Click()
Text3.Text = "|index=" & Text4.Text

End Sub

Private Sub Command6_Click()
Text3.Text = "|地位级别=" & Text5.Text
End Sub

Private Sub Command7_Click()
Text3.Text = "|描述=" & Text6.Text

End Sub

Private Sub Command8_Click()
Text3.Text = "|攻击方式=" & Text7.Text
End Sub

Private Sub Command9_Click()

List1.AddItem "|耐久=" & Text8.Text
List1.AddItem "|攻击力=" & Text9.Text
List1.AddItem "|防御力=" & Text10.Text
List1.AddItem "|法术抗性=" & Text11.Text

End Sub

Private Sub Form_Load()
sfbcg = 0
bczzmc = Int(Rnd * (888888888 - 1 + 1)) + 1
End Sub
