VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "prts敌人页面生成器-零下一度的冰晶beta1.2"
   ClientHeight    =   6795
   ClientLeft      =   6660
   ClientTop       =   5940
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   13470
   Begin VB.CommandButton Command22 
      Caption         =   "退出"
      Height          =   495
      Left            =   12840
      TabIndex        =   89
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "技能"
      Height          =   375
      Left            =   480
      TabIndex        =   87
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command21 
      Caption         =   "作者"
      Height          =   495
      Left            =   12840
      TabIndex        =   86
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "一键提交"
      Height          =   375
      Left            =   9960
      TabIndex        =   85
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton Command18 
      Caption         =   "直接提交"
      Height          =   375
      Left            =   4560
      TabIndex        =   83
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "清空"
      Height          =   375
      Left            =   5040
      TabIndex        =   80
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text30 
      Height          =   375
      Left            =   10800
      TabIndex        =   79
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text29 
      Height          =   375
      Left            =   10800
      TabIndex        =   78
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text28 
      Height          =   375
      Left            =   10800
      TabIndex        =   77
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text27 
      Height          =   375
      Left            =   10800
      TabIndex        =   76
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   10800
      TabIndex        =   75
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   10800
      TabIndex        =   69
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   7080
      TabIndex        =   67
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   7080
      TabIndex        =   65
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   8520
      TabIndex        =   63
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   6840
      TabIndex        =   61
      Text            =   "100"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   6840
      TabIndex        =   59
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   8520
      TabIndex        =   57
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text17 
      Height          =   390
      Left            =   8400
      TabIndex        =   55
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   6840
      TabIndex        =   54
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   6840
      TabIndex        =   51
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   7080
      TabIndex        =   49
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   7080
      TabIndex        =   47
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text13 
      Height          =   390
      Left            =   7080
      TabIndex        =   45
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command17 
      Caption         =   "加|index=0"
      Height          =   375
      Left            =   9720
      TabIndex        =   43
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      Caption         =   "加尾}}"
      Height          =   375
      Left            =   11160
      TabIndex        =   42
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      Caption         =   "加头{{敌人信息/level"
      Height          =   375
      Left            =   7320
      TabIndex        =   41
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command14 
      Caption         =   "加尾}}"
      Height          =   375
      Left            =   4080
      TabIndex        =   38
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "带颜色的文本"
      Height          =   375
      Left            =   480
      TabIndex        =   37
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "保存"
      Height          =   375
      Left            =   5040
      TabIndex        =   35
      Top             =   6240
      Width           =   855
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
      Height          =   1680
      Left            =   5880
      TabIndex        =   31
      Top             =   4920
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
      Left            =   5040
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   5040
      Width           =   3855
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
      Caption         =   "加头{{敌人信息/common"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label38 
      Caption         =   "vb特性，不要点X，否则会留下进程"
      Height          =   375
      Left            =   11280
      TabIndex        =   90
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label37 
      Caption         =   "警告：prts中预览不要点击保存更改（千万不要）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   12120
      TabIndex        =   88
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label36 
      Caption         =   "仅支持一键提交"
      Height          =   1335
      Left            =   9600
      TabIndex        =   84
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label35 
      Caption         =   "再做这边"
      Height          =   975
      Left            =   5880
      TabIndex        =   82
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label34 
      Caption         =   "先做这边"
      Height          =   2775
      Left            =   5040
      TabIndex        =   81
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label33 
      Caption         =   "浮空抗性"
      Height          =   375
      Left            =   9960
      TabIndex        =   74
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label32 
      Caption         =   "冻结抗性"
      Height          =   375
      Left            =   9960
      TabIndex        =   73
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label31 
      Caption         =   "沉睡抗性"
      Height          =   375
      Left            =   9960
      TabIndex        =   72
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label30 
      Caption         =   "沉默抗性"
      Height          =   375
      Left            =   9960
      TabIndex        =   71
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "眩晕抗性"
      Height          =   375
      Left            =   9960
      TabIndex        =   70
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label28 
      Caption         =   "重量等级"
      Height          =   375
      Left            =   9960
      TabIndex        =   68
      Top             =   720
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   9720
      X2              =   9720
      Y1              =   480
      Y2              =   5040
   End
   Begin VB.Label Label27 
      Caption         =   "sp恢复速度"
      Height          =   375
      Left            =   6000
      TabIndex        =   66
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label26 
      Caption         =   "生命恢复速度"
      Height          =   375
      Left            =   5880
      TabIndex        =   64
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label25 
      Caption         =   "攻击间隔"
      Height          =   375
      Left            =   7800
      TabIndex        =   62
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "攻击速度"
      Height          =   375
      Left            =   6000
      TabIndex        =   60
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "移动速度"
      Height          =   375
      Left            =   6000
      TabIndex        =   58
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "法术抗性"
      Height          =   255
      Left            =   7800
      TabIndex        =   56
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label21 
      Caption         =   "防御力"
      Height          =   375
      Left            =   6240
      TabIndex        =   53
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label20 
      Caption         =   "攻击力"
      Height          =   255
      Left            =   7800
      TabIndex        =   52
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "最大生命值"
      Height          =   255
      Left            =   5880
      TabIndex        =   50
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "攻击范围半径"
      Height          =   255
      Left            =   5880
      TabIndex        =   48
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "数量"
      Height          =   375
      Left            =   6600
      TabIndex        =   46
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "描述"
      Height          =   255
      Left            =   6600
      TabIndex        =   44
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "/level类型"
      Height          =   255
      Left            =   6000
      TabIndex        =   40
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   5760
      Y1              =   0
      Y2              =   5040
   End
   Begin VB.Label Label14 
      Caption         =   "/common类型"
      Height          =   495
      Left            =   120
      TabIndex        =   39
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "工具箱"
      Height          =   255
      Left            =   480
      TabIndex        =   36
      Top             =   5640
      Width           =   975
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

Private Sub Command12_Click()
Form2.Show

End Sub

Private Sub Command13_Click()
MsgBox "还没做"
End Sub

Private Sub Command14_Click()
List1.AddItem "}}"
End Sub

Private Sub Command15_Click()
List1.AddItem "{{敌人信息/level"
End Sub

Private Sub Command16_Click()
List1.AddItem "}}"
End Sub

Private Sub Command17_Click()
List1.AddItem "|index=0"
End Sub

Private Sub Command18_Click()
Text3.Text = "|名称=" & Text1.Text
List1.AddItem Text3.Text

Text3.Text = "|种类=" & Text2.Text
List1.AddItem Text3.Text

Text3.Text = "|index=" & Text4.Text
List1.AddItem Text3.Text

Text3.Text = "|地位级别=" & Text5.Text
List1.AddItem Text3.Text

Text3.Text = "|描述=" & Text6.Text
List1.AddItem Text3.Text

Text3.Text = "|攻击方式=" & Text7.Text
List1.AddItem Text3.Text

List1.AddItem "|耐久=" & Text8.Text
List1.AddItem "|攻击力=" & Text9.Text
List1.AddItem "|防御力=" & Text10.Text
List1.AddItem "|法术抗性=" & Text11.Text

Text3.Text = "|能力=" & Text12.Text
List1.AddItem Text3.Text

End Sub

Private Sub Command19_Click()
List1.Clear
End Sub

Private Sub Command2_Click()
Text3.Text = "|名称=" & Text1.Text
End Sub

Private Sub Command20_Click()


List1.AddItem "|描述=" & Text13.Text
List1.AddItem "|数量=" & Text14.Text
List1.AddItem "|攻击范围半径=" & Text15.Text
List1.AddItem "|最大生命值" & Text16.Text
List1.AddItem "|攻击力=" & Text17.Text
List1.AddItem "|防御力=" & Text18.Text
List1.AddItem "|法术抗性=" & Text19.Text
List1.AddItem "|移动速度=" & Text20.Text
List1.AddItem "|攻击速度" & Text21.Text
List1.AddItem "|攻击间隔" & Text22.Text
List1.AddItem "|生命恢复速度=" & Text23.Text
List1.AddItem "|sp恢复速度=" & Text24.Text
List1.AddItem "|重量等级=" & Text25.Text
List1.AddItem "|眩晕抗性=" & Text26.Text
List1.AddItem "|沉默抗性=" & Text27.Text
List1.AddItem "|沉睡抗性=" & Text28.Text
List1.AddItem "|冻结抗性=" & Text29.Text
List1.AddItem "|浮空抗性=" & Text30.Text

End Sub

Private Sub Command21_Click()
Form3.Show
End Sub

Private Sub Command22_Click()
End
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

Form1.Hide
Form4.Show

sfbcg = 0
bczzmc = Int(Rnd * (888888888 - 1 + 1)) + 1
End Sub

