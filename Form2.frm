VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4485
   ClientLeft      =   9105
   ClientTop       =   6540
   ClientWidth     =   9285
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   9285
   Begin VB.CommandButton Command6 
      Caption         =   "加尾}}"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "加头{{"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "误删恢复"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清空"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   2640
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "插入"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "插入"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "#00FFFF"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "一般不用"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "输出"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "彩色文本"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "普通文本"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "颜色"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wb As String

Private Sub Command1_Click()
If Text4.Text = "" Then
wb = Text2.Text
Text4.Text = wb
Else
wb = Text4.Text
wb = wb & Text2.Text
Text4.Text = wb
End If
End Sub

Private Sub Command2_Click()
If Text4.Text = "" Then
wb = "{{color|" & Text1.Text & "|" & Text3.Text & "}}"
Text4.Text = wb
Else
wb = Text4.Text
wb = wb & "{{color|" & Text1.Text & "|" & Text3.Text & "}}"
Text4.Text = wb
End If

End Sub

Private Sub Command3_Click()
Text4.Text = ""
End Sub

Private Sub Command4_Click()
Text4.Text = wb
End Sub

Private Sub Command5_Click()
wb = Text4.Text
wb = wb & "{{"
Text4.Text = wb
End Sub

Private Sub Command6_Click()
wb = Text4.Text
wb = wb & "}}"
Text4.Text = wb
End Sub

