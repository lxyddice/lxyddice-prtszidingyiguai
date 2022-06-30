VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "作者"
   ClientHeight    =   5085
   ClientLeft      =   11760
   ClientTop       =   7350
   ClientWidth     =   4365
   LinkTopic       =   "Form3"
   ScaleHeight     =   5085
   ScaleWidth      =   4365
   Begin VB.CommandButton Command2 
      Caption         =   "官网"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开源"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "版本beta1.2"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Github"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "哔哩哔哩个人空间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   3735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "explorer https://github.com/lxyddice/lxyddice-prtszidingyiguai"

End Sub

Private Sub Command2_Click()
MsgBox "搭建中..."
End Sub

Private Sub Label1_Click()
Shell "explorer https://space.bilibili.com/401299476"
End Sub

Private Sub Label2_Click()
Shell "explorer https://github.com/lxyddice"

End Sub

Private Sub Label3_Click()
MsgBox "哎嘿"
End Sub
