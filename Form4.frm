VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   2700
   ClientLeft      =   10950
   ClientTop       =   8970
   ClientWidth     =   4275
   LinkTopic       =   "Form4"
   ScaleHeight     =   2700
   ScaleWidth      =   4275
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3600
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "提交"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "验证码："
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "密码："
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim a As Long
Dim time As Long
Dim sfbcg As Long
Dim bczzmc As Long
Dim i As Long
Dim b As Long



Private Sub Command1_Click()

If Text2.Text = "114514" Then

Inet1.Protocol = icHTTP
    
       ' MsgBox (Inet1.OpenURL(Text1.Text))
a = Inet1.OpenURL("http://mc.lxyddice.top:16666/mima")

b = Text1.Text

Else
MsgBox "密码错误"
End
End If


If a = b Then
Form1.Show
Form4.Hide
Else
MsgBox "验证码错误"
End
End If





End Sub

Private Sub Command2_Click()
End
End Sub


