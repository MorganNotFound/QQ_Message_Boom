VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "QQ消息发送器"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1200
      TabIndex        =   11
      Text            =   "C:\Users\Administrator\Desktop\编程软件\Works\VisualBasic 6.0\QQ消息轰炸终版\boom.VBS"
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2760
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "结束"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "boom.VBS地址："
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   " 毫秒"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "间隔："
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "发送数量："
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "内容："
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "输入要轰炸的QQ号："
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m As Integer
Dim n As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
Dim VBS As String
a = Text2.Text
m = Text3.Text
n = Text4.Text
Shell Environ("PROGRAMFILES") & "\Internet Explorer\iexplore.exe " & "Tencent://Message/?Menu=YES&Exe=&Uin=" & Text1.Text, vbNormalFocus
VBS = Text5.Text '文件路径
ShellExecute ByVal 0, "open", VBS, ByVal 0, ByVal 0, ByVal 0
End Sub
Private Sub Command2_Click()
End
End Sub
