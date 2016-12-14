VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10890
   LinkTopic       =   "Form8"
   ScaleHeight     =   6120
   ScaleWidth      =   10890
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "密码修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   4
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton queren 
         Caption         =   "确认"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   2
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton fanhui 
         Caption         =   "返回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         TabIndex        =   1
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1920
         TabIndex        =   5
         Top             =   1920
         Width           =   720
      End
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -240
      Picture         =   "用户管理.frx":0000
      Top             =   -480
      Width           =   18420
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
   Set cn = New ADODB.Connection
   Set rs = New ADODB.Recordset
   connectstr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=校运会管理2.0"
   cn.ConnectionString = connectstr
   cn.Open
   
   Text1.Text = gsUserName
   Text1.Enabled = False
End Sub
Private Sub fanhui_Click()
  Unload Me
  Form2.Show
End Sub

Private Sub queren_Click()
  Dim strSQL1 As String
  Dim strSQL2 As String
 strSQL1 = " update 登录表 set password='" & Text2.Text & "' where username='" & Text1.Text & "'"
 If rs.State = 1 Then
   rs.Close
   End If
     rs.Open strSQL1, cn, 1, 3
       MsgBox "密码修改成功", 64, "信息提示"


End Sub

Private Sub Text2_Change()
  Text2.PasswordChar = "*"
End Sub
