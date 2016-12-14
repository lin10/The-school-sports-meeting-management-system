VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   9540
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton exit 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "登录"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Width           =   1875
   End
   Begin VB.ComboBox cboUserType 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "权限："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label lalLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "校运会管理系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   3150
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -2160
      Picture         =   "登录.frx":0000
      Top             =   -720
      Width           =   18420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Form_Load()
    Dim gsUserName As String
    Dim connectstr As String
    Set cn = New ADODB.Connection
   Set rs = New ADODB.Recordset
   connectstr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=校运会管理2.0"
   cn.ConnectionString = connectstr
   cn.Open
   
    cboUserType.Clear
    '在cboUserType中加载用于表示用户类型的字符串
    cboUserType.AddItem "学生", 0
    cboUserType.AddItem "裁判", 1
    cboUserType.AddItem "管理员", 2
    
    cboUserType.ListIndex = 0

End Sub
Private Sub cmdOK_Click()
  If Text1.Text = "" Then
      MsgBox "用户名不能为空！", vbOKOnly + vbInformation, "友情提示"
      Text1.SetFocus
      Exit Sub
    End If
    
   If Text2.Text = "" Then
      MsgBox "密码不能为空！", vbOKOnly + vbInformation, "友情提示"
      Text2.SetFocus
      Exit Sub
    End If
    '取得用户输入的用户名和密码
    gnUserType = cboUserType.ListIndex
    gsUserName = Text1.Text
    Dim strSQL As String
    strSQL = "select * from 登录表 where username='" & Text1.Text & "'" & "and password='" & Text2.Text & "'and qxid='" & gnUserType & "'"

    If rs.State = adStateOpen Then
      rs.Close
    End If
      rs.Open strSQL, cn, 1, 3
    If rs.EOF Then
       try_times = try_times + 1
       If try_time >= 3 Then
         MsgBox "您已经三次尝试进入本系统，均不成功，系统将自动关闭", vbOKOnly + vbCritical, "警告"
         Unload Me
      Else
         MsgBox "对不起，用户名不存在或密码或权限错误！", vbOKOnly + vbQuestion, "警告"
         Text1.SetFocus
         Text1.Text = ""
         Text2.Text = ""
       End If
      Else
        Unload Me
        Form2.Show
    End If
End Sub
Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Text2_Change()
   Text2.PasswordChar = "*"
End Sub
