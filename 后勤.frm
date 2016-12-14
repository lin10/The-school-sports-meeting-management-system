VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form7"
   ScaleHeight     =   6525
   ScaleWidth      =   10140
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "后勤"
      Height          =   5175
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   8535
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton chaxun 
         Caption         =   "查询"
         Height          =   495
         Left            =   2880
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton baoming 
         Caption         =   "报名"
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton shanchu 
         Caption         =   "删除"
         Height          =   495
         Left            =   6240
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton xiugai 
         Caption         =   "修改"
         Height          =   495
         Left            =   4680
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton fanhui 
         Caption         =   "返回"
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   4560
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   2520
         Top             =   1920
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"后勤.frx":0000
         OLEDBString     =   $"后勤.frx":00AF
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "后勤"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "后勤.frx":015E
         Height          =   1935
         Left            =   1080
         TabIndex        =   3
         Top             =   2400
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3413
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "后勤"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "负责物品"
         Height          =   180
         Left            =   4200
         TabIndex        =   15
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "值班时间"
         Height          =   180
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "值班编号"
         Height          =   180
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "值班人学号"
         Height          =   180
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   480
         TabIndex        =   10
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询情况如下："
         Height          =   180
         Left            =   1080
         TabIndex        =   9
         Top             =   2040
         Width           =   1260
      End
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -240
      Picture         =   "后勤.frx":0173
      Top             =   -120
      Width           =   18420
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
   Select Case gnUserType
     Case 0:
      Text2.Text = gsUserName
      Text2.Enabled = False
      baoming.Enabled = True
      shanchu.Enabled = False
      xiugai.Enabled = False
     Case 1:
      baoming.Enabled = False
      shanchu.Enabled = False
      xiugai.Enabled = False
     Case 2:
      baoming.Enabled = True
      shanchu.Enabled = True
      xiugai.Enabled = True
   End Select
   
   Set cn = New ADODB.Connection
   Set rs = New ADODB.Recordset
   connectstr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=校运会管理2.0"
   cn.ConnectionString = connectstr
   cn.Open
   
   Text1.ToolTipText = "请输入值班编号，例1001"
End Sub

Private Sub baoming_Click()
 strSQL1 = "select * from 后勤 where zbid='" & Text1.Text & "'" & "and zbrid='" & Text2.Text & "'"
   If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
      rs.AddNew
      rs("zbid") = Text1.Text
      rs("zbrid") = Text2.Text
      rs("zbtime") = Text3.Text
    rs.Update
    MsgBox "报名成功", 64, "信息提示"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from 后勤 inner join 物品 on 后勤.wpid = 物品.wpid"
    Adodc1.Refresh
         Text1.Text = ""
        Text2.Text = ""
  Else
  
    MsgBox "该值班安排已存在", 16, "警告"
  End If
End Sub

Private Sub chaxun_Click()
   Dim strSQL1 As String
   strSQL1 = "select * from 后勤 where zbrid='" & Text2.Text & "' union select * from 后勤 where zbid='" & Text1.Text & "' "
   If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
    MsgBox "该值班信息不存在"
     Adodc1.CommandType = adCmdText
     Adodc1.RecordSource = strSQL1
     Adodc1.Refresh
     
     Else
     
     Text1.Text = rs("zbid")
     Text2.Text = rs("zbrid")
     Text3.Text = rs("zbtime")
     Text4.Text = rs("wpid")

     End If
End Sub




Private Sub xiugai_Click()
   strSQL1 = "select * from 后勤 where zbid='" & Text1.Text & "'"
    If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
    MsgBox "信息修改成功", 64, "信息提示"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = strSQL1
    Set DataGrid1.DataSource = Adodc1
    
    DataGrid1.AllowUpdate = True
  Else
    MsgBox "信息修改失败", 16, "警告"
  End If
End Sub
Private Sub shanchu_Click()
   If Adodc1.Recordset.EOF = False Then
      c = MsgBox("您确定要删除该记录吗？", 32 + 4, "特别提示")
       If c = vbYes Then
       
       strSQL1 = "select * from 后勤 where zbid='" & Text1.Text & "'"
       If rs.State = 1 Then
         rs.Close
        End If
       rs.Open strSQL1, cn, 1, 3
       Adodc1.Recordset.Delete
       Adodc1.CommandType = adCmdText
       Adodc1.RecordSource = "select * from 后勤"
       Adodc1.Refresh
         MsgBox "值班的信息已成功删除", 64, "信息提示"
       Adodc1.Refresh
         Text1.Text = ""
         Text2.Text = ""
         Text3.Text = ""
         Text4.Text = ""
       End If
      Else
          MsgBox "当前数据库没有可删除的记录", 64, "警告"
      End If
End Sub

Private Sub fanhui_Click()
   Unload Me
   Form2.Show
End Sub

