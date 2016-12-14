VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780
   LinkTopic       =   "Form4"
   ScaleHeight     =   5700
   ScaleWidth      =   9780
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "积分管理"
      Height          =   5295
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      Begin VB.CommandButton chaxunxy 
         Caption         =   "学院积分查询"
         Height          =   495
         Left            =   2880
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton shanchu 
         Caption         =   "删除"
         Height          =   495
         Left            =   6600
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton tianjia 
         Caption         =   "添加"
         Height          =   495
         Left            =   6720
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton chaxunath 
         Caption         =   "个人积分查询"
         Height          =   495
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton fanhui 
         Caption         =   "返回"
         Height          =   495
         Left            =   3120
         TabIndex        =   3
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton xiugai 
         Caption         =   "打分"
         Height          =   495
         Left            =   4680
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "积分管理.frx":0000
         Height          =   1935
         Left            =   600
         TabIndex        =   2
         Top             =   2400
         Width           =   7215
         _ExtentX        =   12726
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
         Caption         =   "项目积分"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2520
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
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
         Connect         =   $"积分管理.frx":0015
         OLEDBString     =   $"积分管理.frx":00C4
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "报名"
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询情况如下："
         Height          =   180
         Left            =   1080
         TabIndex        =   17
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运动员积分"
         Height          =   180
         Left            =   5520
         TabIndex        =   16
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "学院积分"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学院编号"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运动员编号"
         Height          =   180
         Left            =   3000
         TabIndex        =   13
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "比赛编号"
         Height          =   180
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -840
      Picture         =   "积分管理.frx":0173
      Top             =   -960
      Width           =   18420
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Form_Load()
   Select Case gnUserType
     Case 0:
      tianjia.Enabled = False
      shanchu.Enabled = False
      xiugai.Enabled = False
    Text3.Enabled = False
     Case 1:
      tianjia.Enabled = True
      shanchu.Enabled = False
      xiugai.Enabled = True
     Case 2:
      tianjia.Enabled = True
      shanchu.Enabled = True
      xiugai.Enabled = True
   End Select
   
   Set cn = New ADODB.Connection
   Set rs = New ADODB.Recordset
   connectstr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=校运会管理2.0"
   cn.ConnectionString = connectstr
   cn.Open
   
   tianjia.Visible = False
   Text5.Enabled = False

   
End Sub
Private Sub chaxunath_Click()
    Dim strSQL1 As String
   strSQL1 = "select * from 报名 where athid='" & Text2.Text & "'"
   If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
    MsgBox "该运动员不存在"
     Adodc1.CommandType = adCmdText
     Adodc1.RecordSource = strSQL1
     Adodc1.Refresh
     
     Else
     
     Text1.Text = rs("matid")
     Text2.Text = rs("athid")
     Text3.Text = rs("jf")
     End If
End Sub
Private Sub chaxunxy_Click()
  Dim strSQL1 As String
   strSQL1 = "select xyid,sum(jf) AS xyjf from 报名 where xyid='" & Text4.Text & "' group by  xyid "
   If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
    MsgBox "该学院积分不存在"
    Adodc1.CommandType = adCmdText
     Adodc1.RecordSource = strSQL1
     Adodc1.Refresh
     
     Else
     
     Text4.Text = rs("xyid")
     Text5.Text = rs("xyjf")
     End If
End Sub




Private Sub xiugai_Click()
   strSQL1 = "select * from 报名 where athid='" & Text2.Text & "'"
    If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
    DataGrid1.AllowUpdate = True
      MsgBox "打分成功", 64, "信息提示"
  Else
  
    MsgBox "信息修改失败", 16, "警告"
  End If
End Sub
Private Sub tianjia_Click()
    strSQL1 = "select * from 报名 where matid='" & Text1.Text & "'"
   If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
      rs.AddNew
      rs("matid") = Text1.Text
      rs("athid") = Text2.Text
      rs("jf") = Text3.Text
    rs.Update
    MsgBox "信息添加成功", 64, "信息提示"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from 报名"
    Adodc1.Refresh
         Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
  Else
  
    MsgBox "该信息已经存在了", 16, "警告"
  End If
End Sub
Private Sub shanchu_Click()
    If Adodc1.Recordset.EOF = False Then
      c = MsgBox("您确定要删除该记录吗？", 32 + 4, "特别提示")
      X = Adodc1.Recordset.Fields(0)
       If c = vbYes Then
       
       strSQL1 = "select * from 报名 where matid='" & Text1.Text & "'"
       If rs.State = 1 Then
         rs.Close
        End If
       rs.Open strSQL1, cn, 1, 3
       Adodc1.Recordset.Delete
       Adodc1.CommandType = adCmdText
       Adodc1.RecordSource = "select * from 报名"
       Adodc1.Refresh
         MsgBox "比赛项目的信息已成功删除", 64, "信息提示"
       Adodc1.Refresh
         Text1.Text = ""
         Text2.Text = ""
         Text3.Text = ""
       End If
      Else
          MsgBox "当前数据库没有可删除的记录", 64, "警告"
      End If
End Sub

Private Sub fanhui_Click()
  Unload Me
  Form2.Show
End Sub
