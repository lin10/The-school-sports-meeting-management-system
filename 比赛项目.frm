VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   LinkTopic       =   "Form3"
   ScaleHeight     =   5595
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "比赛项目"
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      Begin VB.CommandButton xuanchua 
         Caption         =   "宣传"
         Height          =   615
         Left            =   1320
         TabIndex        =   13
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton fanhui 
         Caption         =   "返回"
         Height          =   615
         Left            =   5160
         TabIndex        =   12
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   6480
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton chaxun 
         Caption         =   "查询"
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton tianjia 
         Caption         =   "添加"
         Height          =   495
         Left            =   4320
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton shanchu 
         Caption         =   "删除"
         Height          =   495
         Left            =   6240
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton xiugai 
         Caption         =   "修改"
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "比赛项目.frx":0000
         Height          =   1335
         Left            =   360
         TabIndex        =   2
         Top             =   2640
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2355
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
         Caption         =   "比赛项目"
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
         Width           =   4815
         _ExtentX        =   8493
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
         Connect         =   $"比赛项目.frx":0015
         OLEDBString     =   $"比赛项目.frx":00C4
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "比赛项目"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "比赛编号"
         Height          =   180
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目名称"
         Height          =   180
         Left            =   3000
         TabIndex        =   19
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参赛人数"
         Height          =   180
         Left            =   5640
         TabIndex        =   18
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预赛时间"
         Height          =   180
         Left            =   480
         TabIndex        =   17
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "决赛时间"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "裁判员编号"
         Height          =   180
         Left            =   5520
         TabIndex        =   15
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询情况如下："
         Height          =   180
         Left            =   960
         TabIndex        =   14
         Top             =   2160
         Width           =   1260
      End
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -3960
      Picture         =   "比赛项目.frx":0173
      Top             =   -1200
      Width           =   18420
   End
End
Attribute VB_Name = "Form3"
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
   
    Text1.ToolTipText = "请输入4位比赛编号，例如1001"
   Text4.ToolTipText = "请输入具体时间例如，例2016-04-01 15:00"
   Text5.ToolTipText = "请输入具体时间例如，例2016-04-01 15:00"
   Text6.ToolTipText = "请输入4位裁判员编号，例如1001"
End Sub
Private Sub chaxun_Click()
   Dim strSQL1 As String
   strSQL1 = "select * from 比赛项目 where matid='" & Text1.Text & "' union select * from 比赛项目 where matname='" & Text2.Text & "'"
   If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
    MsgBox "该比赛项目不存在"
     Adodc1.CommandType = adCmdText
     Adodc1.RecordSource = strSQL1
     Adodc1.Refresh
     
     Else
     
     Text1.Text = rs("matid")
     Text2.Text = rs("matname")
     Text3.Text = rs("athnum")
     Text4.Text = rs("ystime")
     Text5.Text = rs("jstime")
     Text6.Text = rs("cpid")
     End If
End Sub


Private Sub xiugai_Click()
   strSQL1 = "select * from 比赛项目 where matid='" & Text1.Text & "' union select * from 比赛项目 where matname='" & Text2.Text & "'"
   If rs.State = 1 Then
   rs.Close
   End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = strSQL1
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.AllowUpdate = True
    MsgBox "信息修改成功", 64, "信息提示"

  Else
  
    MsgBox "信息修改失败", 16, "警告"
  End If
End Sub
Private Sub tianjia_Click()
    strSQL1 = "select * from 比赛项目 where matid='" & Text1.Text & "'"
    If rs.State = 1 Then
      rs.Close
     End If
   rs.Open strSQL1, cn, 1, 3
   If rs.EOF Then
      rs.AddNew
      rs("matid") = Text1.Text
      rs("matname") = Text2.Text
      rs("athnum") = Text3.Text
      rs("ystime") = Text4.Text
      rs("jstime") = Text5.Text
      rs("cpid") = Text6.Text
    rs.Update
    MsgBox "信息添加成功", 64, "信息提示"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from 比赛项目"
    Adodc1.Refresh
         Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
         Text6.Text = ""
  Else
  
    MsgBox "该信息已经存在了", 16, "警告"
  End If
End Sub
Private Sub shanchu_Click()
 If Adodc1.Recordset.EOF = False Then
      c = MsgBox("您确定要删除该记录吗？", 32 + 4, "特别提示")
      X = Adodc1.Recordset.Fields(0)
       If c = vbYes Then
       strSQL1 = "select * from 比赛项目 where matid='" & Text1.Text & "' union select * from 比赛项目 where matname='" & Text2.Text & "'"
       If rs.State = 1 Then
        rs.Close
        End If
       rs.Open strSQL1, cn, 1, 3
       Adodc1.Recordset.Delete
       Adodc1.CommandType = adCmdText
       Adodc1.RecordSource = "select * from 比赛项目"
       Adodc1.Refresh
         MsgBox "比赛项目的信息已成功删除", 64, "信息提示"
       Adodc1.Refresh
         Text1.Text = ""
         Text2.Text = ""
         Text3.Text = ""
         Text4.Text = ""
         Text5.Text = ""
         Text6.Text = ""
       End If
      Else
          MsgBox "当前数据库没有可删除的记录", 64, "警告"
      End If
End Sub
Private Sub xuanchua_Click()
  Unload Me
  Form9.Show
End Sub

Private Sub fanhui_Click()
 Unload Me
 Form2.Show
End Sub
