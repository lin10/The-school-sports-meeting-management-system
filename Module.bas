Attribute VB_Name = "Module1"
    '改变用来保存当前的用户类型，其中0表示学生，1表示裁判，2表示管理员
     Public gnUserType As Integer
    '表示当前用户名
       Public gsUserName As String
   '数据连接，在打开记录集和操作记录数据时使用
      Public cn As New ADODB.Connection
      Public rs As New ADODB.Recordset
