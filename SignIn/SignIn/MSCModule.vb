Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Resources
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Net.Sockets
Imports System.Text

Module MSCModule

    '需要写入注册表的常量
    Public Const Version As Integer = 1
    Public Const SoftName = "科中管理系统_签到"

    '数据库相关
    Public Sign_Identity As String = String.Empty               '本次签到在数据库上对应的ID值

    'Public connstr As String = "server=192.168.1.252;database=student;uid=ta;pwd=elab2013;connection timeout=5;"  '连接远程服务器
    Public connstr As String = "Data Source=.;Integrated Security=True;Database=student"              '连接本地
    Public conn As SqlConnection, cmd As SqlCommand, myreader As SqlDataReader  '数据库连接读取等相关变量
    Public conn_1 As SqlConnection, cmd_1 As SqlCommand, myreader_1 As SqlDataReader  '数据库连接读取等相关变量
    Public conn_2 As SqlConnection, cmd_2 As SqlCommand, myreader_2 As SqlDataReader  '数据库连接读取等相关变量

    Public strSql, selectcmd, insertcmd, updatecmd, delcmd As String            '数据库字符串储存变量

    Public StartDate As Date                 '程序启动时间
    Public EndDate As Date                   '程序退出时间

    Public schoolcalendar As String    '学期（例：2013秋）

    '获取基本信息：用户名，学号，组别，手机号，座右铭，年级
    Public UserName As String = String.Empty
    Public StuNum As String = String.Empty
    Public Team As String = String.Empty
    Public Phone As String = String.Empty
    Public HappyMotto As String = String.Empty
    Public Grade As String = String.Empty
    Public user As String = String.Empty

    '已注册学号
    Public sti(1) As String
    '该生学号
    Public number_me As String
    '解决注册application问题
    'public 

#Region "执行SQL语句函数"

    Public Function sql(ByVal sqlcmd As String) As Boolean
        conn = New SqlConnection(connstr)
        Try
            conn.Open()
            cmd = New SqlCommand(sqlcmd, conn)
            myreader = cmd.ExecuteReader
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            conn.Close()
            Return False
        End Try
    End Function

    Public Function sql_1(ByVal sqlcmd As String) As Boolean
        conn_1 = New SqlConnection(connstr)
        Try
            conn_1.Open()
            cmd_1 = New SqlCommand(sqlcmd, conn_1)
            myreader_1 = cmd_1.ExecuteReader
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            conn_1.Close()
            Return False
        End Try
    End Function

    Public Function sql_2(ByVal sqlcmd As String) As Boolean
        conn_2 = New SqlConnection(connstr)
        Try
            conn_2.Open()
            cmd_2 = New SqlCommand(sqlcmd, conn_2)
            myreader_2 = cmd_2.ExecuteReader
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            conn_2.Close()
            Return False
        End Try
    End Function

    Public Function gettable(ByVal tablename As String, ByVal sqlstr As String) As DataTable
        conn = New SqlConnection(connstr)
        Dim adapter As New SqlDataAdapter(sqlstr, conn)
        Dim newtable As New DataTable
        Dim timeds As New DataSet
        Try
            conn.Open()
            timeds.Clear()
            adapter.Fill(timeds, tablename)
            newtable = timeds.Tables(tablename)
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
            conn.Close()
        End Try
        Return timeds.Tables(tablename)
    End Function

    Public Sub sql_rele(ByVal SqlCon As SqlConnection, ByVal SqlCmd As SqlCommand)              '释放数据库占用资源
        If Not SqlCmd Is Nothing Then
            SqlCmd.Dispose()
            SqlCmd = Nothing
        End If
        If Not SqlCon Is Nothing Then
            If SqlCon.State <> ConnectionState.Closed Then
                SqlCon.Close()
            End If
            SqlCon.Dispose()
            SqlCon = Nothing
        End If
    End Sub

#End Region

#Region "读取上课学期信息"

    Public Sub init()
        selectcmd = "select 学期 from [开学日期]"
        If sql(selectcmd) Then
            If myreader.Read Then
                schoolcalendar = myreader.Item(0).ToString.Trim
            End If
        End If
        sql_rele(conn, cmd)
    End Sub

#End Region

    Public Function GetInfo() As Boolean
        '******用一次请求完成，获取基本信息
        Dim year As Integer
        Dim mac As String = String.Empty
        Dim ip As String = String.Empty

        sql("select * from [newmember] where USERNAME='" & Environment.GetEnvironmentVariable("username") & "'")
        If myreader.Read() Then
            StuNum = myreader("学号")
            sql("select * from [newmember] where [学号]=" & StuNum)
            While myreader.Read
                UserName = myreader("姓名")
                year = DateTime.Today.Year() - 1000 * CType(Val(StuNum(0)), Integer) - 100 * CType(Val(StuNum(1)), Integer) - 10 * CType(Val(StuNum(2)), Integer) - CType(Val(StuNum(3)), Integer)
                If IsDBNull(myreader("组别")) Then
                    Phone = String.Empty
                Else
                    Team = myreader("组别")
                End If
                If IsDBNull(myreader("电话")) Then
                    Phone = String.Empty
                Else
                    Phone = myreader("电话")
                End If
                If DateTime.Today.Month > 8 Then
                    year = year + 1
                End If
                If year = 1 Then
                    Grade = "大一"
                ElseIf year = 2 Then
                    Grade = "大二"
                ElseIf year = 3 Then
                    Grade = "大三"
                ElseIf year = 4 Then
                    Grade = "大四"
                Else
                    Grade = "往届"
                End If
                If IsDBNull(myreader("座右铭")) Then
                    HappyMotto = String.Empty
                Else
                    HappyMotto = myreader("座右铭")
                End If
            End While
        Else
            Dim res As New 注册
            Main.Hide()
            res.ShowDialog()
        End If
        sql_rele(conn, cmd)

        
        Return True
    End Function

    '获取周次(这周是这学期的第几周)
    Public NowWeek As Integer
    Public Function GetWeek() As Integer
        Dim week As Integer
        Dim str As String
        str = "select top 1 datediff(day,[开学日期],getdate())/7+1 as cache from [开学日期]"
        sql(str)
        If myreader.Read Then
            week = myreader.Item("cache")
        Else
            week = -1
        End If
        sql_rele(conn, cmd)
        Return week
    End Function

    '返回当前服务器时间，错误返回"1991-01-01 00:00:00"
    Public Function ServerTime() As Date
        Dim nowtime As Date
        sql("select convert(varchar(20),getdate(),20) as date")
        If myreader.Read = True Then
            nowtime = myreader.Item("date")
        Else
            nowtime = New Date(1991, 1, 1, 0, 0, 0)
        End If
        sql_rele(conn, cmd)
        Return nowtime
    End Function
    '返回单双周
    Public Function 单双周() As String
        If NowWeek Mod 2 = 0 Then
            Return "双周"
        Else
            Return "单周"
        End If
    End Function

    '返回两个时间的差值，以小时为单位，用了系统内置的函数DateDiff
    Public Function TimeDiff(ByVal time1 As Date, ByVal time2 As Date) As String
        If time1.Year = time2.Year And time1.Month = time2.Month And time1.Day = time2.Day Then
            Return Format(DateDiff(DateInterval.Minute, time1, time2) / 60, "0.00")
        Else
            Return "0"
        End If
    End Function

    '返回指定时间状态，早上，中午。。。
    Public Function TimeStat(ByVal cachetime As Date) As String
        Select Case cachetime.Hour
            Case 5 To 7
                Return "早上"
            Case 8 To 10
                Return "上午"
            Case 11 To 12
                Return "中午"
            Case 13 To 16
                Return "下午"
            Case 17 To 23
                Return "晚上"
            Case Else
                Return "凌晨"
        End Select
    End Function

End Module
