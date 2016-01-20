Imports Microsoft.Win32
Public Class Main

    Public IsFirstRun As Boolean             '此程序是否第一次在本机上运行
    Public IsBackTip As Boolean             '是否后台提示

    Public NowTime As Date         '当前时间，需要时时更新，所以不给初始值

    Public IsSignIn As Boolean = False            '是否已签到
    Public online_time As String = "0.00"         '当前在线时间

#Region "窗体事件"

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.ApplicationExitCall Then      '退出程序
            If Me.IsSignIn Then
                EndDate = ServerTime()
                If EndDate.Year = 1991 OrElse Not sql("update 时间统计 set 离开='" & EndDate.ToString("T") & "',合计时间='" & TimeDiff(StartDate, EndDate) & "' where ID=" & Sign_Identity) Then
                    If MsgBox("无法正常签退，可能是网络连接故障。" & vbCrLf _
                             & "点击 确定 将强行退出，此次签到记录作废。" & vbCrLf _
                             & "点击 取消 将不会退出，可修复网络连接后再尝试退出。", _
                             MsgBoxStyle.OkCancel Or MsgBoxStyle.Critical, SoftName) = MsgBoxResult.Ok Then
                        '强行签退
                    Else
                        e.Cancel = True
                    End If
                End If
                sql_rele(conn, cmd)
            End If
        ElseIf e.CloseReason = CloseReason.UserClosing Then            '隐藏窗口
            e.Cancel = True
            Me.ShowOrHide(False)
            If Me.IsFirstRun AndAlso Me.IsBackTip Then          '第一次点×提示后台运行
                With Me.NotifyIcon1
                    .Visible = True
                    .BalloonTipIcon = ToolTipIcon.Info
                    .BalloonTipTitle = UserName
                    .BalloonTipText = "我在后台运行呢~~嘿嘿"
                    .ShowBalloonTip(30)
                End With
                Me.IsBackTip = False
            End If
        Else                         '关机自动签退，如果签退失败不提示         包含e.CloseReason = CloseReason.WindowsShutDown
            If Me.IsSignIn Then
                EndDate = ServerTime()
                sql("update 时间统计 set 离开='" & EndDate.ToString("T") & "',合计时间='" & TimeDiff(StartDate, EndDate) & "' where ID=" & Sign_Identity)
                sql_rele(conn, cmd)
            End If
        End If
        NotifyIcon1.Visible = False              '新加，解决托盘图标遗留问题

    End Sub

    Private Sub Main_HandleCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.HandleCreated
        Me.Opacity = 0
        Me.TextBMotto.Focus()
        Me.Label1.Text = "版本号：" & Version

        Dim STR1 As String
        STR1 = Application.StartupPath
        If STR1.Substring(0, 2) = "\\" Then
            MsgBox("请不要在服务器上运行，请安装到本地后再运行。")
            Application.Exit()
            Exit Sub
        End If

        '检测程序是否已经在运行
        If Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1 Then
            MsgBox("检测到程序正在运行！", MsgBoxStyle.Critical, SoftName)
            Application.Exit()
            Exit Sub
        End If

        '检测网络
        If Not My.Computer.Network.IsAvailable Then
            MsgBox("电脑不能联网，请检查网络连接", MsgBoxStyle.Critical, SoftName)
            Application.Exit()
            Exit Sub
        End If

        init()

        NowTime = ServerTime()

        '将本程序和注册表统一
        Dim reg As RegistryKey = My.Computer.Registry.CurrentUser
        reg = reg.CreateSubKey("Software\ELAB\" & SoftName)
        If CInt(reg.GetValue("Version")) = Version AndAlso reg.GetValue("Run") = True Then
            Me.IsFirstRun = False
            Me.IsBackTip = False
        Else                           '第一次运行
            Me.IsFirstRun = True
            Me.IsBackTip = True
            reg.SetValue("Version", Version)
            reg.SetValue("Run", True)
        End If
        reg.Close()

        '检测是否开机启动
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", SoftName, Nothing) Is Nothing Then
            Me.开机启动ToolStripMenuItem.Checked = False
        Else
            Me.开机启动ToolStripMenuItem.Checked = True
        End If

        NowWeek = GetWeek()                   '获取周次

        '获取基本信息：用户名，学号，组别，手机号，座右铭
        If Not GetInfo() OrElse NowWeek < 0 Then
            MsgBox("出现未知错误！程序将退出！", MsgBoxStyle.Critical, SoftName)
            Application.Exit() : Exit Sub
        End If

        StartDate = ServerTime()

        '检测上次是否正常退出(今天范围内)，如果不是，将上次未签退的记录以0小时退掉，并将离开字段=登入字段(因为智能考核系统是以离开字段=0获取在线用户的)
        If sql("update 时间统计 set 离开=登入 where 离开='0' and 日期='" & StartDate.Year & "-" & StartDate.Month & "-" & StartDate.Day & "' and 学号='" & StuNum & "'") Then
        Else
            MsgBox("出现致命错误！程序将强制退出！", MsgBoxStyle.Exclamation, SoftName)
            IsSignIn = False
            Application.Exit()
            Exit Sub
        End If
        sql_rele(conn, cmd)

        '签到
        If StartDate.Year <> 1991 Then
            Dim str As String = String.Empty
            str = "insert into 时间统计(学号,姓名,组别,日期,周次,登入,离开,合计时间,学期) values ('" _
            & StuNum & "','" & UserName & "','" & Team & "','" & StartDate.Year & "-" & StartDate.Month _
            & "-" & StartDate.Day & "','" & NowWeek & "','" & StartDate.ToString("T") & "','" & "0" & "','" & "0" & "','" & schoolcalendar & "')" _
            & vbCrLf & "select scope_identity() as id"
            Try
                sql(str)
                If myreader.Read = True Then            '获取Identity
                    Sign_Identity = myreader.Item("id").ToString
                    Me.IsSignIn = True
                End If
                sql_rele(conn, cmd)
            Catch ex As Exception
                MsgBox("出现未知错误！程序将退出！", MsgBoxStyle.Critical, SoftName)
                Application.Exit() : Exit Sub
            End Try
        End If

        Me.UpdataInfo()

        '第一次运行，显示帮助信息
        If IsFirstRun Then
            Me.版本特性ToolStripMenuItem.PerformClick()
        End If

        ' Dim NotifyIc As New NotifyIcon
        NotifyIcon1.Visible = True
        With NotifyIcon1
            .Visible = True
            .BalloonTipIcon = ToolTipIcon.Info
            .BalloonTipTitle = TimeStat(StartDate) & "好！ " & Team & " " & UserName
            If HappyMotto = "" Then              'BalloonTip不允许显示为空
                .BalloonTipText = "已签到！"
            Else
                .BalloonTipText = HappyMotto
            End If
            If IsFirstRun Then
                .BalloonTipText &= vbCrLf & "如果显示得不是你的名字，请找软件组组长更改！"
            End If

            .ShowBalloonTip(30)          '显示小气球
        End With

        '计算当前在线时间
        Dim str_time As String = "select SUM(合计时间) from 时间统计 where 学号 = '" & StuNum & "' and 周次 = " & NowWeek & " and 学期 ='" & schoolcalendar & "'"
        If sql(str_time) AndAlso myreader.Read Then
            online_time = myreader.Item(0)
        End If
        sql_rele(conn, cmd)
    End Sub

    'Esc键隐藏窗口
    Private Sub Main_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.ShowOrHide(False)
        End If
    End Sub

    '显示/隐藏 窗口
    Sub ShowOrHide(ByVal isshow As Boolean)
        Me.详细信息ToolStripMenuItem.Checked = isshow
        If isshow Then
            NowTime = ServerTime()
            If NowTime.Year = 1991 Then
                '年份1991年表示获取服务器时间出错
                Me.ToolStripStatusLabel2.Text = "获取失败"
            Else
                Me.ToolStripStatusLabel2.Text = NowTime.ToShortDateString & " " & NowTime.ToLongTimeString
                NowTime = NowTime.AddSeconds(1)
            End If
            'info_notify()
        End If
        Me.Visible = isshow
    End Sub

    '还给其他窗口焦点
    Private Sub Main_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.Hide()
        Me.Opacity = 1
        详细信息ToolStripMenuItem.Checked = False
    End Sub

#End Region

#Region "右下角事件"

    '双击右下角
    Private Sub NotifyIcon_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.ShowOrHide(Not Me.Visible)
    End Sub

    Private Sub 详细信息ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 详细信息ToolStripMenuItem.Click
        Me.ShowOrHide(Not Me.Visible)
    End Sub

    '注册表实现
    Private Sub 开机启动ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 开机启动ToolStripMenuItem.Click
        Me.开机启动ToolStripMenuItem.Checked = Not Me.开机启动ToolStripMenuItem.Checked
        If Me.开机启动ToolStripMenuItem.Checked Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", SoftName, """" & Application.StartupPath & "\" & SoftName & ".exe" & """")
        Else
            My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(SoftName, False)
        End If
    End Sub

    Private Sub 版本特性ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 版本特性ToolStripMenuItem.Click
        MsgBox("考核系统数据库结构更新" & vbCrLf & vbCrLf _
             & "注意：旧版本已停用，须重新从共享上下载！", , "新版功能")
    End Sub

    Private Sub 使用须知ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 使用须知ToolStripMenuItem.Click
        MsgBox("1.程序运行后就会自动签到，不过别忘了添加开机启动" & vbCrLf & vbCrLf _
             & "2.本客户端支持签到和值班，更多功能请打开""科中管理系统""" & vbCrLf & vbCrLf _
             & "3.值班信息：下午时段为12:30-17:30，晚上时段为17:30-22:00，允许半小时的迟到早退。请大家按时值班" & vbCrLf & vbCrLf _
             & "以上                                                              by Vigi, Zhao", , "使用须知")
    End Sub

    Private Sub 退出程序ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 退出程序ToolStripMenuItem.Click
        Application.Exit()
    End Sub

    '鼠标放在托盘图片上显示当前在线时间
    Private Sub NotifyIcon_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseMove
        NotifyIcon1.Text = "当前在线时间为：" & online_time & "小时"
    End Sub
#End Region

#Region "个人信息控制及按钮事件"

    '对手机号文本框的约束
    Private Sub TextB_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBPhone.TextChanged, TextBMotto.TextChanged
        Me.ButChange.Enabled = True
        Me.ButChange.Text = "应用设置"
    End Sub

    '更新详细信息
    Sub UpdataInfo()
        Me.TextBName.Text = UserName
        Me.TextBNum.Text = StuNum
        Me.TextBTeam.Text = Team
        Me.TextBGrade.Text = Grade
        Me.TextBSign.Text = StartDate.ToString("T")
        Me.TextBPhone.Text = Phone
        Me.TextBMotto.Text = HappyMotto
        Me.ButChange.Enabled = False
    End Sub

    '应用设置（保存手机号码和座右铭）
    Private Sub ButChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButChange.Click
        Dim i As Short : Dim islegal As Boolean = True   '标识检测文本是否合法
        For i = 1 To Me.TextBPhone.Text.Length
            If Not Char.IsNumber(Me.TextBPhone.Text, i - 1) Then
                islegal = False
            End If
        Next
        If islegal Then
            Me.TextBPhone.Text = StrConv(Me.TextBPhone.Text, VbStrConv.Narrow)
            Me.ButChange.Enabled = False
            Me.ButChange.Text = "稍后...."
            If Not sql("update [newmember] set [电话]='" & Me.TextBPhone.Text.Trim & "',[座右铭]='" & Me.TextBMotto.Text.Trim & "' where [学号]='" & StuNum & "'") Then
                MsgBox("更新失败！请检查网络连接", MsgBoxStyle.Critical, SoftName)
                Me.ButChange.Enabled = True
                Me.ButChange.Text = "应用设置"
                Me.ButChange.Focus()
            Else
                Phone = Me.TextBPhone.Text.Trim
                HappyMotto = Me.TextBMotto.Text.Trim
                Me.ButChange.Text = "设置成功"
                Me.TextBMotto.Focus()
            End If
            sql_rele(conn, cmd)
        Else
            Me.ErrorProvider.SetError(Me.TextBPhone, "只允许使用数字！")
            Me.TextBPhone.Focus()
        End If
    End Sub

#End Region

    
    
    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class