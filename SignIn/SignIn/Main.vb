Imports Microsoft.Win32
Public Class Main

    Public IsFirstRun As Boolean             '�˳����Ƿ��һ���ڱ���������
    Public IsBackTip As Boolean             '�Ƿ��̨��ʾ

    Public NowTime As Date         '��ǰʱ�䣬��Ҫʱʱ���£����Բ�����ʼֵ

    Public IsSignIn As Boolean = False            '�Ƿ���ǩ��
    Public online_time As String = "0.00"         '��ǰ����ʱ��

#Region "�����¼�"

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.ApplicationExitCall Then      '�˳�����
            If Me.IsSignIn Then
                EndDate = ServerTime()
                If EndDate.Year = 1991 OrElse Not sql("update ʱ��ͳ�� set �뿪='" & EndDate.ToString("T") & "',�ϼ�ʱ��='" & TimeDiff(StartDate, EndDate) & "' where ID=" & Sign_Identity) Then
                    If MsgBox("�޷�����ǩ�ˣ��������������ӹ��ϡ�" & vbCrLf _
                             & "��� ȷ�� ��ǿ���˳����˴�ǩ����¼���ϡ�" & vbCrLf _
                             & "��� ȡ�� �������˳������޸��������Ӻ��ٳ����˳���", _
                             MsgBoxStyle.OkCancel Or MsgBoxStyle.Critical, SoftName) = MsgBoxResult.Ok Then
                        'ǿ��ǩ��
                    Else
                        e.Cancel = True
                    End If
                End If
                sql_rele(conn, cmd)
            End If
        ElseIf e.CloseReason = CloseReason.UserClosing Then            '���ش���
            e.Cancel = True
            Me.ShowOrHide(False)
            If Me.IsFirstRun AndAlso Me.IsBackTip Then          '��һ�ε����ʾ��̨����
                With Me.NotifyIcon1
                    .Visible = True
                    .BalloonTipIcon = ToolTipIcon.Info
                    .BalloonTipTitle = UserName
                    .BalloonTipText = "���ں�̨������~~�ٺ�"
                    .ShowBalloonTip(30)
                End With
                Me.IsBackTip = False
            End If
        Else                         '�ػ��Զ�ǩ�ˣ����ǩ��ʧ�ܲ���ʾ         ����e.CloseReason = CloseReason.WindowsShutDown
            If Me.IsSignIn Then
                EndDate = ServerTime()
                sql("update ʱ��ͳ�� set �뿪='" & EndDate.ToString("T") & "',�ϼ�ʱ��='" & TimeDiff(StartDate, EndDate) & "' where ID=" & Sign_Identity)
                sql_rele(conn, cmd)
            End If
        End If
        NotifyIcon1.Visible = False              '�¼ӣ��������ͼ����������

    End Sub

    Private Sub Main_HandleCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.HandleCreated
        Me.Opacity = 0
        Me.TextBMotto.Focus()
        Me.Label1.Text = "�汾�ţ�" & Version

        Dim STR1 As String
        STR1 = Application.StartupPath
        If STR1.Substring(0, 2) = "\\" Then
            MsgBox("�벻Ҫ�ڷ����������У��밲װ�����غ������С�")
            Application.Exit()
            Exit Sub
        End If

        '�������Ƿ��Ѿ�������
        If Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1 Then
            MsgBox("��⵽�����������У�", MsgBoxStyle.Critical, SoftName)
            Application.Exit()
            Exit Sub
        End If

        '�������
        If Not My.Computer.Network.IsAvailable Then
            MsgBox("���Բ���������������������", MsgBoxStyle.Critical, SoftName)
            Application.Exit()
            Exit Sub
        End If

        init()

        NowTime = ServerTime()

        '���������ע���ͳһ
        Dim reg As RegistryKey = My.Computer.Registry.CurrentUser
        reg = reg.CreateSubKey("Software\ELAB\" & SoftName)
        If CInt(reg.GetValue("Version")) = Version AndAlso reg.GetValue("Run") = True Then
            Me.IsFirstRun = False
            Me.IsBackTip = False
        Else                           '��һ������
            Me.IsFirstRun = True
            Me.IsBackTip = True
            reg.SetValue("Version", Version)
            reg.SetValue("Run", True)
        End If
        reg.Close()

        '����Ƿ񿪻�����
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", SoftName, Nothing) Is Nothing Then
            Me.��������ToolStripMenuItem.Checked = False
        Else
            Me.��������ToolStripMenuItem.Checked = True
        End If

        NowWeek = GetWeek()                   '��ȡ�ܴ�

        '��ȡ������Ϣ���û�����ѧ�ţ�����ֻ��ţ�������
        If Not GetInfo() OrElse NowWeek < 0 Then
            MsgBox("����δ֪���󣡳����˳���", MsgBoxStyle.Critical, SoftName)
            Application.Exit() : Exit Sub
        End If

        StartDate = ServerTime()

        '����ϴ��Ƿ������˳�(���췶Χ��)��������ǣ����ϴ�δǩ�˵ļ�¼��0Сʱ�˵��������뿪�ֶ�=�����ֶ�(��Ϊ���ܿ���ϵͳ�����뿪�ֶ�=0��ȡ�����û���)
        If sql("update ʱ��ͳ�� set �뿪=���� where �뿪='0' and ����='" & StartDate.Year & "-" & StartDate.Month & "-" & StartDate.Day & "' and ѧ��='" & StuNum & "'") Then
        Else
            MsgBox("�����������󣡳���ǿ���˳���", MsgBoxStyle.Exclamation, SoftName)
            IsSignIn = False
            Application.Exit()
            Exit Sub
        End If
        sql_rele(conn, cmd)

        'ǩ��
        If StartDate.Year <> 1991 Then
            Dim str As String = String.Empty
            str = "insert into ʱ��ͳ��(ѧ��,����,���,����,�ܴ�,����,�뿪,�ϼ�ʱ��,ѧ��) values ('" _
            & StuNum & "','" & UserName & "','" & Team & "','" & StartDate.Year & "-" & StartDate.Month _
            & "-" & StartDate.Day & "','" & NowWeek & "','" & StartDate.ToString("T") & "','" & "0" & "','" & "0" & "','" & schoolcalendar & "')" _
            & vbCrLf & "select scope_identity() as id"
            Try
                sql(str)
                If myreader.Read = True Then            '��ȡIdentity
                    Sign_Identity = myreader.Item("id").ToString
                    Me.IsSignIn = True
                End If
                sql_rele(conn, cmd)
            Catch ex As Exception
                MsgBox("����δ֪���󣡳����˳���", MsgBoxStyle.Critical, SoftName)
                Application.Exit() : Exit Sub
            End Try
        End If

        Me.UpdataInfo()

        '��һ�����У���ʾ������Ϣ
        If IsFirstRun Then
            Me.�汾����ToolStripMenuItem.PerformClick()
        End If

        ' Dim NotifyIc As New NotifyIcon
        NotifyIcon1.Visible = True
        With NotifyIcon1
            .Visible = True
            .BalloonTipIcon = ToolTipIcon.Info
            .BalloonTipTitle = TimeStat(StartDate) & "�ã� " & Team & " " & UserName
            If HappyMotto = "" Then              'BalloonTip��������ʾΪ��
                .BalloonTipText = "��ǩ����"
            Else
                .BalloonTipText = HappyMotto
            End If
            If IsFirstRun Then
                .BalloonTipText &= vbCrLf & "�����ʾ�ò���������֣�����������鳤���ģ�"
            End If

            .ShowBalloonTip(30)          '��ʾС����
        End With

        '���㵱ǰ����ʱ��
        Dim str_time As String = "select SUM(�ϼ�ʱ��) from ʱ��ͳ�� where ѧ�� = '" & StuNum & "' and �ܴ� = " & NowWeek & " and ѧ�� ='" & schoolcalendar & "'"
        If sql(str_time) AndAlso myreader.Read Then
            online_time = myreader.Item(0)
        End If
        sql_rele(conn, cmd)
    End Sub

    'Esc�����ش���
    Private Sub Main_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.ShowOrHide(False)
        End If
    End Sub

    '��ʾ/���� ����
    Sub ShowOrHide(ByVal isshow As Boolean)
        Me.��ϸ��ϢToolStripMenuItem.Checked = isshow
        If isshow Then
            NowTime = ServerTime()
            If NowTime.Year = 1991 Then
                '���1991���ʾ��ȡ������ʱ�����
                Me.ToolStripStatusLabel2.Text = "��ȡʧ��"
            Else
                Me.ToolStripStatusLabel2.Text = NowTime.ToShortDateString & " " & NowTime.ToLongTimeString
                NowTime = NowTime.AddSeconds(1)
            End If
            'info_notify()
        End If
        Me.Visible = isshow
    End Sub

    '�����������ڽ���
    Private Sub Main_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.Hide()
        Me.Opacity = 1
        ��ϸ��ϢToolStripMenuItem.Checked = False
    End Sub

#End Region

#Region "���½��¼�"

    '˫�����½�
    Private Sub NotifyIcon_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.ShowOrHide(Not Me.Visible)
    End Sub

    Private Sub ��ϸ��ϢToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ��ϸ��ϢToolStripMenuItem.Click
        Me.ShowOrHide(Not Me.Visible)
    End Sub

    'ע���ʵ��
    Private Sub ��������ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ��������ToolStripMenuItem.Click
        Me.��������ToolStripMenuItem.Checked = Not Me.��������ToolStripMenuItem.Checked
        If Me.��������ToolStripMenuItem.Checked Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", SoftName, """" & Application.StartupPath & "\" & SoftName & ".exe" & """")
        Else
            My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(SoftName, False)
        End If
    End Sub

    Private Sub �汾����ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles �汾����ToolStripMenuItem.Click
        MsgBox("����ϵͳ���ݿ�ṹ����" & vbCrLf & vbCrLf _
             & "ע�⣺�ɰ汾��ͣ�ã������´ӹ��������أ�", , "�°湦��")
    End Sub

    Private Sub ʹ����֪ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ʹ����֪ToolStripMenuItem.Click
        MsgBox("1.�������к�ͻ��Զ�ǩ����������������ӿ�������" & vbCrLf & vbCrLf _
             & "2.���ͻ���֧��ǩ����ֵ�࣬���๦�����""���й���ϵͳ""" & vbCrLf & vbCrLf _
             & "3.ֵ����Ϣ������ʱ��Ϊ12:30-17:30������ʱ��Ϊ17:30-22:00�������Сʱ�ĳٵ����ˡ����Ұ�ʱֵ��" & vbCrLf & vbCrLf _
             & "����                                                              by Vigi, Zhao", , "ʹ����֪")
    End Sub

    Private Sub �˳�����ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles �˳�����ToolStripMenuItem.Click
        Application.Exit()
    End Sub

    '����������ͼƬ����ʾ��ǰ����ʱ��
    Private Sub NotifyIcon_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseMove
        NotifyIcon1.Text = "��ǰ����ʱ��Ϊ��" & online_time & "Сʱ"
    End Sub
#End Region

#Region "������Ϣ���Ƽ���ť�¼�"

    '���ֻ����ı����Լ��
    Private Sub TextB_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBPhone.TextChanged, TextBMotto.TextChanged
        Me.ButChange.Enabled = True
        Me.ButChange.Text = "Ӧ������"
    End Sub

    '������ϸ��Ϣ
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

    'Ӧ�����ã������ֻ��������������
    Private Sub ButChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButChange.Click
        Dim i As Short : Dim islegal As Boolean = True   '��ʶ����ı��Ƿ�Ϸ�
        For i = 1 To Me.TextBPhone.Text.Length
            If Not Char.IsNumber(Me.TextBPhone.Text, i - 1) Then
                islegal = False
            End If
        Next
        If islegal Then
            Me.TextBPhone.Text = StrConv(Me.TextBPhone.Text, VbStrConv.Narrow)
            Me.ButChange.Enabled = False
            Me.ButChange.Text = "�Ժ�...."
            If Not sql("update [newmember] set [�绰]='" & Me.TextBPhone.Text.Trim & "',[������]='" & Me.TextBMotto.Text.Trim & "' where [ѧ��]='" & StuNum & "'") Then
                MsgBox("����ʧ�ܣ�������������", MsgBoxStyle.Critical, SoftName)
                Me.ButChange.Enabled = True
                Me.ButChange.Text = "Ӧ������"
                Me.ButChange.Focus()
            Else
                Phone = Me.TextBPhone.Text.Trim
                HappyMotto = Me.TextBMotto.Text.Trim
                Me.ButChange.Text = "���óɹ�"
                Me.TextBMotto.Focus()
            End If
            sql_rele(conn, cmd)
        Else
            Me.ErrorProvider.SetError(Me.TextBPhone, "ֻ����ʹ�����֣�")
            Me.TextBPhone.Focus()
        End If
    End Sub

#End Region

    
    
    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class