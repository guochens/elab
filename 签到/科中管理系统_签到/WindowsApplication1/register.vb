Public Class register

    '限制输入内容
    Private Sub TextBox_num_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_num.KeyPress
        If (Char.IsDigit(e.KeyChar) And TextBox_num.TextLength < 9) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox_tel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_tel.KeyPress
        If (Char.IsDigit(e.KeyChar) And TextBox_tel.TextLength < 11) Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    'Private Sub TextBox_user_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '    If TextBox_user.TextLength < 3 Or e.KeyChar = Chr(8) Then
    '        e.Handled = False
    '    Else
    '        e.Handled = True
    '    End If
    'End Sub

    Private Sub TextBox_name_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_name.KeyPress
        If TextBox_name.TextLength < 3 Or e.KeyChar = Chr(8) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim scs As String = ""
        If TextBox_num.Text <> "" And TextBox_user.Text <> "" And TextBox_name.Text <> "" And ComboBox_sex.Text <> "" And TextBox_col.Text <> "" And TextBox_group.Text <> "" And TextBox_tel.Text <> "" Then
            scs = "update newmember set 性别='" & ComboBox_sex.Text & "',USERNAME='" & TextBox_user.Text & "',院系='" & TextBox_col.Text & "',职务='" & TextBox6.Text & "',年级='" & TextBox7.Text & "',文件版本='" & TextBox8.Text & "',座右铭='" & RichTextBox1.Text & "' where 学号='" & number_me & "'"
            sql(scs)
            sql_rele(conn, cmd)

            '检验
            Dim num_temp As String
            Dim sqlstr2 As String = "select * from newmember where 学号='" & TextBox_num.Text & "'"
            sql(sqlstr2)
            If myreader.Read() Then
                num_temp = myreader.Item(1)
                If num_temp = TextBox_user.Text Then
                    MsgBox("注册成功！请重新打开本软件")
                    Application.Exit()
                Else
                    MsgBox("注册失败！")
                End If
            Else
                MsgBox("注册失败！")
            End If
            sql_rele(conn, cmd)
        Else
            MsgBox("请填写完所有信息之后再提交！", , "警告")

        End If
    End Sub

    Private Sub register_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox_user.Text = Environment.GetEnvironmentVariable("username")
        TextBox_user.ReadOnly = True
        Dim sqlconn As String = "select * from newmember where 学号='" & number_me & "'"
        sql(sqlconn)
        While myreader.Read
            TextBox_num.Text = myreader.Item(0)
            TextBox_name.Text = myreader.Item(2)
            TextBox_group.Text = myreader.Item(5)
            TextBox_tel.Text = myreader.Item(6)
            TextBox_num.ReadOnly = True
            TextBox_name.ReadOnly = True
            TextBox_group.ReadOnly = True
            TextBox_tel.ReadOnly = True
        End While
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Application.Exit()

    End Sub
End Class
