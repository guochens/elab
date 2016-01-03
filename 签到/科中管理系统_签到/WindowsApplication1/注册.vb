Public Class 注册

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Application.Exit()

    End Sub

    Private Sub 注册_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim stu_num As String = "select * from newmember where 学号='" & TextBox1.Text & "'"
        sql(stu_num)
        If myreader.Read Then
            If MsgBox("如果你的名字是" & myreader.Item(2) & ",请继续注册，如果不是请找学长联系并退出", vbYesNo) = vbYes Then   '6
                Dim resg As New register
                Me.Hide()
                number_me = TextBox1.Text
                resg.ShowDialog()
            Else
                Application.Exit()

            End If
        Else
            MsgBox("抱歉，该学号非科中新生学号，请重新确认")
        End If
        sql_rele(conn, cmd)
        

    End Sub
End Class