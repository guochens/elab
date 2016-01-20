<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ConMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.详细信息ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.开机启动ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.科中管理系统ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.版本特性ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.使用须知ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.退出程序ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LabName = New System.Windows.Forms.Label()
        Me.TextBName = New System.Windows.Forms.TextBox()
        Me.TextBTeam = New System.Windows.Forms.TextBox()
        Me.LabTeam = New System.Windows.Forms.Label()
        Me.TextBPhone = New System.Windows.Forms.TextBox()
        Me.LabPhone = New System.Windows.Forms.Label()
        Me.TextBGrade = New System.Windows.Forms.TextBox()
        Me.LabGrade = New System.Windows.Forms.Label()
        Me.TextBMotto = New System.Windows.Forms.TextBox()
        Me.LabMotto = New System.Windows.Forms.Label()
        Me.ButChange = New System.Windows.Forms.Button()
        Me.Timer = New System.Windows.Forms.Timer(Me.components)
        Me.TextBSign = New System.Windows.Forms.TextBox()
        Me.LabSign = New System.Windows.Forms.Label()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.LabNum = New System.Windows.Forms.Label()
        Me.TextBNum = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ConMenu.SuspendLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon1.BalloonTipText = "123"
        Me.NotifyIcon1.BalloonTipTitle = "456"
        Me.NotifyIcon1.ContextMenuStrip = Me.ConMenu
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        '
        'ConMenu
        '
        Me.ConMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.详细信息ToolStripMenuItem, Me.开机启动ToolStripMenuItem, Me.科中管理系统ToolStripMenuItem, Me.版本特性ToolStripMenuItem, Me.使用须知ToolStripMenuItem, Me.退出程序ToolStripMenuItem})
        Me.ConMenu.Name = "ContextMenu"
        Me.ConMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ConMenu.Size = New System.Drawing.Size(149, 136)
        '
        '详细信息ToolStripMenuItem
        '
        Me.详细信息ToolStripMenuItem.Checked = True
        Me.详细信息ToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.详细信息ToolStripMenuItem.Name = "详细信息ToolStripMenuItem"
        Me.详细信息ToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.详细信息ToolStripMenuItem.Text = "详细信息"
        '
        '开机启动ToolStripMenuItem
        '
        Me.开机启动ToolStripMenuItem.Checked = True
        Me.开机启动ToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.开机启动ToolStripMenuItem.Name = "开机启动ToolStripMenuItem"
        Me.开机启动ToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.开机启动ToolStripMenuItem.Text = "开机启动"
        '
        '科中管理系统ToolStripMenuItem
        '
        Me.科中管理系统ToolStripMenuItem.Name = "科中管理系统ToolStripMenuItem"
        Me.科中管理系统ToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.科中管理系统ToolStripMenuItem.Text = "科中管理系统"
        '
        '版本特性ToolStripMenuItem
        '
        Me.版本特性ToolStripMenuItem.Name = "版本特性ToolStripMenuItem"
        Me.版本特性ToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.版本特性ToolStripMenuItem.Text = "版本特性"
        '
        '使用须知ToolStripMenuItem
        '
        Me.使用须知ToolStripMenuItem.Name = "使用须知ToolStripMenuItem"
        Me.使用须知ToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.使用须知ToolStripMenuItem.Text = "使用须知"
        '
        '退出程序ToolStripMenuItem
        '
        Me.退出程序ToolStripMenuItem.Name = "退出程序ToolStripMenuItem"
        Me.退出程序ToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.退出程序ToolStripMenuItem.Text = "退出程序"
        '
        'LabName
        '
        Me.LabName.AutoSize = True
        Me.LabName.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LabName.Location = New System.Drawing.Point(31, 15)
        Me.LabName.Name = "LabName"
        Me.LabName.Size = New System.Drawing.Size(41, 12)
        Me.LabName.TabIndex = 9
        Me.LabName.Text = "姓名："
        Me.LabName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBName
        '
        Me.TextBName.Enabled = False
        Me.TextBName.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBName.Location = New System.Drawing.Point(83, 12)
        Me.TextBName.Name = "TextBName"
        Me.TextBName.Size = New System.Drawing.Size(115, 21)
        Me.TextBName.TabIndex = 0
        Me.TextBName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBTeam
        '
        Me.TextBTeam.Enabled = False
        Me.TextBTeam.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBTeam.Location = New System.Drawing.Point(83, 66)
        Me.TextBTeam.Name = "TextBTeam"
        Me.TextBTeam.Size = New System.Drawing.Size(115, 21)
        Me.TextBTeam.TabIndex = 2
        Me.TextBTeam.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LabTeam
        '
        Me.LabTeam.AutoSize = True
        Me.LabTeam.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LabTeam.Location = New System.Drawing.Point(31, 69)
        Me.LabTeam.Name = "LabTeam"
        Me.LabTeam.Size = New System.Drawing.Size(41, 12)
        Me.LabTeam.TabIndex = 11
        Me.LabTeam.Text = "组别："
        Me.LabTeam.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBPhone
        '
        Me.TextBPhone.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBPhone.Location = New System.Drawing.Point(83, 154)
        Me.TextBPhone.MaxLength = 11
        Me.TextBPhone.Name = "TextBPhone"
        Me.TextBPhone.Size = New System.Drawing.Size(115, 21)
        Me.TextBPhone.TabIndex = 5
        Me.TextBPhone.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LabPhone
        '
        Me.LabPhone.AutoSize = True
        Me.LabPhone.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LabPhone.Location = New System.Drawing.Point(19, 157)
        Me.LabPhone.Name = "LabPhone"
        Me.LabPhone.Size = New System.Drawing.Size(53, 12)
        Me.LabPhone.TabIndex = 14
        Me.LabPhone.Text = "手机号："
        Me.LabPhone.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBGrade
        '
        Me.TextBGrade.Enabled = False
        Me.TextBGrade.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBGrade.Location = New System.Drawing.Point(83, 94)
        Me.TextBGrade.Name = "TextBGrade"
        Me.TextBGrade.Size = New System.Drawing.Size(115, 21)
        Me.TextBGrade.TabIndex = 3
        Me.TextBGrade.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LabGrade
        '
        Me.LabGrade.AutoSize = True
        Me.LabGrade.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LabGrade.Location = New System.Drawing.Point(31, 97)
        Me.LabGrade.Name = "LabGrade"
        Me.LabGrade.Size = New System.Drawing.Size(41, 12)
        Me.LabGrade.TabIndex = 12
        Me.LabGrade.Text = "年级："
        Me.LabGrade.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBMotto
        '
        Me.TextBMotto.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBMotto.Location = New System.Drawing.Point(83, 183)
        Me.TextBMotto.MaxLength = 1000
        Me.TextBMotto.Multiline = True
        Me.TextBMotto.Name = "TextBMotto"
        Me.TextBMotto.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBMotto.Size = New System.Drawing.Size(115, 89)
        Me.TextBMotto.TabIndex = 6
        '
        'LabMotto
        '
        Me.LabMotto.AutoSize = True
        Me.LabMotto.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LabMotto.Location = New System.Drawing.Point(19, 186)
        Me.LabMotto.Name = "LabMotto"
        Me.LabMotto.Size = New System.Drawing.Size(53, 12)
        Me.LabMotto.TabIndex = 15
        Me.LabMotto.Text = "座右铭："
        Me.LabMotto.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ButChange
        '
        Me.ButChange.Location = New System.Drawing.Point(65, 290)
        Me.ButChange.Name = "ButChange"
        Me.ButChange.Size = New System.Drawing.Size(68, 28)
        Me.ButChange.TabIndex = 7
        Me.ButChange.Text = "应用设置"
        Me.ButChange.UseVisualStyleBackColor = True
        '
        'TextBSign
        '
        Me.TextBSign.Enabled = False
        Me.TextBSign.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBSign.Location = New System.Drawing.Point(83, 124)
        Me.TextBSign.Name = "TextBSign"
        Me.TextBSign.Size = New System.Drawing.Size(115, 21)
        Me.TextBSign.TabIndex = 4
        Me.TextBSign.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LabSign
        '
        Me.LabSign.AutoSize = True
        Me.LabSign.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LabSign.Location = New System.Drawing.Point(7, 127)
        Me.LabSign.Name = "LabSign"
        Me.LabSign.Size = New System.Drawing.Size(65, 12)
        Me.LabSign.TabIndex = 13
        Me.LabSign.Text = "签到时间："
        Me.LabSign.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'StatusStrip
        '
        Me.StatusStrip.Location = New System.Drawing.Point(0, 351)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(221, 22)
        Me.StatusStrip.SizingGrip = False
        Me.StatusStrip.TabIndex = 14
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(150, 17)
        Me.ToolStripStatusLabel1.Spring = True
        Me.ToolStripStatusLabel1.Text = "服务器时间："
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(56, 17)
        Me.ToolStripStatusLabel2.Text = "00:00:00"
        Me.ToolStripStatusLabel2.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'LabNum
        '
        Me.LabNum.AutoSize = True
        Me.LabNum.Location = New System.Drawing.Point(31, 42)
        Me.LabNum.Name = "LabNum"
        Me.LabNum.Size = New System.Drawing.Size(41, 12)
        Me.LabNum.TabIndex = 10
        Me.LabNum.Text = "学号："
        Me.LabNum.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBNum
        '
        Me.TextBNum.Enabled = False
        Me.TextBNum.Location = New System.Drawing.Point(83, 39)
        Me.TextBNum.MaxLength = 9
        Me.TextBNum.Name = "TextBNum"
        Me.TextBNum.Size = New System.Drawing.Size(115, 21)
        Me.TextBNum.TabIndex = 1
        Me.TextBNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(81, 336)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 12)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Label1"
        '
        'ErrorProvider
        '
        Me.ErrorProvider.ContainerControl = Me
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(221, 373)
        Me.Controls.Add(Me.TextBNum)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LabNum)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.TextBSign)
        Me.Controls.Add(Me.LabSign)
        Me.Controls.Add(Me.TextBMotto)
        Me.Controls.Add(Me.TextBPhone)
        Me.Controls.Add(Me.TextBGrade)
        Me.Controls.Add(Me.LabMotto)
        Me.Controls.Add(Me.TextBTeam)
        Me.Controls.Add(Me.ButChange)
        Me.Controls.Add(Me.LabGrade)
        Me.Controls.Add(Me.TextBName)
        Me.Controls.Add(Me.LabPhone)
        Me.Controls.Add(Me.LabTeam)
        Me.Controls.Add(Me.LabName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.KeyPreview = True
        Me.Name = "Main"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "科中管理系统_签到"
        Me.TopMost = True
        Me.ConMenu.ResumeLayout(False)
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ConMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents 退出程序ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 使用须知ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 详细信息ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LabName As System.Windows.Forms.Label
    Friend WithEvents TextBName As System.Windows.Forms.TextBox
    Friend WithEvents 开机启动ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TextBTeam As System.Windows.Forms.TextBox
    Friend WithEvents LabTeam As System.Windows.Forms.Label
    Friend WithEvents TextBPhone As System.Windows.Forms.TextBox
    Friend WithEvents LabPhone As System.Windows.Forms.Label
    Friend WithEvents TextBGrade As System.Windows.Forms.TextBox
    Friend WithEvents LabGrade As System.Windows.Forms.Label
    Friend WithEvents TextBMotto As System.Windows.Forms.TextBox
    Friend WithEvents LabMotto As System.Windows.Forms.Label
    Friend WithEvents ButChange As System.Windows.Forms.Button
    Friend WithEvents Timer As System.Windows.Forms.Timer
    Friend WithEvents TextBSign As System.Windows.Forms.TextBox
    Friend WithEvents LabSign As System.Windows.Forms.Label
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents 版本特性ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LabNum As System.Windows.Forms.Label
    Friend WithEvents TextBNum As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider As System.Windows.Forms.ErrorProvider
    Friend WithEvents 科中管理系统ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon

End Class
