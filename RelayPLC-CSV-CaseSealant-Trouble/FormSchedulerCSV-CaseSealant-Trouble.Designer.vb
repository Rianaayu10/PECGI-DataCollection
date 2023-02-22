<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormSchedulerCSV_CaseSealant_Trouble
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormSchedulerCSV_CaseSealant_Trouble))
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.btnConfig = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnManual = New System.Windows.Forms.Button()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.txtMsg = New System.Windows.Forms.TextBox()
        Me.grpMsg = New System.Windows.Forms.GroupBox()
        Me.TimerProcess = New System.Windows.Forms.Timer(Me.components)
        Me.timerCurr = New System.Windows.Forms.Timer(Me.components)
        Me.stpStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.lblCurrDate = New System.Windows.Forms.Label()
        Me.lblCurrTime = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pic1 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.grid = New C1.Win.C1FlexGrid.C1FlexGrid()
        Me.grpMsg.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.pic1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "RELAY PLC CSV TROUBLE (TAMPING)"
        Me.NotifyIcon1.Visible = True
        '
        'btnConfig
        '
        Me.btnConfig.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnConfig.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConfig.Image = CType(resources.GetObject("btnConfig.Image"), System.Drawing.Image)
        Me.btnConfig.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnConfig.Location = New System.Drawing.Point(105, 612)
        Me.btnConfig.Name = "btnConfig"
        Me.btnConfig.Size = New System.Drawing.Size(88, 28)
        Me.btnConfig.TabIndex = 74
        Me.btnConfig.Text = "C&onfig"
        Me.btnConfig.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Image = CType(resources.GetObject("btnClose.Image"), System.Drawing.Image)
        Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose.Location = New System.Drawing.Point(11, 612)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(88, 28)
        Me.btnClose.TabIndex = 73
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnManual
        '
        Me.btnManual.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnManual.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManual.Image = CType(resources.GetObject("btnManual.Image"), System.Drawing.Image)
        Me.btnManual.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnManual.Location = New System.Drawing.Point(652, 612)
        Me.btnManual.Name = "btnManual"
        Me.btnManual.Size = New System.Drawing.Size(127, 28)
        Me.btnManual.TabIndex = 72
        Me.btnManual.Text = "     &Manual Process"
        Me.btnManual.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStop.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Image = CType(resources.GetObject("btnStop.Image"), System.Drawing.Image)
        Me.btnStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnStop.Location = New System.Drawing.Point(785, 612)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(88, 28)
        Me.btnStop.TabIndex = 71
        Me.btnStop.Text = "S&top"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStart.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Image = CType(resources.GetObject("btnStart.Image"), System.Drawing.Image)
        Me.btnStart.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnStart.Location = New System.Drawing.Point(879, 612)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(88, 28)
        Me.btnStart.TabIndex = 70
        Me.btnStart.Text = "&Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'txtMsg
        '
        Me.txtMsg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMsg.BackColor = System.Drawing.Color.LightSteelBlue
        Me.txtMsg.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtMsg.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMsg.ForeColor = System.Drawing.Color.Blue
        Me.txtMsg.Location = New System.Drawing.Point(13, 18)
        Me.txtMsg.Multiline = True
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.ReadOnly = True
        Me.txtMsg.Size = New System.Drawing.Size(924, 25)
        Me.txtMsg.TabIndex = 1
        Me.txtMsg.TabStop = False
        Me.txtMsg.Text = "Message"
        Me.txtMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'grpMsg
        '
        Me.grpMsg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpMsg.Controls.Add(Me.txtMsg)
        Me.grpMsg.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMsg.Location = New System.Drawing.Point(8, 542)
        Me.grpMsg.Name = "grpMsg"
        Me.grpMsg.Size = New System.Drawing.Size(959, 55)
        Me.grpMsg.TabIndex = 69
        Me.grpMsg.TabStop = False
        '
        'TimerProcess
        '
        Me.TimerProcess.Interval = 15000
        '
        'timerCurr
        '
        '
        'stpStatus
        '
        Me.stpStatus.BackColor = System.Drawing.Color.White
        Me.stpStatus.Name = "stpStatus"
        Me.stpStatus.Size = New System.Drawing.Size(967, 17)
        Me.stpStatus.Spring = True
        Me.stpStatus.Text = "Status"
        Me.stpStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.stpStatus})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 663)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(982, 22)
        Me.StatusStrip1.TabIndex = 67
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'lblCurrDate
        '
        Me.lblCurrDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCurrDate.BackColor = System.Drawing.Color.White
        Me.lblCurrDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrDate.Location = New System.Drawing.Point(667, 39)
        Me.lblCurrDate.Name = "lblCurrDate"
        Me.lblCurrDate.Size = New System.Drawing.Size(302, 20)
        Me.lblCurrDate.TabIndex = 13
        Me.lblCurrDate.Text = "NAMA HARI DALAM INGGRIS , 08 Feb 2021"
        Me.lblCurrDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurrTime
        '
        Me.lblCurrTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCurrTime.BackColor = System.Drawing.Color.White
        Me.lblCurrTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrTime.Location = New System.Drawing.Point(798, 5)
        Me.lblCurrTime.Name = "lblCurrTime"
        Me.lblCurrTime.Size = New System.Drawing.Size(171, 29)
        Me.lblCurrTime.TabIndex = 12
        Me.lblCurrTime.Text = "17:05:01"
        Me.lblCurrTime.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblVersion.Location = New System.Drawing.Point(102, 37)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(53, 13)
        Me.lblVersion.TabIndex = 2
        Me.lblVersion.Text = "Ver. 1.0.0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DarkBlue
        Me.Label1.Location = New System.Drawing.Point(102, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(389, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "RELAY PLC CSV TROUBLE (CaseSealant)"
        '
        'pic1
        '
        Me.pic1.Image = CType(resources.GetObject("pic1.Image"), System.Drawing.Image)
        Me.pic1.Location = New System.Drawing.Point(10, 3)
        Me.pic1.Name = "pic1"
        Me.pic1.Size = New System.Drawing.Size(78, 61)
        Me.pic1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pic1.TabIndex = 0
        Me.pic1.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.lblCurrDate)
        Me.Panel1.Controls.Add(Me.lblCurrTime)
        Me.Panel1.Controls.Add(Me.lblVersion)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.pic1)
        Me.Panel1.Location = New System.Drawing.Point(-2, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(987, 67)
        Me.Panel1.TabIndex = 66
        '
        'grid
        '
        Me.grid.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None
        Me.grid.AllowEditing = False
        Me.grid.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None
        Me.grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grid.AutoClipboard = True
        Me.grid.AutoResize = True
        Me.grid.ColumnInfo = "3,0,0,0,0,100,Columns:1{Width:118;}" & Global.Microsoft.VisualBasic.ChrW(9) & "2{Width:120;}" & Global.Microsoft.VisualBasic.ChrW(9)
        Me.grid.FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None
        Me.grid.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grid.HighLight = C1.Win.C1FlexGrid.HighLightEnum.Never
        Me.grid.Location = New System.Drawing.Point(11, 95)
        Me.grid.Name = "grid"
        Me.grid.Rows.Count = 5
        Me.grid.Rows.DefaultSize = 20
        Me.grid.Size = New System.Drawing.Size(956, 441)
        Me.grid.StyleInfo = resources.GetString("grid.StyleInfo")
        Me.grid.TabIndex = 68
        '
        'FormSchedulerCSV_CaseSealant_Trouble
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(982, 685)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnConfig)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnManual)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.grpMsg)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.grid)
        Me.Name = "FormSchedulerCSV_CaseSealant_Trouble"
        Me.Text = "PT. Panasonic Gobel Energy Indonesia (PECGI)"
        Me.grpMsg.ResumeLayout(False)
        Me.grpMsg.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.pic1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents NotifyIcon1 As NotifyIcon
    Friend WithEvents btnConfig As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents btnManual As Button
    Friend WithEvents btnStop As Button
    Friend WithEvents btnStart As Button
    Public WithEvents txtMsg As TextBox
    Public WithEvents grpMsg As GroupBox
    Friend WithEvents TimerProcess As Timer
    Private WithEvents timerCurr As Timer
    Friend WithEvents stpStatus As ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As StatusStrip
    Private WithEvents lblCurrDate As Label
    Private WithEvents lblCurrTime As Label
    Friend WithEvents lblVersion As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents pic1 As PictureBox
    Friend WithEvents Panel1 As Panel
    Public WithEvents grid As C1.Win.C1FlexGrid.C1FlexGrid
End Class
