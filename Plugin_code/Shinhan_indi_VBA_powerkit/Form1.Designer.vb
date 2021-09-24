<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.AutoStatus = New System.Windows.Forms.Label()
        Me.oShin_balance = New AxGIEXPERTCONTROLLib.AxGiExpertControl()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.oShin_PriceAmount = New AxGIEXPERTCONTROLLib.AxGiExpertControl()
        Me.oShin_name = New AxGIEXPERTCONTROLLib.AxGiExpertControl()
        Me.oShin_Priceonly = New AxGIEXPERTCONTROLLib.AxGiExpertControl()
        Me.oShin_favname = New AxGIEXPERTCONTROLLib.AxGiExpertControl()
        Me.oShin_buysell = New AxGIEXPERTCONTROLLib.AxGiExpertControl()
        Me.G_use = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_Status = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_Box = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.G_Read = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_Name = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_Write = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_Own = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_GoodPrice = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_BadSell = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_BadPer = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_GoodSell = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G_GoodPer = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        CType(Me.oShin_balance, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.oShin_PriceAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.oShin_name, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.oShin_Priceonly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.oShin_favname, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.oShin_buysell, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(303, 599)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(141, 25)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Launch Excel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.AutoStatus)
        Me.GroupBox1.Controls.Add(Me.oShin_balance)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(747, 575)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Debug Message"
        '
        'AutoStatus
        '
        Me.AutoStatus.AutoSize = True
        Me.AutoStatus.ForeColor = System.Drawing.SystemColors.ControlText
        Me.AutoStatus.Location = New System.Drawing.Point(122, 118)
        Me.AutoStatus.Name = "AutoStatus"
        Me.AutoStatus.Size = New System.Drawing.Size(57, 12)
        Me.AutoStatus.TabIndex = 11
        Me.AutoStatus.Text = "Unknown"
        '
        'oShin_balance
        '
        Me.oShin_balance.Enabled = True
        Me.oShin_balance.Location = New System.Drawing.Point(1878, 1152)
        Me.oShin_balance.Name = "oShin_balance"
        Me.oShin_balance.OcxState = CType(resources.GetObject("oShin_balance.OcxState"), System.Windows.Forms.AxHost.State)
        Me.oShin_balance.Size = New System.Drawing.Size(136, 50)
        Me.oShin_balance.TabIndex = 10
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.ControlDark
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeight = 20
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.G_use, Me.G_Status, Me.G_Box, Me.G_Read, Me.G_Name, Me.G_Write, Me.G_Amount, Me.G_Own, Me.G_GoodPrice, Me.G_BadSell, Me.G_BadPer, Me.G_GoodSell, Me.G_GoodPer})
        Me.DataGridView1.Location = New System.Drawing.Point(21, 138)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(1)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.RowTemplate.Height = 20
        Me.DataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.DataGridView1.Size = New System.Drawing.Size(707, 421)
        Me.DataGridView1.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(20, 118)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 12)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Autodeal status : "
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.ForeColor = System.Drawing.SystemColors.WindowFrame
        Me.Label1.Location = New System.Drawing.Point(21, 20)
        Me.Label1.Multiline = True
        Me.Label1.Name = "Label1"
        Me.Label1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.Label1.Size = New System.Drawing.Size(707, 90)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Debug Message Displays Here"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(23, 605)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "LastCalledFunction"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(459, 599)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(141, 25)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "run Macros"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(618, 599)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(141, 25)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Launch shinhan module"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'oShin_PriceAmount
        '
        Me.oShin_PriceAmount.Enabled = True
        Me.oShin_PriceAmount.Location = New System.Drawing.Point(4819, 2759)
        Me.oShin_PriceAmount.Name = "oShin_PriceAmount"
        Me.oShin_PriceAmount.OcxState = CType(resources.GetObject("oShin_PriceAmount.OcxState"), System.Windows.Forms.AxHost.State)
        Me.oShin_PriceAmount.Size = New System.Drawing.Size(136, 50)
        Me.oShin_PriceAmount.TabIndex = 3
        '
        'oShin_name
        '
        Me.oShin_name.Enabled = True
        Me.oShin_name.Location = New System.Drawing.Point(3923, 2248)
        Me.oShin_name.Name = "oShin_name"
        Me.oShin_name.OcxState = CType(resources.GetObject("oShin_name.OcxState"), System.Windows.Forms.AxHost.State)
        Me.oShin_name.Size = New System.Drawing.Size(136, 50)
        Me.oShin_name.TabIndex = 5
        '
        'oShin_Priceonly
        '
        Me.oShin_Priceonly.Enabled = True
        Me.oShin_Priceonly.Location = New System.Drawing.Point(3663, 2109)
        Me.oShin_Priceonly.Name = "oShin_Priceonly"
        Me.oShin_Priceonly.OcxState = CType(resources.GetObject("oShin_Priceonly.OcxState"), System.Windows.Forms.AxHost.State)
        Me.oShin_Priceonly.Size = New System.Drawing.Size(136, 50)
        Me.oShin_Priceonly.TabIndex = 6
        '
        'oShin_favname
        '
        Me.oShin_favname.Enabled = True
        Me.oShin_favname.Location = New System.Drawing.Point(4191, 2395)
        Me.oShin_favname.Name = "oShin_favname"
        Me.oShin_favname.OcxState = CType(resources.GetObject("oShin_favname.OcxState"), System.Windows.Forms.AxHost.State)
        Me.oShin_favname.Size = New System.Drawing.Size(136, 50)
        Me.oShin_favname.TabIndex = 7
        '
        'oShin_buysell
        '
        Me.oShin_buysell.Enabled = True
        Me.oShin_buysell.Location = New System.Drawing.Point(1623, 1069)
        Me.oShin_buysell.Name = "oShin_buysell"
        Me.oShin_buysell.OcxState = CType(resources.GetObject("oShin_buysell.OcxState"), System.Windows.Forms.AxHost.State)
        Me.oShin_buysell.Size = New System.Drawing.Size(136, 50)
        Me.oShin_buysell.TabIndex = 8
        '
        'G_use
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.G_use.DefaultCellStyle = DataGridViewCellStyle2
        Me.G_use.HeaderText = "◎"
        Me.G_use.Name = "G_use"
        Me.G_use.Width = 33
        '
        'G_Status
        '
        Me.G_Status.HeaderText = "결과"
        Me.G_Status.Name = "G_Status"
        Me.G_Status.Width = 35
        '
        'G_Box
        '
        Me.G_Box.HeaderText = "사용"
        Me.G_Box.Name = "G_Box"
        Me.G_Box.Width = 35
        '
        'G_Read
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.G_Read.DefaultCellStyle = DataGridViewCellStyle3
        Me.G_Read.HeaderText = "종목C"
        Me.G_Read.Name = "G_Read"
        Me.G_Read.Width = 50
        '
        'G_Name
        '
        Me.G_Name.HeaderText = "이름"
        Me.G_Name.Name = "G_Name"
        Me.G_Name.Width = 90
        '
        'G_Write
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.Padding = New System.Windows.Forms.Padding(0, 1, 0, 0)
        Me.G_Write.DefaultCellStyle = DataGridViewCellStyle4
        Me.G_Write.HeaderText = "현재가"
        Me.G_Write.Name = "G_Write"
        Me.G_Write.Width = 70
        '
        'G_Amount
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.Padding = New System.Windows.Forms.Padding(0, 1, 0, 0)
        Me.G_Amount.DefaultCellStyle = DataGridViewCellStyle5
        Me.G_Amount.HeaderText = "거래량"
        Me.G_Amount.Name = "G_Amount"
        Me.G_Amount.Width = 65
        '
        'G_Own
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.G_Own.DefaultCellStyle = DataGridViewCellStyle6
        Me.G_Own.HeaderText = "보유"
        Me.G_Own.Name = "G_Own"
        Me.G_Own.Width = 50
        '
        'G_GoodPrice
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.Padding = New System.Windows.Forms.Padding(0, 1, 0, 0)
        Me.G_GoodPrice.DefaultCellStyle = DataGridViewCellStyle7
        Me.G_GoodPrice.HeaderText = "제비용단가"
        Me.G_GoodPrice.Name = "G_GoodPrice"
        Me.G_GoodPrice.Width = 70
        '
        'G_BadSell
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle8.Padding = New System.Windows.Forms.Padding(0, 1, 0, 0)
        Me.G_BadSell.DefaultCellStyle = DataGridViewCellStyle8
        Me.G_BadSell.HeaderText = "손절가"
        Me.G_BadSell.Name = "G_BadSell"
        Me.G_BadSell.Width = 70
        '
        'G_BadPer
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.G_BadPer.DefaultCellStyle = DataGridViewCellStyle9
        Me.G_BadPer.HeaderText = "%"
        Me.G_BadPer.Name = "G_BadPer"
        Me.G_BadPer.Width = 33
        '
        'G_GoodSell
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle10.Padding = New System.Windows.Forms.Padding(0, 1, 0, 0)
        Me.G_GoodSell.DefaultCellStyle = DataGridViewCellStyle10
        Me.G_GoodSell.HeaderText = "목표가"
        Me.G_GoodSell.Name = "G_GoodSell"
        Me.G_GoodSell.Width = 70
        '
        'G_GoodPer
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.G_GoodPer.DefaultCellStyle = DataGridViewCellStyle11
        Me.G_GoodPer.HeaderText = "%"
        Me.G_GoodPer.Name = "G_GoodPer"
        Me.G_GoodPer.Width = 33
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(771, 639)
        Me.Controls.Add(Me.oShin_buysell)
        Me.Controls.Add(Me.oShin_favname)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.oShin_Priceonly)
        Me.Controls.Add(Me.oShin_name)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.oShin_PriceAmount)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form1"
        Me.Text = "FastVBA Plugin"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.oShin_balance, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.oShin_PriceAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.oShin_name, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.oShin_Priceonly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.oShin_favname, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.oShin_buysell, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Button2 As Button
    Friend WithEvents oShin_PriceAmount As AxGIEXPERTCONTROLLib.AxGiExpertControl
    Friend WithEvents Button3 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents oShin_name As AxGIEXPERTCONTROLLib.AxGiExpertControl
    Friend WithEvents oShin_Priceonly As AxGIEXPERTCONTROLLib.AxGiExpertControl
    Friend WithEvents oShin_favname As AxGIEXPERTCONTROLLib.AxGiExpertControl
    Friend WithEvents Label1 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents oShin_balance As AxGIEXPERTCONTROLLib.AxGiExpertControl
    Friend WithEvents oShin_buysell As AxGIEXPERTCONTROLLib.AxGiExpertControl
    Friend WithEvents AutoStatus As Label
    Friend WithEvents G_use As DataGridViewTextBoxColumn
    Friend WithEvents G_Status As DataGridViewTextBoxColumn
    Friend WithEvents G_Box As DataGridViewCheckBoxColumn
    Friend WithEvents G_Read As DataGridViewTextBoxColumn
    Friend WithEvents G_Name As DataGridViewTextBoxColumn
    Friend WithEvents G_Write As DataGridViewTextBoxColumn
    Friend WithEvents G_Amount As DataGridViewTextBoxColumn
    Friend WithEvents G_Own As DataGridViewTextBoxColumn
    Friend WithEvents G_GoodPrice As DataGridViewTextBoxColumn
    Friend WithEvents G_BadSell As DataGridViewTextBoxColumn
    Friend WithEvents G_BadPer As DataGridViewTextBoxColumn
    Friend WithEvents G_GoodSell As DataGridViewTextBoxColumn
    Friend WithEvents G_GoodPer As DataGridViewTextBoxColumn
End Class
