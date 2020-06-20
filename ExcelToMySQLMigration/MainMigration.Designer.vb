<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMigration
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
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.MyDataGridView_Trace = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MysqlServerSettingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtStatus = New System.Windows.Forms.Label()
        Me.txtHost = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdMigrate = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtTotalRows = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.txtTableName = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        CType(Me.MyDataGridView_Trace, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MyDataGridView_Trace
        '
        Me.MyDataGridView_Trace.AllowUserToAddRows = False
        Me.MyDataGridView_Trace.AllowUserToDeleteRows = False
        Me.MyDataGridView_Trace.AllowUserToResizeColumns = False
        Me.MyDataGridView_Trace.AllowUserToResizeRows = False
        Me.MyDataGridView_Trace.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MyDataGridView_Trace.BackgroundColor = System.Drawing.Color.White
        Me.MyDataGridView_Trace.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.MyDataGridView_Trace.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle10.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.MyDataGridView_Trace.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle10
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Segoe UI", 8.25!)
        DataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.MyDataGridView_Trace.DefaultCellStyle = DataGridViewCellStyle11
        Me.MyDataGridView_Trace.GridColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.MyDataGridView_Trace.Location = New System.Drawing.Point(359, 34)
        Me.MyDataGridView_Trace.Margin = New System.Windows.Forms.Padding(2)
        Me.MyDataGridView_Trace.MultiSelect = False
        Me.MyDataGridView_Trace.Name = "MyDataGridView_Trace"
        Me.MyDataGridView_Trace.ReadOnly = True
        Me.MyDataGridView_Trace.RowHeadersVisible = False
        Me.MyDataGridView_Trace.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        DataGridViewCellStyle12.SelectionBackColor = System.Drawing.Color.CornflowerBlue
        DataGridViewCellStyle12.SelectionForeColor = System.Drawing.Color.White
        Me.MyDataGridView_Trace.RowsDefaultCellStyle = DataGridViewCellStyle12
        Me.MyDataGridView_Trace.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.MyDataGridView_Trace.Size = New System.Drawing.Size(669, 615)
        Me.MyDataGridView_Trace.TabIndex = 396
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Image = Global.ExcelToMySQLMigration.My.Resources.Resources.document_excel_table
        Me.Button1.Location = New System.Drawing.Point(316, 52)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(25, 23)
        Me.Button1.TabIndex = 397
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(12, 52)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(301, 23)
        Me.TextBox1.TabIndex = 398
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MysqlServerSettingToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1031, 24)
        Me.MenuStrip1.TabIndex = 399
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MysqlServerSettingToolStripMenuItem
        '
        Me.MysqlServerSettingToolStripMenuItem.Name = "MysqlServerSettingToolStripMenuItem"
        Me.MysqlServerSettingToolStripMenuItem.Size = New System.Drawing.Size(137, 20)
        Me.MysqlServerSettingToolStripMenuItem.Text = "MySQL Server Settings"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 84)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 15)
        Me.Label1.TabIndex = 400
        Me.Label1.Text = "Connection Status:"
        '
        'txtStatus
        '
        Me.txtStatus.AutoSize = True
        Me.txtStatus.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.Location = New System.Drawing.Point(124, 85)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(77, 15)
        Me.txtStatus.TabIndex = 401
        Me.txtStatus.Text = "CONNECTED"
        '
        'txtHost
        '
        Me.txtHost.AutoSize = True
        Me.txtHost.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHost.Location = New System.Drawing.Point(124, 103)
        Me.txtHost.Name = "txtHost"
        Me.txtHost.Size = New System.Drawing.Size(88, 15)
        Me.txtHost.TabIndex = 403
        Me.txtHost.Text = "192.168.254.119"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(52, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 15)
        Me.Label4.TabIndex = 402
        Me.Label4.Text = "Server Host:"
        '
        'cmdMigrate
        '
        Me.cmdMigrate.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMigrate.Location = New System.Drawing.Point(92, 213)
        Me.cmdMigrate.Name = "cmdMigrate"
        Me.cmdMigrate.Size = New System.Drawing.Size(221, 40)
        Me.cmdMigrate.TabIndex = 404
        Me.cmdMigrate.Text = "Start Data Transfer"
        Me.cmdMigrate.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(154, 15)
        Me.Label2.TabIndex = 405
        Me.Label2.Text = "Please load correct excel file"
        '
        'txtTotalRows
        '
        Me.txtTotalRows.AutoSize = True
        Me.txtTotalRows.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalRows.Location = New System.Drawing.Point(124, 123)
        Me.txtTotalRows.Name = "txtTotalRows"
        Me.txtTotalRows.Size = New System.Drawing.Size(0, 15)
        Me.txtTotalRows.TabIndex = 407
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(54, 123)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(67, 15)
        Me.Label5.TabIndex = 406
        Me.Label5.Text = "Total Rows:"
        '
        'pb1
        '
        Me.pb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb1.Location = New System.Drawing.Point(1, 654)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(1027, 23)
        Me.pb1.TabIndex = 408
        '
        'lbl1
        '
        Me.lbl1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbl1.AutoSize = True
        Me.lbl1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.Location = New System.Drawing.Point(1, 637)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(189, 15)
        Me.lbl1.TabIndex = 409
        Me.lbl1.Text = "Converting from Excel to Database"
        '
        'txtTableName
        '
        Me.txtTableName.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtTableName.FormattingEnabled = True
        Me.txtTableName.Location = New System.Drawing.Point(92, 159)
        Me.txtTableName.Name = "txtTableName"
        Me.txtTableName.Size = New System.Drawing.Size(221, 23)
        Me.txtTableName.TabIndex = 412
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(15, 163)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 15)
        Me.Label3.TabIndex = 413
        Me.Label3.Text = "Target Table"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.CheckBox1.Location = New System.Drawing.Point(92, 188)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(142, 19)
        Me.CheckBox1.TabIndex = 414
        Me.CheckBox1.Text = "Drop and Insert Query"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'MainMigration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1031, 679)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtTableName)
        Me.Controls.Add(Me.lbl1)
        Me.Controls.Add(Me.pb1)
        Me.Controls.Add(Me.txtTotalRows)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmdMigrate)
        Me.Controls.Add(Me.txtHost)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.MyDataGridView_Trace)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainMigration"
        Me.Text = "DaddyRoger Migration Tool"
        CType(Me.MyDataGridView_Trace, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MyDataGridView_Trace As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents MysqlServerSettingToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtStatus As System.Windows.Forms.Label
    Friend WithEvents txtHost As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdMigrate As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtTotalRows As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents txtTableName As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox

End Class
