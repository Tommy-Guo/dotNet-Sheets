<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.infoLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.spacerLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SpreadSheet1 = New dotNet_Sheets.Sheet()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox1.Font = New System.Drawing.Font("Consolas", 8.25!)
        Me.TextBox1.Location = New System.Drawing.Point(36, 10)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(959, 20)
        Me.TextBox1.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Padding = New System.Windows.Forms.Padding(10)
        Me.Panel1.Size = New System.Drawing.Size(1005, 40)
        Me.Panel1.TabIndex = 3
        '
        'PictureBox1
        '
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Left
        Me.PictureBox1.Image = Global.dotNet_Sheets.My.Resources.Resources._51661
        Me.PictureBox1.Location = New System.Drawing.Point(10, 10)
        Me.PictureBox1.MaximumSize = New System.Drawing.Size(26, 20)
        Me.PictureBox1.MinimumSize = New System.Drawing.Size(26, 20)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Padding = New System.Windows.Forms.Padding(0, 0, 6, 0)
        Me.PictureBox1.Size = New System.Drawing.Size(26, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 14
        Me.PictureBox1.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.AutoScroll = True
        Me.Panel2.AutoSize = True
        Me.Panel2.Controls.Add(Me.SpreadSheet1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 41)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1005, 539)
        Me.Panel2.TabIndex = 0
        '
        'infoLabel
        '
        Me.infoLabel.Name = "infoLabel"
        Me.infoLabel.Size = New System.Drawing.Size(231, 17)
        Me.infoLabel.Text = "Grid Size: {100, 100}   |   Selected Cell: None"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.spacerLabel, Me.infoLabel})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 580)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1005, 22)
        Me.StatusStrip1.TabIndex = 4
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'spacerLabel
        '
        Me.spacerLabel.Name = "spacerLabel"
        Me.spacerLabel.Size = New System.Drawing.Size(759, 17)
        Me.spacerLabel.Spring = True
        Me.spacerLabel.Text = " "
        '
        'SpreadSheet1
        '
        Me.SpreadSheet1.ActiveCell = Nothing
        Me.SpreadSheet1.CellSize = New System.Drawing.Size(80, 30)
        Me.SpreadSheet1.FormulaBox = Me.TextBox1
        Me.SpreadSheet1.GridSize = New System.Drawing.Size(100, 100)
        Me.SpreadSheet1.Location = New System.Drawing.Point(0, 0)
        Me.SpreadSheet1.Name = "SpreadSheet1"
        Me.SpreadSheet1.Size = New System.Drawing.Size(8081, 3031)
        Me.SpreadSheet1.StatusLabel = Me.infoLabel
        Me.SpreadSheet1.TabIndex = 2
        Me.SpreadSheet1.Text = "SpreadSheet1"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(8)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1005, 1)
        Me.Panel3.TabIndex = 5
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1005, 602)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Name = "Form1"
        Me.Text = "dotNet Sheets"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents SpreadSheet1 As dotNet_Sheets.Sheet
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents spacerLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents infoLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel

End Class
