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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.StartToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NeuePartieToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PartieLadenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CasinoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BetVoyagerNonZeroToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EinstellungToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AboutToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.Version101ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Button1 = New System.Windows.Forms.Button
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.DarkGreen
        Me.MenuStrip1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 0)
        Me.MenuStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StartToolStripMenuItem, Me.CasinoToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.MenuStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
        Me.MenuStrip1.Size = New System.Drawing.Size(338, 23)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'StartToolStripMenuItem
        '
        Me.StartToolStripMenuItem.BackColor = System.Drawing.Color.DarkGreen
        Me.StartToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.StartToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NeuePartieToolStripMenuItem, Me.PartieLadenToolStripMenuItem})
        Me.StartToolStripMenuItem.ForeColor = System.Drawing.Color.Silver
        Me.StartToolStripMenuItem.Name = "StartToolStripMenuItem"
        Me.StartToolStripMenuItem.Size = New System.Drawing.Size(39, 19)
        Me.StartToolStripMenuItem.Text = "File"
        '
        'NeuePartieToolStripMenuItem
        '
        Me.NeuePartieToolStripMenuItem.BackColor = System.Drawing.Color.DarkGreen
        Me.NeuePartieToolStripMenuItem.ForeColor = System.Drawing.Color.Silver
        Me.NeuePartieToolStripMenuItem.Name = "NeuePartieToolStripMenuItem"
        Me.NeuePartieToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.NeuePartieToolStripMenuItem.Text = "Neue Partie beginnen"
        '
        'PartieLadenToolStripMenuItem
        '
        Me.PartieLadenToolStripMenuItem.BackColor = System.Drawing.Color.DarkGreen
        Me.PartieLadenToolStripMenuItem.ForeColor = System.Drawing.Color.Silver
        Me.PartieLadenToolStripMenuItem.Name = "PartieLadenToolStripMenuItem"
        Me.PartieLadenToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.PartieLadenToolStripMenuItem.Text = "Partie laden"
        '
        'CasinoToolStripMenuItem
        '
        Me.CasinoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BetVoyagerNonZeroToolStripMenuItem})
        Me.CasinoToolStripMenuItem.ForeColor = System.Drawing.Color.Silver
        Me.CasinoToolStripMenuItem.Name = "CasinoToolStripMenuItem"
        Me.CasinoToolStripMenuItem.Size = New System.Drawing.Size(59, 19)
        Me.CasinoToolStripMenuItem.Text = "Casino"
        '
        'BetVoyagerNonZeroToolStripMenuItem
        '
        Me.BetVoyagerNonZeroToolStripMenuItem.Checked = True
        Me.BetVoyagerNonZeroToolStripMenuItem.CheckOnClick = True
        Me.BetVoyagerNonZeroToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BetVoyagerNonZeroToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EinstellungToolStripMenuItem})
        Me.BetVoyagerNonZeroToolStripMenuItem.Image = Global.P__Bot.My.Resources.Resources.bv
        Me.BetVoyagerNonZeroToolStripMenuItem.Name = "BetVoyagerNonZeroToolStripMenuItem"
        Me.BetVoyagerNonZeroToolStripMenuItem.Size = New System.Drawing.Size(189, 22)
        Me.BetVoyagerNonZeroToolStripMenuItem.Text = "BetVoyager Non Zero"
        '
        'EinstellungToolStripMenuItem
        '
        Me.EinstellungToolStripMenuItem.Name = "EinstellungToolStripMenuItem"
        Me.EinstellungToolStripMenuItem.Size = New System.Drawing.Size(136, 22)
        Me.EinstellungToolStripMenuItem.Text = "Einstellung"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem1, Me.Version101ToolStripMenuItem})
        Me.AboutToolStripMenuItem.ForeColor = System.Drawing.Color.Silver
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Overflow = System.Windows.Forms.ToolStripItemOverflow.Always
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(39, 19)
        Me.AboutToolStripMenuItem.Text = "Info"
        '
        'AboutToolStripMenuItem1
        '
        Me.AboutToolStripMenuItem1.BackColor = System.Drawing.Color.DarkGreen
        Me.AboutToolStripMenuItem1.ForeColor = System.Drawing.Color.Silver
        Me.AboutToolStripMenuItem1.Name = "AboutToolStripMenuItem1"
        Me.AboutToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.AboutToolStripMenuItem1.Text = "About P²"
        '
        'Version101ToolStripMenuItem
        '
        Me.Version101ToolStripMenuItem.BackColor = System.Drawing.Color.DarkGreen
        Me.Version101ToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.None
        Me.Version101ToolStripMenuItem.ForeColor = System.Drawing.Color.Silver
        Me.Version101ToolStripMenuItem.Name = "Version101ToolStripMenuItem"
        Me.Version101ToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.Version101ToolStripMenuItem.Text = "Version 1.01"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.DarkGreen
        Me.Button1.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.YellowGreen
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Image = Global.P__Bot.My.Resources.Resources.Delete
        Me.Button1.Location = New System.Drawing.Point(310, 0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(28, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.UseVisualStyleBackColor = False
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.DimGray
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 14
        Me.ListBox1.Location = New System.Drawing.Point(0, 26)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(71, 392)
        Me.ListBox1.TabIndex = 2
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(338, 769)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Silver
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "P²"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StartToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NeuePartieToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PartieLadenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Version101ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CasinoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BetVoyagerNonZeroToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EinstellungToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox

End Class
