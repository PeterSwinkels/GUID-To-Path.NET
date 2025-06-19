Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class InterfaceWindow
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(InterfaceWindow))
        Me.GUIDListBox = New System.Windows.Forms.TextBox()
        Me.GUIDsLabel = New System.Windows.Forms.Label()
        Me.ResultsBox = New System.Windows.Forms.TextBox()
        Me.ResultsLabel = New System.Windows.Forms.Label()
        Me.SearchButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'GUIDListBox
        '
        Me.GUIDListBox.BackColor = System.Drawing.SystemColors.Window
        Me.GUIDListBox.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GUIDListBox.Location = New System.Drawing.Point(12, 29)
        Me.GUIDListBox.Multiline = True
        Me.GUIDListBox.Name = "GUIDListBox"
        Me.GUIDListBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.GUIDListBox.Size = New System.Drawing.Size(880, 221)
        Me.GUIDListBox.TabIndex = 0
        '
        'GUIDsLabel
        '
        Me.GUIDsLabel.AutoSize = True
        Me.GUIDsLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GUIDsLabel.Location = New System.Drawing.Point(12, 9)
        Me.GUIDsLabel.Name = "GUIDsLabel"
        Me.GUIDsLabel.Size = New System.Drawing.Size(48, 13)
        Me.GUIDsLabel.TabIndex = 1
        Me.GUIDsLabel.Text = "GUIDs:"
        '
        'ResultsBox
        '
        Me.ResultsBox.BackColor = System.Drawing.SystemColors.Window
        Me.ResultsBox.Font = New System.Drawing.Font("Consolas", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ResultsBox.Location = New System.Drawing.Point(15, 310)
        Me.ResultsBox.Multiline = True
        Me.ResultsBox.Name = "ResultsBox"
        Me.ResultsBox.ReadOnly = True
        Me.ResultsBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.ResultsBox.Size = New System.Drawing.Size(880, 223)
        Me.ResultsBox.TabIndex = 2
        '
        'ResultsLabel
        '
        Me.ResultsLabel.AutoSize = True
        Me.ResultsLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ResultsLabel.Location = New System.Drawing.Point(12, 281)
        Me.ResultsLabel.Name = "ResultsLabel"
        Me.ResultsLabel.Size = New System.Drawing.Size(53, 13)
        Me.ResultsLabel.TabIndex = 3
        Me.ResultsLabel.Text = "Results:"
        '
        'SearchButton
        '
        Me.SearchButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.SearchButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SearchButton.Location = New System.Drawing.Point(798, 271)
        Me.SearchButton.Name = "SearchButton"
        Me.SearchButton.Size = New System.Drawing.Size(75, 23)
        Me.SearchButton.TabIndex = 1
        Me.SearchButton.Text = "&Search"
        Me.SearchButton.UseVisualStyleBackColor = True
        '
        'InterfaceWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(904, 560)
        Me.Controls.Add(Me.SearchButton)
        Me.Controls.Add(Me.ResultsLabel)
        Me.Controls.Add(Me.ResultsBox)
        Me.Controls.Add(Me.GUIDsLabel)
        Me.Controls.Add(Me.GUIDListBox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "InterfaceWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GUIDListBox As TextBox
   Friend WithEvents GUIDsLabel As Label
    Friend WithEvents ResultsBox As TextBox
    Friend WithEvents ResultsLabel As Label
    Friend WithEvents SearchButton As Button
End Class
