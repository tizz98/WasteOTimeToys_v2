<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmToys
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
        Me.btnMagic = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnMagic
        '
        Me.btnMagic.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMagic.Location = New System.Drawing.Point(12, 12)
        Me.btnMagic.Name = "btnMagic"
        Me.btnMagic.Size = New System.Drawing.Size(260, 238)
        Me.btnMagic.TabIndex = 0
        Me.btnMagic.Text = "Magic"
        Me.btnMagic.UseVisualStyleBackColor = True
        '
        'frmToys
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.btnMagic)
        Me.Name = "frmToys"
        Me.Text = "Waste O Time Toys v2"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnMagic As Button
End Class
