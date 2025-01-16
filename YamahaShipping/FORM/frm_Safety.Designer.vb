<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Safety
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
        Me.btnCasemark = New System.Windows.Forms.Button()
        Me.btnTag = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnCasemark
        '
        Me.btnCasemark.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCasemark.BackColor = System.Drawing.Color.SteelBlue
        Me.btnCasemark.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCasemark.ForeColor = System.Drawing.Color.White
        Me.btnCasemark.Location = New System.Drawing.Point(12, 12)
        Me.btnCasemark.Name = "btnCasemark"
        Me.btnCasemark.Size = New System.Drawing.Size(206, 80)
        Me.btnCasemark.TabIndex = 186
        Me.btnCasemark.Text = "CASEMARK"
        Me.btnCasemark.UseVisualStyleBackColor = False
        '
        'btnTag
        '
        Me.btnTag.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTag.BackColor = System.Drawing.Color.SteelBlue
        Me.btnTag.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTag.ForeColor = System.Drawing.Color.White
        Me.btnTag.Location = New System.Drawing.Point(335, 12)
        Me.btnTag.Name = "btnTag"
        Me.btnTag.Size = New System.Drawing.Size(206, 80)
        Me.btnTag.TabIndex = 187
        Me.btnTag.Text = "TAG"
        Me.btnTag.UseVisualStyleBackColor = False
        '
        'frm_Safety
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(553, 113)
        Me.Controls.Add(Me.btnTag)
        Me.Controls.Add(Me.btnCasemark)
        Me.Name = "frm_Safety"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnCasemark As Button
    Friend WithEvents btnTag As Button
End Class
