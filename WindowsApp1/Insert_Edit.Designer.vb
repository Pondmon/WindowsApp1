<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Insert_Edit
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TEXT_INSERT_BARCODE = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(84, 83)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "ตกลง"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TEXT_INSERT_BARCODE
        '
        Me.TEXT_INSERT_BARCODE.Location = New System.Drawing.Point(12, 42)
        Me.TEXT_INSERT_BARCODE.Name = "TEXT_INSERT_BARCODE"
        Me.TEXT_INSERT_BARCODE.Size = New System.Drawing.Size(232, 20)
        Me.TEXT_INSERT_BARCODE.TabIndex = 1
        '
        'Insert_Edit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(256, 137)
        Me.Controls.Add(Me.TEXT_INSERT_BARCODE)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Insert_Edit"
        Me.Text = "Insert_Edit"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents TEXT_INSERT_BARCODE As TextBox
End Class
