<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Dim Button1 As System.Windows.Forms.Button
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Button1 = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Button1.AutoEllipsis = True
        Button1.BackColor = System.Drawing.Color.PowderBlue
        Button1.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Button1.ForeColor = System.Drawing.Color.Black
        Button1.Location = New System.Drawing.Point(107, 243)
        Button1.Name = "Button1"
        Button1.Size = New System.Drawing.Size(549, 55)
        Button1.TabIndex = 0
        Button1.TabStop = False
        Button1.Text = "" & Global.Microsoft.VisualBasic.ChrW(9) & "คำร้องขอลงทะเบียนเพิ่ม/เปลี่ยนกลุ่ม กรณีกลุ่มเต็ม/ ถอนรายวิชา (ท.1)"
        Button1.UseVisualStyleBackColor = False
        AddHandler Button1.Click, AddressOf Me.Button1_Click
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(143, 86)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1100, 117)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'BackgroundWorker1
        '
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.PowderBlue
        Me.Button2.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(107, 325)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(549, 55)
        Me.Button2.TabIndex = 1
        Me.Button2.TabStop = False
        Me.Button2.Text = "" & Global.Microsoft.VisualBasic.ChrW(9) & "คำร้องขอถอนรายวิชา (ท.8)"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.PowderBlue
        Me.Button3.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(107, 417)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(549, 55)
        Me.Button3.TabIndex = 2
        Me.Button3.TabStop = False
        Me.Button3.Text = "" & Global.Microsoft.VisualBasic.ChrW(9) & "คำร้องขอลงทะเบียนเรียนเกิน/ต่ำกว่าหน่วยกิตที่กำหนด (ท.16)"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.PowderBlue
        Me.Button4.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(107, 509)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(549, 52)
        Me.Button4.TabIndex = 3
        Me.Button4.TabStop = False
        Me.Button4.Text = "" & Global.Microsoft.VisualBasic.ChrW(9) & "คำร้องขอลงทะเบียนเรียนโดยมีเวลาสอบซ้ำซ้อน (ท.17)"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.PowderBlue
        Me.Button5.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.Black
        Me.Button5.Location = New System.Drawing.Point(722, 327)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(549, 53)
        Me.Button5.TabIndex = 4
        Me.Button5.TabStop = False
        Me.Button5.Text = "คำร้องขอลาระหว่างเรียน (ท.95)"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.PowderBlue
        Me.Button6.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Black
        Me.Button6.Location = New System.Drawing.Point(722, 243)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(549, 53)
        Me.Button6.TabIndex = 5
        Me.Button6.TabStop = False
        Me.Button6.Text = "คำร้องขอแบ่งชาระค่าธรรมเนียม(ท.20)"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.PowderBlue
        Me.Button7.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button7.ForeColor = System.Drawing.Color.Black
        Me.Button7.Location = New System.Drawing.Point(107, 596)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(549, 53)
        Me.Button7.TabIndex = 6
        Me.Button7.TabStop = False
        Me.Button7.Text = "คำร้องขอลงทะเบียนเรียนรายวิชานอกเหนือหลักสูตร หรือมีเงื่อนไขเฉพาะ (ท.18)"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.Color.PowderBlue
        Me.Button8.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button8.ForeColor = System.Drawing.Color.Black
        Me.Button8.Location = New System.Drawing.Point(722, 509)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(549, 53)
        Me.Button8.TabIndex = 7
        Me.Button8.TabStop = False
        Me.Button8.Text = "คำร้องขอสอบวิชาโครงงาน ระดับปริญญาตรี"
        Me.Button8.UseVisualStyleBackColor = False
        '
        'Button9
        '
        Me.Button9.BackColor = System.Drawing.Color.PowderBlue
        Me.Button9.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button9.ForeColor = System.Drawing.Color.Black
        Me.Button9.Location = New System.Drawing.Point(722, 417)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(549, 53)
        Me.Button9.TabIndex = 8
        Me.Button9.TabStop = False
        Me.Button9.Text = "คำร้องขอใช้คะแนนสอบ TOEFL/IELTS/TOEIC "
        Me.Button9.UseVisualStyleBackColor = False
        '
        'Button10
        '
        Me.Button10.BackColor = System.Drawing.Color.PowderBlue
        Me.Button10.Font = New System.Drawing.Font("TH SarabunPSK", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button10.ForeColor = System.Drawing.Color.Black
        Me.Button10.Location = New System.Drawing.Point(722, 594)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(549, 53)
        Me.Button10.TabIndex = 9
        Me.Button10.TabStop = False
        Me.Button10.Text = "คำร้องยืนยันกำรจัดทำโครงงานระดับปริญญาตรี สาขาวิชาวิศวกรรมโทรคมนาคม"
        Me.Button10.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.RoyalBlue
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(1354, 685)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Button1)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button5)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "โปรแกรมคำร้อง"
        Me.TransparencyKey = System.Drawing.Color.LightSkyBlue
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents Button10 As Button
End Class
