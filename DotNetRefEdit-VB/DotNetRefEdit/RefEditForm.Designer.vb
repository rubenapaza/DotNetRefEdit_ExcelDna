Namespace DotNetRefEdit
	Partial Public Class RefEditForm
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
            Me.InputBox1 = New System.Windows.Forms.RichTextBox()
            Me.label1 = New System.Windows.Forms.Label()
            Me.InputBox2 = New System.Windows.Forms.RichTextBox()
            Me.label2 = New System.Windows.Forms.Label()
            Me.label3 = New System.Windows.Forms.Label()
            Me.DestinationBox = New System.Windows.Forms.RichTextBox()
            Me.label4 = New System.Windows.Forms.Label()
            Me.InsertButton = New System.Windows.Forms.Button()
            Me.label5 = New System.Windows.Forms.Label()
            Me.EvaluationBox = New System.Windows.Forms.RichTextBox()
            Me.btnAug = New System.Windows.Forms.Button()
            Me.btnAdd = New System.Windows.Forms.Button()
            Me.btnDest = New System.Windows.Forms.Button()
            Me.SuspendLayout()
            '
            'InputBox1
            '
            Me.InputBox1.Location = New System.Drawing.Point(62, 44)
            Me.InputBox1.Multiline = False
            Me.InputBox1.Name = "InputBox1"
            Me.InputBox1.Size = New System.Drawing.Size(351, 28)
            Me.InputBox1.TabIndex = 0
            Me.InputBox1.Text = ""
            '
            'label1
            '
            Me.label1.AutoSize = True
            Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.label1.Location = New System.Drawing.Point(12, 9)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(295, 20)
            Me.label1.TabIndex = 12
            Me.label1.Text = "Select input ranges and sum the values !"
            '
            'InputBox2
            '
            Me.InputBox2.Location = New System.Drawing.Point(62, 78)
            Me.InputBox2.Multiline = False
            Me.InputBox2.Name = "InputBox2"
            Me.InputBox2.Size = New System.Drawing.Size(351, 28)
            Me.InputBox2.TabIndex = 3
            Me.InputBox2.Text = ""
            '
            'label2
            '
            Me.label2.AutoSize = True
            Me.label2.Location = New System.Drawing.Point(12, 44)
            Me.label2.Name = "label2"
            Me.label2.Size = New System.Drawing.Size(44, 13)
            Me.label2.TabIndex = 2
            Me.label2.Text = "Augend"
            '
            'label3
            '
            Me.label3.AutoSize = True
            Me.label3.Location = New System.Drawing.Point(12, 78)
            Me.label3.Name = "label3"
            Me.label3.Size = New System.Drawing.Size(44, 13)
            Me.label3.TabIndex = 5
            Me.label3.Text = "Addend"
            '
            'DestinationBox
            '
            Me.DestinationBox.Location = New System.Drawing.Point(79, 148)
            Me.DestinationBox.Multiline = False
            Me.DestinationBox.Name = "DestinationBox"
            Me.DestinationBox.Size = New System.Drawing.Size(334, 28)
            Me.DestinationBox.TabIndex = 6
            Me.DestinationBox.Text = ""
            '
            'label4
            '
            Me.label4.AutoSize = True
            Me.label4.Location = New System.Drawing.Point(13, 151)
            Me.label4.Name = "label4"
            Me.label4.Size = New System.Drawing.Size(60, 13)
            Me.label4.TabIndex = 8
            Me.label4.Text = "Destination"
            '
            'InsertButton
            '
            Me.InsertButton.BackColor = System.Drawing.SystemColors.ScrollBar
            Me.InsertButton.Location = New System.Drawing.Point(79, 194)
            Me.InsertButton.Name = "InsertButton"
            Me.InsertButton.Size = New System.Drawing.Size(219, 30)
            Me.InsertButton.TabIndex = 9
            Me.InsertButton.Text = "Insert"
            Me.InsertButton.UseVisualStyleBackColor = False
            '
            'label5
            '
            Me.label5.AutoSize = True
            Me.label5.Location = New System.Drawing.Point(13, 240)
            Me.label5.Name = "label5"
            Me.label5.Size = New System.Drawing.Size(57, 13)
            Me.label5.TabIndex = 11
            Me.label5.Text = "Evaluation"
            '
            'EvaluationBox
            '
            Me.EvaluationBox.Location = New System.Drawing.Point(79, 237)
            Me.EvaluationBox.Name = "EvaluationBox"
            Me.EvaluationBox.ReadOnly = True
            Me.EvaluationBox.Size = New System.Drawing.Size(334, 28)
            Me.EvaluationBox.TabIndex = 10
            Me.EvaluationBox.Text = ""
            '
            'btnAug
            '
            Me.btnAug.Location = New System.Drawing.Point(414, 43)
            Me.btnAug.Name = "btnAug"
            Me.btnAug.Size = New System.Drawing.Size(29, 29)
            Me.btnAug.TabIndex = 1
            Me.btnAug.Text = "_"
            Me.btnAug.UseVisualStyleBackColor = True
            '
            'btnAdd
            '
            Me.btnAdd.Location = New System.Drawing.Point(414, 77)
            Me.btnAdd.Name = "btnAdd"
            Me.btnAdd.Size = New System.Drawing.Size(29, 29)
            Me.btnAdd.TabIndex = 4
            Me.btnAdd.Text = "_"
            Me.btnAdd.UseVisualStyleBackColor = True
            '
            'btnDest
            '
            Me.btnDest.Location = New System.Drawing.Point(414, 147)
            Me.btnDest.Name = "btnDest"
            Me.btnDest.Size = New System.Drawing.Size(29, 29)
            Me.btnDest.TabIndex = 7
            Me.btnDest.Text = "_"
            Me.btnDest.UseVisualStyleBackColor = True
            '
            'RefEditForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(468, 273)
            Me.Controls.Add(Me.btnDest)
            Me.Controls.Add(Me.btnAdd)
            Me.Controls.Add(Me.btnAug)
            Me.Controls.Add(Me.EvaluationBox)
            Me.Controls.Add(Me.label5)
            Me.Controls.Add(Me.InsertButton)
            Me.Controls.Add(Me.label4)
            Me.Controls.Add(Me.DestinationBox)
            Me.Controls.Add(Me.label3)
            Me.Controls.Add(Me.label2)
            Me.Controls.Add(Me.InputBox2)
            Me.Controls.Add(Me.label1)
            Me.Controls.Add(Me.InputBox1)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
            Me.MaximizeBox = False
            Me.MaximumSize = New System.Drawing.Size(484, 312)
            Me.MinimizeBox = False
            Me.MinimumSize = New System.Drawing.Size(484, 312)
            Me.Name = "RefEditForm"
            Me.Text = "RefEdit"
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

#End Region

        Private InputBox1 As System.Windows.Forms.RichTextBox
		Private label1 As System.Windows.Forms.Label
		Private InputBox2 As System.Windows.Forms.RichTextBox
		Private label2 As System.Windows.Forms.Label
		Private label3 As System.Windows.Forms.Label
		Private DestinationBox As System.Windows.Forms.RichTextBox
		Private label4 As System.Windows.Forms.Label
		Private WithEvents InsertButton As System.Windows.Forms.Button
		Private label5 As System.Windows.Forms.Label
		Private EvaluationBox As System.Windows.Forms.RichTextBox
        Friend WithEvents btnAug As Forms.Button
        Friend WithEvents btnAdd As Forms.Button
        Friend WithEvents btnDest As Forms.Button
    End Class
End Namespace
