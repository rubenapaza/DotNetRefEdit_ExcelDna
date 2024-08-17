<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRefEdit_Plus
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
        Me.btnOk = New System.Windows.Forms.Button()
        Me.rtxtbxInputRange = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(448, 11)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(29, 29)
        Me.btnOk.TabIndex = 3
        Me.btnOk.Text = "_"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'rtxtbxInputRange
        '
        Me.rtxtbxInputRange.Location = New System.Drawing.Point(12, 12)
        Me.rtxtbxInputRange.Multiline = False
        Me.rtxtbxInputRange.Name = "rtxtbxInputRange"
        Me.rtxtbxInputRange.Size = New System.Drawing.Size(435, 28)
        Me.rtxtbxInputRange.TabIndex = 2
        Me.rtxtbxInputRange.Text = ""
        '
        'frmRefEdit_Plus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(488, 53)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.rtxtbxInputRange)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRefEdit_Plus"
        Me.Text = "RefEdit Plus"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnOk As Forms.Button
    Private WithEvents rtxtbxInputRange As Forms.RichTextBox
End Class
