Imports System.Windows.Forms

Namespace DotNetRefEdit
    Public Class frmRefEdit_Plus
        Inherits Form

        Private WithEvents btnOkInputRange As Button
        Private _focusedBox As System.Windows.Forms.RichTextBox
        Private WithEvents rtxtbxInputRange As RichTextBox
        Private _selectionTracker As ExcelSelectionTracker

        Public Sub New(ByVal textRefEfit As String)
            Me.exit = False
            Me.isOkRefEd = False
            Me.InitializeComponent()
            Me.Text = textRefEfit
            Me.rtxtbxInputRange.Text = ""

            _selectionTracker = New ExcelSelectionTracker(WindowsInterop.GetCurrentThreadId())
            AddHandler Me.Closed, Sub() '0
                                      RemoveHandler _selectionTracker.NewSelection, AddressOf ChangeText
                                      _selectionTracker.Stop()
                                  End Sub
            AddHandler _selectionTracker.NewSelection, AddressOf ChangeText

            AddHandler Me.Deactivate, AddressOf CheckFocus

        End Sub

        Private Sub ChangeText(ByVal sender As Object, ByVal args As RangeAddressEventArgs)

            Invoke(New System.Action(Sub()

                                         If _focusedBox IsNot Nothing Then
                                             _focusedBox.Text = args.Address
                                             _focusedBox.Select(_focusedBox.Text.Length, 0)
                                         End If
                                     End Sub))
        End Sub

        Private Sub CheckFocus(ByVal sender As Object, ByVal eventArgs As EventArgs)

            If rtxtbxInputRange.Focused Then
                _focusedBox = rtxtbxInputRange
            Else
                _focusedBox = Nothing
            End If
        End Sub

        Private Sub frmRefEdit_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        End Sub

        Private Sub frmRefEdit_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
            If Not Me.isOkRefEd Then
                Me.strReturnRefEdit = ""
            End If
            Me.exit = True
        End Sub

        Public Property strTextBox As String
            Get
                Return Me.rtxtbxInputRange.Text
            End Get
            Set(ByVal value As String)
                Me.rtxtbxInputRange.Text = value
            End Set
        End Property

        Public Property strReturnRefEdit As String

        Public Property [exit] As Boolean

        Public Property isOkRefEd As Boolean

        Private Sub btnAug_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOkInputRange.Click
            Me.isOkRefEd = True
            Me.strReturnRefEdit = Me.rtxtbxInputRange.Text
            Me.exit = True
            Me.Close()
        End Sub

        Private Sub InitializeComponent()
            Me.btnOkInputRange = New System.Windows.Forms.Button()
            Me.rtxtbxInputRange = New System.Windows.Forms.RichTextBox()
            Me.SuspendLayout()
            '
            'btnOkInputRange
            '
            Me.btnOkInputRange.Location = New System.Drawing.Point(427, 11)
            Me.btnOkInputRange.Name = "btnOkInputRange"
            Me.btnOkInputRange.Size = New System.Drawing.Size(29, 29)
            Me.btnOkInputRange.TabIndex = 3
            Me.btnOkInputRange.Text = "_"
            Me.btnOkInputRange.UseVisualStyleBackColor = True
            '
            'rtxtbxInputRange
            '
            Me.rtxtbxInputRange.Location = New System.Drawing.Point(12, 12)
            Me.rtxtbxInputRange.Multiline = False
            Me.rtxtbxInputRange.Name = "rtxtbxInputRange"
            Me.rtxtbxInputRange.Size = New System.Drawing.Size(409, 28)
            Me.rtxtbxInputRange.TabIndex = 2
            Me.rtxtbxInputRange.Text = ""
            '
            'frmRefEdit_Plus
            '
            Me.ClientSize = New System.Drawing.Size(468, 55)
            Me.Controls.Add(Me.btnOkInputRange)
            Me.Controls.Add(Me.rtxtbxInputRange)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "frmRefEdit_Plus"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.ResumeLayout(False)

        End Sub

    End Class

End Namespace
