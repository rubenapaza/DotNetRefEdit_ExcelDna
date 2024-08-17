Imports System
Imports System.Windows.Forms
Imports ExcelDna.Integration

Namespace DotNetRefEdit
	Partial Public Class RefEditForm
		Inherits System.Windows.Forms.Form

        Private Shared callfrmRefEdit As frmRefEdit_Plus
        Private Shared myFormError As frmRefEdit_Plus
        Private isInputOn As Boolean

        Private ReadOnly _selectionTracker As ExcelSelectionTracker
        Private ReadOnly _application As Microsoft.Office.Interop.Excel.Application
        Private _focusedBox As RichTextBox

        Public Sub New(ByVal excelThreadId As Integer)
            InitializeComponent()
            _selectionTracker = New ExcelSelectionTracker(excelThreadId)
            _application = DirectCast(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)

            AddHandler Me.Closed, Sub()
                                      RemoveHandler _selectionTracker.NewSelection, AddressOf ChangeText
                                      _selectionTracker.Stop()
                                  End Sub

            AddHandler _selectionTracker.NewSelection, AddressOf ChangeText

            AddHandler Me.Deactivate, AddressOf CheckFocus

            AddHandler InputBox1.TextChanged, AddressOf OnNewInput
            AddHandler InputBox2.TextChanged, AddressOf OnNewInput

            AddHandler InputBox1.KeyDown, AddressOf CheckF4
            AddHandler InputBox2.KeyDown, AddressOf CheckF4
            AddHandler DestinationBox.KeyDown, AddressOf CheckF4
        End Sub

        Private Sub CheckFocus(ByVal sender As Object, ByVal eventArgs As EventArgs)
            If InputBox1.Focused Then
                _focusedBox = InputBox1
            ElseIf InputBox2.Focused Then
                _focusedBox = InputBox2
            ElseIf DestinationBox.Focused Then
                _focusedBox = DestinationBox
            Else
                _focusedBox = Nothing
            End If
        End Sub

        ''' <summary>
        ''' Build final formula: to be run in UI thread
        ''' </summary>
        ''' <returns></returns>
        Private Function BuildFormula() As String
            Return String.Format("=sum({0},{1})", InputBox1.Text, InputBox2.Text)
        End Function

        ''' <summary>
        ''' Evaluate the formula: to be run in Excel thread
        ''' </summary>
        Private Sub UpdateEvaluation(ByVal formula As String)
            Dim formulaResult As Object = ExcelHelper.EvaluateFormula(formula, _application)
            Invoke(New Action(Sub() EvaluationBox.Text = (If(formulaResult, "")).ToString()))
        End Sub

        Private Sub OnNewInput(ByVal sender As Object, ByVal e As EventArgs)
            Dim formula As String = BuildFormula()
            ExcelAsyncUtil.QueueAsMacro(Sub() UpdateEvaluation(formula))
        End Sub

        Private Sub ChangeText(ByVal sender As Object, ByVal args As RangeAddressEventArgs)
            Invoke(New Action(Sub()
                                  If _focusedBox IsNot Nothing Then
                                      _focusedBox.Text = args.Address
                                      _focusedBox.Select(_focusedBox.Text.Length, 0)
                                  End If
                              End Sub))
        End Sub

        Private Sub CheckF4(ByVal sender As Object, ByVal e As KeyEventArgs)
            Dim textBox As RichTextBox = TryCast(sender, RichTextBox)
            If e.KeyCode = Keys.F4 AndAlso textBox IsNot Nothing Then
                Dim text As String = textBox.Text

                ExcelAsyncUtil.QueueAsMacro(Sub()
                                                Dim newAddress As String = Nothing
                                                If ExcelHelper.TryF4(text, _application, newAddress) Then
                                                    Invoke(New Action(Sub()
                                                                          textBox.Text = newAddress
                                                                          textBox.Select(textBox.Text.Length, 0)
                                                                      End Sub))
                                                End If
                                            End Sub)
            End If
        End Sub

        Private Sub InsertButton_Click(ByVal sender As Object, ByVal e As EventArgs) Handles InsertButton.Click
            If Not Me.InsertButton.IsHandleCreated Then Return

            Dim formula As String = BuildFormula()
            Dim destination As String = DestinationBox.Text

            If formula IsNot Nothing AndAlso Not String.IsNullOrEmpty(destination) Then
                ExcelAsyncUtil.QueueAsMacro(Sub() ExcelHelper.InsertFormula(formula, _application, destination))
            End If
            Me.isOKForm = True
            Me.exit = True
            Me.Close()
        End Sub

        ' ------------
        Public Property [exit] As Boolean

        Public Property isOKForm As Boolean

        Public Property strTextBox As String
            Get
                Return ""
            End Get
            Set(ByVal value As String)
                If Not Me.isInputOn Then
                    Return
                End If
                RefEditForm.callfrmRefEdit.strTextBox = value
            End Set
        End Property

        Private Sub BtnAug_Click(sender As Object, e As EventArgs) Handles btnAug.Click
            If Not Me.btnAug.IsHandleCreated Then Return

            Me.Visible = False
            RefEditForm.callfrmRefEdit = New frmRefEdit_Plus("Select a range")
            RefEditForm.callfrmRefEdit.ShowInTaskbar = False
            RefEditForm.callfrmRefEdit.Show(DirectCast(form_Interface_Plus.NativeWindowWrapper.ExcelWindow, IWin32Window))
            Me.isInputOn = True
            Do While Not RefEditForm.callfrmRefEdit.exit
                System.Windows.Forms.Application.DoEvents()
            Loop
            Me.isInputOn = False
            Me.Visible = True
            If Not RefEditForm.callfrmRefEdit.isOkRefEd OrElse Not (RefEditForm.callfrmRefEdit.strReturnRefEdit <> "") Then
                Return
            End If
            Me.InputBox1.Text = RefEditForm.callfrmRefEdit.strReturnRefEdit
        End Sub

        Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
            If Not Me.btnAdd.IsHandleCreated Then Return

            Me.Visible = False
            RefEditForm.callfrmRefEdit = New frmRefEdit_Plus("Select a range")
            RefEditForm.callfrmRefEdit.ShowInTaskbar = False '0 
            RefEditForm.callfrmRefEdit.Show(DirectCast(form_Interface_Plus.NativeWindowWrapper.ExcelWindow, IWin32Window)) '0 
            Me.isInputOn = True
            Do While Not RefEditForm.callfrmRefEdit.exit
                System.Windows.Forms.Application.DoEvents()
            Loop
            Me.isInputOn = False
            Me.Visible = True
            If Not RefEditForm.callfrmRefEdit.isOkRefEd OrElse Not (RefEditForm.callfrmRefEdit.strReturnRefEdit <> "") Then
                Return
            End If
            Me.InputBox2.Text = RefEditForm.callfrmRefEdit.strReturnRefEdit
        End Sub

        Private Sub BtnDest_Click(sender As Object, e As EventArgs) Handles btnDest.Click
            If Not Me.btnDest.IsHandleCreated Then Return

            Me.Visible = False
            RefEditForm.callfrmRefEdit = New frmRefEdit_Plus("Select a Cell")
            RefEditForm.callfrmRefEdit.ShowInTaskbar = False '0 
            RefEditForm.callfrmRefEdit.Show(DirectCast(form_Interface_Plus.NativeWindowWrapper.ExcelWindow, IWin32Window)) '0 
            Me.isInputOn = True
            Do While Not RefEditForm.callfrmRefEdit.exit
                System.Windows.Forms.Application.DoEvents()
            Loop
            Me.isInputOn = False
            Me.Visible = True
            If Not RefEditForm.callfrmRefEdit.isOkRefEd OrElse Not (RefEditForm.callfrmRefEdit.strReturnRefEdit <> "") Then
                Return
            End If
            Me.DestinationBox.Text = RefEditForm.callfrmRefEdit.strReturnRefEdit
        End Sub
    End Class
End Namespace
