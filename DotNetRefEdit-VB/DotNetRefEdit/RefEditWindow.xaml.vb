Imports System
Imports System.Windows
Imports System.Windows.Input
Imports ExcelDna.Integration

Imports TextBox = System.Windows.Controls.TextBox

Namespace DotNetRefEdit
	Partial Public Class RefEditWindow
		Private ReadOnly _selectionTracker As ExcelSelectionTracker
		Private ReadOnly _application As Microsoft.Office.Interop.Excel.Application
		Private _focusedBox As TextBox

		Public Sub New(ByVal excelThreadId As Integer)
			InitializeComponent()

			_selectionTracker = New ExcelSelectionTracker(excelThreadId)
			_application = DirectCast(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)

			AddHandler Me.Closed, Sub()
				RemoveHandler _selectionTracker.NewSelection, AddressOf ChangeText
				_selectionTracker.Stop()
			End Sub

			AddHandler _selectionTracker.NewSelection, AddressOf ChangeText

			AddHandler Me.Deactivated, AddressOf CheckFocus

			AddHandler InputBox1.TextChanged, AddressOf OnNewInput
			AddHandler InputBox2.TextChanged, AddressOf OnNewInput

			AddHandler InputBox1.KeyDown, AddressOf CheckF4
			AddHandler InputBox2.KeyDown, AddressOf CheckF4
			AddHandler DestinationBox.KeyDown, AddressOf CheckF4
		End Sub

		Private Sub CheckFocus(ByVal sender As Object, ByVal eventArgs As EventArgs)
			If InputBox1.IsFocused Then
				_focusedBox = InputBox1
			ElseIf InputBox2.IsFocused Then
				_focusedBox = InputBox2
			ElseIf DestinationBox.IsFocused Then
				_focusedBox = DestinationBox
			Else
				_focusedBox = Nothing
			End If
		End Sub

		Private Sub ChangeText(ByVal sender As Object, ByVal args As RangeAddressEventArgs)
			Dispatcher.Invoke(New Action(Sub()
				If _focusedBox IsNot Nothing Then
					_focusedBox.Text = args.Address
					_focusedBox.CaretIndex = _focusedBox.Text.Length
				End If
			End Sub))
		End Sub

		Private Sub CheckF4(ByVal sender As Object, ByVal e As KeyEventArgs)
			Dim textBox As TextBox = TryCast(sender, TextBox)
			If e.Key = Key.F4 AndAlso textBox IsNot Nothing Then
				Dim text As String = textBox.Text

				ExcelAsyncUtil.QueueAsMacro(Sub()
					Dim newAddress As String = Nothing
					If ExcelHelper.TryF4(text, _application, newAddress) Then
						Dispatcher.Invoke(New Action(Sub()
							textBox.Text = newAddress
							textBox.CaretIndex = textBox.Text.Length
						End Sub))
					End If
				End Sub)
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
			Dispatcher.Invoke(New Action(Sub() EvaluationBox.Text = (If(formulaResult, "")).ToString()))
		End Sub

		Private Sub OnNewInput(ByVal sender As Object, ByVal e As EventArgs)
			Dim formula As String = BuildFormula()
			ExcelAsyncUtil.QueueAsMacro(Sub() UpdateEvaluation(formula))
		End Sub

		Private Sub InsertFormula(ByVal sender As Object, ByVal e As RoutedEventArgs)
			Dim formula As String = BuildFormula()
			Dim destination As String = DestinationBox.Text

			If formula IsNot Nothing AndAlso Not String.IsNullOrEmpty(destination) Then
				ExcelAsyncUtil.QueueAsMacro(Sub() ExcelHelper.InsertFormula(formula, _application, destination))
			End If
		End Sub
	End Class
End Namespace
