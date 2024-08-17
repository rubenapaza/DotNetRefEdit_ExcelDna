Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Namespace DotNetRefEdit
	''' <summary>
	''' Common functions used by the form and the WPF window
	''' </summary>
	Friend Module ExcelHelper
		''' <summary>
		''' Make Excel evaluate the formula.
		''' To be run in Excel thread.
		''' </summary>
		''' <param name="formula"></param>
		''' <param name="application"></param>
		''' <returns></returns>
		Public Function EvaluateFormula(ByVal formula As String, ByVal application As Application) As Object
			Try
				Dim formulaResult As Object = application.Evaluate(formula)

				' Check the Excel error codes
				If TypeOf formulaResult Is Integer Then
					Select Case DirectCast(formulaResult, Integer)
						Case -2146826288, -2146826281, -2146826265, -2146826259, -2146826252, -2146826246, -2146826273
							Return "Could not evaluate function"
					End Select
				End If

				Return formulaResult
			Catch
				Return "Could not evaluate function"
			End Try
		End Function

		''' <summary>
		''' Insert formula into Excel range.
		''' To be run in Excel thread.
		''' </summary>
		''' <param name="formula"></param>
		''' <param name="application"></param>
		''' <param name="destination"></param>
		Public Sub InsertFormula(ByVal formula As String, ByVal application As Application, ByVal destination As String)
			Dim rg As Range = Nothing

			Try
				rg = application.Range(destination)
                ''rg.Formula = formula
                rg.FormulaArray = formula
            Finally
				If rg IsNot Nothing Then
					Marshal.ReleaseComObject(rg)
				End If
			End Try
		End Sub

		''' <summary>
		''' Try to switch the address format, following this sequence:
		''' 1. RowAbsolute=False, ColumnAbsolute=False
		''' 2. RowAbsolute=True, ColumnAbsolute=True
		''' 3. RowAbsolute=True, ColumnAbsolute=False
		''' 4. RowAbsolute=False, ColumnAbsolute=True
		''' This shall reproduce the behaviour of the Excel "Function Arguments" form when the user hits F4.
		''' To be run in Excel thread.
		''' </summary>
		''' <param name="text"></param>
		''' <param name="application"></param>
		''' <param name="newAddress"></param>
		''' <returns></returns>
		Public Function TryF4(ByVal text As String, ByVal application As Application, <System.Runtime.InteropServices.Out()> ByRef newAddress As String) As Boolean
			Try
				Dim formulaResult As Object = application.Evaluate(text)

				If TypeOf formulaResult Is Range Then
					Dim relativePart As String = text

					If text.Contains("!") Then
						relativePart = text.Substring(text.IndexOf("!") + 1, text.Length - text.IndexOf("!") - 1)
					End If

					Dim range As Range = DirectCast(formulaResult, Range)

					Dim addresses As New List(Of String) From {range.Address(False, False, XlReferenceStyle.xlA1, False), range.Address(True, True, XlReferenceStyle.xlA1, False), range.Address(True, False, XlReferenceStyle.xlA1, False), range.Address(False, True, XlReferenceStyle.xlA1, False)}

					Dim found As Boolean = False
					For i As Integer = 0 To addresses.Count - 1
						If addresses(i) = relativePart Then
							relativePart = addresses(If(i + 1 = addresses.Count, 0, i + 1))
							found = True
							Exit For
						End If
					Next i

					If Not found Then
						newAddress = range.Address(False, False, XlReferenceStyle.xlA1, True)
						Return True
					End If

					newAddress = If(text.Contains("!"), String.Concat(text.Substring(0, text.IndexOf("!") + 1), relativePart), relativePart)

					Return True
				End If

				newAddress = Nothing
				Return False
			Catch
				newAddress = Nothing
				Return False
			End Try
		End Function
	End Module
End Namespace
