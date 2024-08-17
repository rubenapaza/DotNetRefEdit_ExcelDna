Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

Namespace DotNetRefEdit
	Public Class RangeAddressEventArgs
		Inherits EventArgs

		Public Property Address As String
	End Class

	Public Class ExcelSelectionTracker
		Private ReadOnly _application As Microsoft.Office.Interop.Excel.Application

		Public Event NewSelection As EventHandler(Of RangeAddressEventArgs)

		Private ReadOnly _hHookCwp As Integer
        Private ReadOnly _procCwp As WindowsInterop.HookProc '  Note: do not make this delegate a local variable within the ExcelSelectionTracker constructor because it must not be collected by the GC before the unhook

        Public Sub New(ByVal excelThreadId As Integer)
            _application = DirectCast(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            AddHandler _application.SheetSelectionChange, AddressOf OnNewSelection

            _procCwp = AddressOf CwpProc

            _hHookCwp = WindowsInterop.SetWindowsHookEx(HookType.WH_CALLWNDPROC, _procCwp, New IntPtr(0), excelThreadId)
            If _hHookCwp = 0 Then
                Throw New Exception("Failed to hook WH_CALLWNDPROC")
            End If
        End Sub

        Public Sub [Stop]()
			RemoveHandler _application.SheetSelectionChange, AddressOf OnNewSelection

			If Not WindowsInterop.UnhookWindowsHookEx(_hHookCwp) Then
				Debug.Print("Error: Failed to unhook WH_CALLWNDPROC")
			End If
		End Sub

		Private Function CwpProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
			Dim cwpStruct As CwpStruct = DirectCast(Marshal.PtrToStructure(lParam, GetType(CwpStruct)), CwpStruct)

			If nCode < 0 Then
				Return WindowsInterop.CallNextHookEx(_hHookCwp, nCode, wParam, lParam)
			End If

			If cwpStruct.message = WindowsInterop.WM_MOUSEACTIVATE Then
				' We got a WM_MOUSEACTIVATE message. Now we will check that the target handle is a workbook window.
				' Workbook windows have the name "EXCEL7".
				Dim isWorkbookWindow As Boolean = False

				Try
					Dim cname As New StringBuilder(256)
					WindowsInterop.GetClassNameW(cwpStruct.hwnd, cname, cname.Capacity)
					If cname.ToString() = "EXCEL7" Then
						isWorkbookWindow = True
					End If
				Catch e As Exception
					Debug.Print("Could not get the window name: {0}", e)
				End Try

				If isWorkbookWindow Then
					' If the window is not activated, then Excel will activate it and then discard the message. That's why the user cannot select a range at the same time.
					' The following statement will activate the window before Excel treats the message, thus it will not activate the window and it will keep proceeding the message. 
					' In that way, it is possible to select the range.
					Try
						WindowsInterop.SetFocus(cwpStruct.hwnd)
					Catch e As Exception
						Debug.Print("Failed to set the focus: {0}", e)
					End Try

					' If the user chooses a cell which was already selected, then the event SheetSelectionChange will not be raised.
					' A workaround is to send the current selection when the Excel window gets the focus. 
					' Note that if the user selects a different range, then 2 events will be raised: a first one with the current selection, 
					' and a second one with the new selection.
					Try
						OnNewSelection(Nothing, DirectCast(_application.Selection, Range))
					Catch
					End Try
				End If
			End If

			Return WindowsInterop.CallNextHookEx(_hHookCwp, nCode, wParam, lParam)
		End Function

		Private Sub OnNewSelection(ByVal sh As Object, ByVal target As Range)
			Try
				Dim newSelection = NewSelectionEvent
				If newSelection IsNot Nothing Then
                    ''newSelection(Me, New RangeAddressEventArgs With {.Address = target.Address(False, False, XlReferenceStyle.xlA1, True)})
                    newSelection(Me, New RangeAddressEventArgs With {.Address = target.Address(True, True, XlReferenceStyle.xlA1, True)})
                End If
			Catch
			End Try
		End Sub
	End Class
End Namespace
