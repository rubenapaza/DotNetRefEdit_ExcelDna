Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI


Namespace DotNetRefEdit
	<ComVisible(True)>
	Public Class MyRibbon
		Inherits ExcelRibbon

        Private ReadOnly _excelThreadId As Integer

        Private _refEditForm1 As RefEditForm
		Private _refEditWindow1 As RefEditWindow
		Private _refEditForm2 As RefEditForm
        Private _refEditWindow2 As RefEditWindow

        Private _refEditForm_Plus As frmRefEdit_Plus

        Public Sub New()
            _excelThreadId = WindowsInterop.GetCurrentThreadId()
        End Sub

		Private Function CheckWorkbook() As Boolean
			Try
				Dim app As Microsoft.Office.Interop.Excel.Application = DirectCast(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
				If app.Workbooks.Count = 0 Then
					MessageBox.Show("Please open a workbook before starting UI.", "Error")
					Return False
				End If

				Return True
			Catch e As Exception
				Debug.Print("Couldn't check workbook: {0}", e)
				Return False
			End Try
		End Function

		Public Sub OpenWinFormInExcelThread(ByVal control As IRibbonControl)
			If Not CheckWorkbook() Then
				Return
			End If

			If _refEditForm1 Is Nothing Then
				Try
                    _refEditForm1 = New RefEditForm(_excelThreadId)
                    _refEditForm1.ShowInTaskbar = False '
                    AddHandler _refEditForm1.Closed, Sub()

                                                         _refEditForm1 = Nothing

                                                     End Sub

                    ''_refEditForm1.Show()
                    _refEditForm1.Show(DirectCast(form_Interface_Plus.NativeWindowWrapper.ExcelWindow, IWin32Window)) '
                Catch e As Exception
					Debug.Print("Error: {0}", e)
				End Try
			Else
				_refEditForm1.Activate()
			End If
		End Sub

		Public Sub OpenWinFormInSeparateThread(ByVal control As IRibbonControl)
			If Not CheckWorkbook() Then
				Return
			End If

			If _refEditForm2 Is Nothing Then
				Dim thread As New Thread(Sub()
                                             Try
                                                 _refEditForm2 = New RefEditForm(_excelThreadId)
                                                 ''_refEditForm2.ShowInTaskbar = False '
                                                 AddHandler _refEditForm2.Closed, Sub()
                                                                                      _refEditForm2 = Nothing

                                                                                  End Sub
                                                 _refEditForm2.ShowDialog()
                                                 ''_refEditForm2.Show(DirectCast(form_Interface_Plus.NativeWindowWrapper.ExcelWindow, IWin32Window))
                                             Catch e As Exception

                                                 Debug.Print("Error: {0}", e)
                                             End Try

                                         End Sub)

				thread.SetApartmentState(ApartmentState.STA)
				thread.Start()
			Else
				_refEditForm2.Invoke(New Action(Sub() _refEditForm2.Activate()))
			End If
		End Sub

		Public Sub OpenWPFInExcelThread(ByVal control As IRibbonControl)
			If Not CheckWorkbook() Then
				Return
			End If

			If _refEditWindow1 Is Nothing Then
				Try
					_refEditWindow1 = New RefEditWindow(_excelThreadId)
					AddHandler _refEditWindow1.Closed, Sub()
						_refEditWindow1 = Nothing
					End Sub
					_refEditWindow1.Show()
				Catch e As Exception
					Debug.Print("Error: {0}", e)
				End Try
			Else
				_refEditWindow1.Activate()
			End If
		End Sub

        Public Sub OpenWPFInSeparateThread(ByVal control As IRibbonControl)
            If Not CheckWorkbook() Then
                Return
            End If

            If _refEditWindow2 Is Nothing Then
                Dim thread As New Thread(Sub()
                                             Try
                                                 _refEditWindow2 = New RefEditWindow(_excelThreadId)
                                                 AddHandler _refEditWindow2.Closed, Sub()
                                                                                        _refEditWindow2 = Nothing
                                                                                    End Sub
                                                 _refEditWindow2.ShowDialog()
                                             Catch e As Exception
                                                 Debug.Print("Error: {0}", e)
                                             End Try
                                         End Sub)

                thread.SetApartmentState(ApartmentState.STA)
                thread.Start()
            Else
                _refEditWindow2.Dispatcher.Invoke(New Action(Sub() _refEditWindow2.Activate()))
            End If
        End Sub

        Public Sub OpenWinFormRefEdit(ByVal control As IRibbonControl)
            If Not CheckWorkbook() Then
                Return
            End If

            If _refEditForm_Plus Is Nothing Then
                Try
                    _refEditForm_Plus = New frmRefEdit_Plus("RefEdit Plus")
                    _refEditForm_Plus.ShowInTaskbar = False 'insert in excel
                    AddHandler _refEditForm_Plus.Closed, Sub()
                                                             _refEditForm_Plus = Nothing
                                                         End Sub
                    ''_refEditForm_Plus.Show()
                    _refEditForm_Plus.Show(DirectCast(form_Interface_Plus.NativeWindowWrapper.ExcelWindow, IWin32Window))
                Catch e As Exception
                    Debug.Print("Error: {0}", e)
                End Try
            Else
                _refEditForm_Plus.Activate()
            End If
        End Sub

        Public Overrides Function GetCustomUI(ByVal uiName As String) As String
			Return "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='CustomExcelAddInTab' label='DotNetRefEdit'>
        <group id='ExcelThreadGroup' label='Excel Thread'>
          <button id='Button1' label='WinForm' imageMso='MacroPlay' onAction='OpenWinFormInExcelThread'/>
          <button id='Button3' label='WPF' imageMso='CancelRequest' onAction='OpenWPFInExcelThread'/>
        </group>
        <group id='SeparateThreadGroup' label='Separate Thread'>
          <button id='Button2' label='WinForm' imageMso='CancelRequest' onAction='OpenWinFormInSeparateThread'/>
          <button id='Button4' label='WPF' imageMso='CancelRequest' onAction='OpenWPFInSeparateThread'/>
        </group>
        <group id='RefEditPlus' label='RefEdit Plus'>
          <button id='Button5' size='large' label='RefEdit' imageMso='MacroPlay' onAction='OpenWinFormRefEdit'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"
		End Function
	End Class
End Namespace
