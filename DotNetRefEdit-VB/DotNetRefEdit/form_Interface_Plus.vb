Imports System.Windows.Forms
Imports ExcelDna.Integration

Namespace DotNetRefEdit
    Public Module form_Interface_Plus
        Public Class NativeWindowWrapper
            Implements IWin32Window

            Public Shared ReadOnly Property ExcelWindow As form_Interface_Plus.NativeWindowWrapper
                Get
                    Return New form_Interface_Plus.NativeWindowWrapper(ExcelDnaUtil.WindowHandle)
                End Get
            End Property

            Public Sub New(ByVal Hwnd As Integer)
                Me.Handle = New IntPtr(Hwnd)
            End Sub

            Public Sub New(ByVal Hwnd As IntPtr)
                Me.Handle = Hwnd
            End Sub

            Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
        End Class

    End Module

End Namespace


