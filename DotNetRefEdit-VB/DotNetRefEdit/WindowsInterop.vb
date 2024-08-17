Imports System
Imports System.Runtime.InteropServices
Imports System.Text

Namespace DotNetRefEdit
	Public Enum HookType
		WH_JOURNALRECORD = 0
		WH_JOURNALPLAYBACK = 1
		WH_KEYBOARD = 2
		WH_GETMESSAGE = 3
		WH_CALLWNDPROC = 4
		WH_CBT = 5
		WH_SYSMSGFILTER = 6
		WH_MOUSE = 7
		WH_HARDWARE = 8
		WH_DEBUG = 9
		WH_SHELL = 10
		WH_FOREGROUNDIDLE = 11
		WH_CALLWNDPROCRET = 12
		WH_KEYBOARD_LL = 13
		WH_MOUSE_LL = 14
	End Enum

	<StructLayout(LayoutKind.Sequential)>
	Public Structure CwpStruct
		Public lparam As IntPtr
		Public wparam As IntPtr
		Public message As Integer
		Public hwnd As IntPtr
	End Structure

	Friend Module WindowsInterop
		Public Delegate Function HookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

		Public Const WM_MOUSEACTIVATE As Integer = &H21

		<DllImport("user32.dll", CharSet := CharSet.Auto, CallingConvention := CallingConvention.StdCall)>
		Public Function SetWindowsHookEx(ByVal idHook As HookType, ByVal lpfn As HookProc, ByVal hInstance As IntPtr, ByVal threadId As Integer) As Integer
		End Function

		<DllImport("user32.dll", CharSet := CharSet.Auto, CallingConvention := CallingConvention.StdCall)>
		Public Function UnhookWindowsHookEx(ByVal idHook As Integer) As Boolean
		End Function

		<DllImport("user32.dll", CharSet := CharSet.Auto, CallingConvention := CallingConvention.StdCall)>
		Public Function CallNextHookEx(ByVal idHook As Integer, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
		End Function

		<DllImport("user32.dll")>
		Public Function SetFocus(ByVal hWnd As IntPtr) As IntPtr
		End Function

		<DllImport("kernel32.dll")>
		Public Function GetCurrentThreadId() As Integer
		End Function

		<DllImport("user32.dll")>
		Public Function GetClassNameW(ByVal hwnd As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal buf As StringBuilder, ByVal nMaxCount As Integer) As Integer
		End Function
	End Module
End Namespace
