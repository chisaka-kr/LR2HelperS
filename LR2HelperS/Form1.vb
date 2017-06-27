Imports System.Runtime.InteropServices

Public Class Form1
    Public Const MOD_ALT As Integer = &H1 'Alt key
    Public Const WM_HOTKEY As Integer = &H312

    <DllImport("User32.dll")>
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr,
                        ByVal id As Integer, ByVal fsModifiers As Integer,
                        ByVal vk As Integer) As Integer
    End Function

    <DllImport("User32.dll")>
    Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr,
                        ByVal id As Integer) As Integer
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object,
        ByVal e As System.EventArgs) Handles MyBase.Load
        RegisterHotKey(Me.Handle, 100, 0, Keys.F10)

    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                Case "100"
                    lbStatus.Text = "START"
                    ExcelMacro()
            End Select
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object,
                        ByVal e As System.Windows.Forms.FormClosingEventArgs) _
                        Handles MyBase.FormClosing
        UnregisterHotKey(Me.Handle, 100)
    End Sub
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Sub ExcelMacro()
        Dim xlApp As Object
        Dim xlBook As Object

        Sleep(1000)
        xlApp = CreateObject("Excel.Application")
        lbStatus.Text = xlApp.Application.Version
        xlBook = xlApp.Workbooks.Open("LR2Helper_analyze.xlsm", 0, True)
        lbStatus.Text = "EXPORT DATA"
        xlApp.Run("TextFile_PullData")
        lbStatus.Text = "CAPTURE SCREEN 1/3"
        xlApp.Run("Capture_Screen")
        lbStatus.Text = "CAPTURE SCREEN 2/3"
        xlApp.Run("Capture_Screen2")
        lbStatus.Text = "CAPTURE SCREEN 3/3"
        xlApp.Run("Capture_Screen3")
        lbStatus.Text = "EXPORT DATA 2"
        xlApp.Run("TextFile_PullData_mini")
        lbStatus.Text = "CAPTURE SCREEN 1/1"
        xlApp.Run("Capture_Screen_mini")
        lbStatus.Text = "FINISH"
        xlApp.Run("Excel_Close")
        lbStatus.Text = "READY"
        xlBook = Nothing
        xlApp = Nothing
    End Sub
End Class