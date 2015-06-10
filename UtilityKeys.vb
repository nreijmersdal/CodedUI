Imports System.Runtime.InteropServices
Imports Microsoft.VisualStudio.TestTools.UITesting
Imports System.Threading

Namespace Utility

    Public Class Keys
        ''' <summary>
        '''  Press the Windows Key and a another string of keys.
        '''  Example PressWindowsKeyCombination("{DOWN}")
        ''' </summary>
        Public Shared Sub PressWindowsKeyCombination(keys As String)
            Const MENU_KEYCODE = 91
            Const KEYEVENTF_KEYUP As UInteger = &H2 '0x0002
            ' Press the button.
            keybd_event(MENU_KEYCODE, 0, 0, 0)
            Keyboard.SendKeys(keys)
            ' Release the button.
            keybd_event(MENU_KEYCODE, 0, KEYEVENTF_KEYUP, 0)
        End Sub

        <DllImport("user32.dll")> _
        Private Shared Function keybd_event(ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As UInteger, ByVal dwExtraInfo As Integer) As Boolean
        End Function

    End Class

End Namespace
