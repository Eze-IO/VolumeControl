Imports System.Runtime.InteropServices

Public Class Condition

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Const WM_APPCOMMAND As UInteger = &H319
    Const APPCOMMAND_VOLUME_UP As UInteger = &HA
    Const APPCOMMAND_VOLUME_DOWN As UInteger = &H9
    Const APPCOMMAND_VOLUME_MUTE As UInteger = &H8

    Enum Volume
        Up = 1
        Down = 2
        Mute = 0
        None = 5
    End Enum

    Private ReadOnly Property PrcoessHandle
        Get
            Return Process.GetCurrentProcess.Handle
        End Get
    End Property

    Public Sub CallAction(selector As Integer)
        On Error Resume Next

        Select Case selector
            Case 1
                SendMessage(PrcoessHandle, WM_APPCOMMAND, &H30292, APPCOMMAND_VOLUME_UP * &H10000)
            Case 2
                SendMessage(PrcoessHandle, WM_APPCOMMAND, &H30292, APPCOMMAND_VOLUME_DOWN * &H10000)
            Case 0
                SendMessage(PrcoessHandle, WM_APPCOMMAND, &H200EB0, APPCOMMAND_VOLUME_MUTE * &H10000)
            Case Else
                'Do nothing
        End Select

        If Err.Number <> 0 Then
            MsgBox("Failed to adjust volume: " & Err.Description, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
        End If
    End Sub

End Class

