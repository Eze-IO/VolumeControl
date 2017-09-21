Imports System.IO

Public Class Sound

    Dim nircmd As String
    Const MAXVOL As Integer = 65535
    Dim locate = Environment.GetEnvironmentVariable("TEMP") & "\nircmd.exe"

    Private Sub CreateExe()
        On error resume next

        Dim resource = My.Resources.nircmd
        File.WriteAllBytes(locate, resource)
        nircmd = locate

    End Sub

    Public Sub SetVolume(ByVal level As Integer)
        On error resume next

        CreateExe()

        Dim p As New ProcessStartInfo
        p.FileName = nircmd
        p.Arguments = "setsysvolume " & (MAXVOL * (level / 100)).ToString
        p.UseShellExecute = False
        p.WindowStyle = ProcessWindowStyle.Hidden
        Process.Start(p)
        Threading.Thread.Sleep(2050)
        File.Delete(locate)

        If Err.Number <> 0 Then
            Return
        End If

    End Sub

End Class
