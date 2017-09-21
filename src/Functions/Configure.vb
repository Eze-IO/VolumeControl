Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Configuration

    Dim config As String = "volume.json"

    Enum Volume
        Mute = 0
        Up = 1
        Down = 2
        Enable = 3
        Disable = 4
        None = 5
    End Enum

    public Name As String = "VolumeControl v1.0"
    public Level As Integer = 10
    public Adjust As Integer = Volume.Up
    public Control As Integer = Volume.Enable

    Dim sb As New StringBuilder()
    Dim sw As New StringWriter(sb)

    Public Sub CreateDefault()
        On error resume next

        Using writer As New JsonTextWriter(sw)

            writer.Formatting = Formatting.Indented
            writer.WriteStartObject()
            writer.WriteComment("(Only edit the values below! *Level: 1-100, *Adjust: ^Up=1| ^Down=2 | ^Mute=0, *Control: Enable=3, Disable=4 .)")
            writer.WritePropertyName("Title")
            writer.WriteValue(Name)
            writer.WritePropertyName("Level")
            writer.WriteValue(level)
            writer.WritePropertyName("Adjust")
            writer.WriteValue(Volume.Up)
            writer.WritePropertyName("Control")
            writer.WriteValue(Volume.Enable)
            writer.WriteEnd()

        End Using

        File.WriteAllText(config, sb.ToString())
        sw.Flush()
        sw.Close()

        If Err.Number <> 0 Then
            MsgBox("Failed to create settings file: " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
        End If

    End Sub

    Public Function GetValue(ByVal item As String) As String
        On Error Resume Next

        Dim json = File.ReadAllText(config)

        Dim parse = JObject.Parse(json)

        Return parse.GetValue(item)

    End Function

End Class