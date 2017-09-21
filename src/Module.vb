Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module Control

    Dim _setting = new Configuration
    Dim _sound = new Sound
    Dim _status = New Condition
    Dim _file As String = Path.Combine(Environment.CurrentDirectory, "volume.json")

    Const Level As String= "Level"
    Const Adjust As String= "Adjust"
    Const Control As String= "Control"

    Enum Volume
        Mute = 0
        Up = 1
        Down = 2
        Enable = 3
        Disable = 4
        None = 5
   End Enum
    
 
    Sub Main()
        On error resume next

        Console.Title = "VolumeControl"

        If File.Exists(_file) Then
            'Continue to read JSON file
        Else
            _setting.CreateDefault()
        End If

        'Reading JSON file

        Dim VolumeLevel = _setting.GetValue(Level)
        Dim VolumeAdjust = _setting.GetValue(Adjust)
        Dim VolumeControl = _setting.GetValue(Control)
        
        If VolumeControl = Volume.Enable Then
            If IsNumeric(VolumeLevel) Then
                _sound.SetVolume(VolumeLevel)
            End if
            If IsNumeric(VolumeAdjust) Then
                _status.CallAction(VolumeAdjust)
            End if
        End If
        

    End Sub

End Module