Imports System.IO
Module GeneralFunctions
    Dim LogFile As String = Environment.CurrentDirectory & "\Logs\" + Now.ToString("yyyy-MM-dd") + "_ExentDocParser.log"

    Enum LogLevels
        Trace = 0
        Info = 1
        Warning = 2
        [Error] = 3
    End Enum
    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub UpdateStatus(f As FormActivity, Message As String)
        Log(LogLevels.Info, Message)
        f.lblCurrentActivity.Text = Message
    End Sub

    Public Sub Log(Level As LogLevels, Message As String)
        Dim sw As StreamWriter

        If Not Directory.Exists(Environment.CurrentDirectory & "\Logs\") Then
            Directory.CreateDirectory(Environment.CurrentDirectory & "\Logs\")
        End If

        If Not File.Exists(LogFile) Then
            sw = File.CreateText(LogFile)
        Else
            sw = File.AppendText(LogFile)
        End If

        Dim LevelString As String = ""
        Select Case Level
            Case LogLevels.Trace
                LevelString = "Trace"
            Case LogLevels.Info
                LevelString = "Info"
            Case LogLevels.Warning
                LevelString = "Warning"
            Case LogLevels.Error
                LevelString = "Error"
        End Select

        sw.WriteLine(Now.ToString("HH:mm:ss.ffff") + " " + LevelString + " " + Message)
        sw.Close()
    End Sub
End Module
