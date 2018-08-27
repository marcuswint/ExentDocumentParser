Imports System.IO
Imports Microsoft.Office.Interop

Module Utils
    Public TextBoxForLogging As Object
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

        Dim LogText As String = Now.ToString("HH:mm:ss.ffff") + " " + LevelString + " " + Message

        If Not IsNothing(TextBoxForLogging) Then
            'TextBoxForLogging.Text = TextBoxForLogging.Text & vbCrLf & LogText
            TextBoxForLogging.AppendText(LogText & vbCrLf)
        End If

        sw.WriteLine(LogText)

        sw.Close()
    End Sub

#Region "Excel Utils"
    Public Function GetCellValue(ws As Excel.Worksheet, Row As Long, Column As Long) As String
        If IsNothing(ws.Cells(Row, Column).Value) Then
            Return ""
        Else
            Return ws.Cells(Row, Column).Value.ToString
        End If
    End Function
#End Region

#Region "Ductionary"
    Public Sub AddHeaderData(ByRef d As Dictionary(Of String, String), Key As String, Data As String)
        If d.ContainsKey(Key) Then
            d(Key) = Data
        Else
            d.Add(Key, Data)
        End If

    End Sub

    Public Function GetHeaderData(d As Dictionary(Of String, String), Key As String, Optional DefaultValue As String = "") As String
        If d.ContainsKey(Key) Then
            Return d(Key)
        Else
            Return DefaultValue
        End If

    End Function

#End Region
End Module
