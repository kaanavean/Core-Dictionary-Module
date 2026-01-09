Imports System.IO

Public Class Main

    'Manual read function is only used for special applications, not for standard XELA applications
    'Issues arrising using manual read functions in standard applications will not be supported and must be fixed by the creator of the application
    'A Object based read function is not provided for manual read functions to avoid possible issues

    Public Function MRead_String(path As String) As String 'Manual Path selection
        If path.EndsWith(".word") Then
            Dim result As String = File.ReadAllText(path)
            Return result
        End If

        Return String.Empty
    End Function

    Public Function MRead_Integer(path As String) As Integer 'Manual Path selection
        If path.EndsWith(".val") Then
            Dim content As String = File.ReadAllText(path)
            Dim result As Integer
            If Integer.TryParse(content, result) Then
                Return result
            End If
        End If
        Return 0
    End Function

    Public Function MRead_Boolean(path As String) As Boolean 'Manual Path selection
        If path.EndsWith(".sta") Then
            Dim content As String = File.ReadAllText(path)
            Dim result As Boolean
            If Boolean.TryParse(content, result) Then
                Return result
            End If
        End If
        Return False
    End Function

    'Auto read function for standard applications
    'Deleting important data can result in an exception being returned
    'Please be aware of this when using the Auto Read function
    'Issues could arrise because of the result being an object, manual fix might be required.
    'See the readme for a solution

    Public Function AutoRead(data As String) As Object 'Auto Path selection, for standard applications
        Dim systemPath As String = MRead_String("") 'Use XELA system path as standard

        Try
            If File.Exists(IO.Path.Combine(systemPath, data & ".word")) Then
                Return MRead_String(IO.Path.Combine(systemPath, data & ".word"))
            ElseIf File.Exists(IO.Path.Combine(systemPath, data & ".val")) Then
                Return MRead_Integer(IO.Path.Combine(systemPath, data & ".val"))
            ElseIf File.Exists(IO.Path.Combine(systemPath, data & ".sta")) Then
                Return MRead_Boolean(IO.Path.Combine(systemPath, data & ".sta"))
            Else
                Return New FileNotFoundException("file '" & data & "' not found")
            End If
        Catch ex As Exception
            Return ex
        End Try
    End Function
End Class
