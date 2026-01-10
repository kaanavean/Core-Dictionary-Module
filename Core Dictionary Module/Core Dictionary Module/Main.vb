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
    'Issues could arrise because of the result being an object, manual fix might be required
    'See the readme for a solution

    Public Enum ObjectType
        SysApps
        SysExt
        Root
    End Enum

    Public Function AutoRead(data As String, type As ObjectType) As Object 'Auto Path selection, for standard applications
        Dim systemPath As String = MRead_String("C:\KAVN\%mela.arc%\bin_path.word") 'Use MELA system path as standard
        Dim path As String = String.Empty

        ' Determine the correct path based on the ObjectType
        If ObjectType.SysApps Then
            path = systemPath & "SYSTEM_APPLICATION\"
        ElseIf ObjectType.SysExt Then
            path = systemPath & "SYSTEM_EXTENSION\"
        ElseIf ObjectType.Root Then
            path = systemPath
        End If

        ' Attempt to read the file based on its extension
        Try
            If File.Exists(IO.Path.Combine(path, data & ".word")) Then
                Return MRead_String(IO.Path.Combine(path, data & ".word"))
            ElseIf File.Exists(IO.Path.Combine(path, data & ".val")) Then
                Return MRead_Integer(IO.Path.Combine(path, data & ".val"))
            ElseIf File.Exists(IO.Path.Combine(path, data & ".sta")) Then
                Return MRead_Boolean(IO.Path.Combine(path, data & ".sta"))
            Else
                Return New FileNotFoundException("file '" & path & data & "' not found")
            End If
        Catch ex As Exception
            Return ex
        End Try
    End Function


    ' Manual write function is only used for special applications, not for standard XELA applications
    ' Issues arrising using manual write functions in standard applications will not be supported and must be fixed by the creator of the application
    ' A Object based write function is not provided for manual write functions to avoid possible issues

    Public Function MWrite_String(path As String, content As String) As Boolean
        If path.EndsWith(".word") Then
            Try
                File.WriteAllText(path, content)
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Function MWrite_Integer(path As String, content As Integer) As Boolean
        If path.EndsWith(".val") Then
            Try
                File.WriteAllText(path, content.ToString())
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Function MWrite_Boolean(path As String, content As Boolean) As Boolean
        If path.EndsWith(".sta") Then
            Try
                File.WriteAllText(path, content.ToString())
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    ' Auto write function for standard applications
    ' Issues could arrise because of the content being an object, manual fix might be required
    ' See the readme for a solution

    Public Function AutoWrite(data As String, content As Object, type As ObjectType) As Boolean
        Dim systemPath As String = MRead_String("C:\KAVN\%mela.arc%\bin_path.word") 'Use MELA system path as standard
        Dim path As String = String.Empty

        ' Determine the correct path based on the ObjectType
        If ObjectType.SysApps Then
            path = systemPath & "SYSTEM_APPLICATION\"
        ElseIf ObjectType.SysExt Then
            path = systemPath & "SYSTEM_EXTENSION\"
        ElseIf ObjectType.Root Then
            path = systemPath
        End If

        ' Attempt to write the file based on the content type
        Try
            If TypeOf content Is Boolean Then
                Return MWrite_String(IO.Path.Combine(path, data & ".word"), CType(content, String))
            ElseIf TypeOf content Is Integer Then
                Return MWrite_Integer(IO.Path.Combine(path, data & ".val"), CType(content, Integer))
            ElseIf TypeOf content Is String Then
                Return MWrite_Boolean(IO.Path.Combine(path, data & ".sta"), CType(content, Boolean))
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
