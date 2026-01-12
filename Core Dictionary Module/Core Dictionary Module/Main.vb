Imports System.IO
Imports System.Windows.Forms

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

    Public Enum ObjectLevel
        SysApps
        SysExt
        Root
    End Enum

    Public Function AutoRead(data As String, type As ObjectType) As Object 'Auto Path selection, for standard applications
        Dim systemPath As String = MRead_String("C:\KAVN\%mela.arc%\bin_path.word") 'Use MELA system path as standard
        Dim path As String = String.Empty

        ' Determine the correct path based on the ObjectType
        If ObjectLevel.SysApps Then
            path = systemPath & "SYSTEM_APPLICATION\"
        ElseIf ObjectLevel.SysExt Then
            path = systemPath & "SYSTEM_EXTENSION\"
        ElseIf ObjectLevel.Root Then
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
        If ObjectLevel.SysApps Then
            path = systemPath & "SYSTEM_APPLICATION\"
        ElseIf ObjectLevel.SysExt Then
            path = systemPath & "SYSTEM_EXTENSION\"
        ElseIf ObjectLevel.Root Then
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


    ' Create an object (file) based on its type
    ' Endings are automatically added based on the object type
    ' Content is optional and can be left empty

    Public Enum ObjectType
        StringObject
        IntegerObject
        BooleanObject
    End Enum

    Public Function CreateObject(objectName As String, path As String, objectType As ObjectType, Optional content As Object = Nothing) As Boolean
        ' Wir setzen den Standard auf Nothing statt auf ""

        If objectType = ObjectType.BooleanObject Then
            Dim fullPath As String = IO.Path.Combine(path, objectName & ".sta")
            ' Nutze IsNot Nothing für Objekt-Vergleiche
            If content IsNot Nothing Then
                Return MWrite_Boolean(fullPath, CType(content, Boolean))
            Else
                Return MWrite_Boolean(fullPath, False) ' Default Wert
            End If

        ElseIf objectType = ObjectType.IntegerObject Then
            Dim fullPath As String = IO.Path.Combine(path, objectName & ".val")
            If content IsNot Nothing Then
                Return MWrite_Integer(fullPath, CType(content, Integer))
            Else
                Return MWrite_Integer(fullPath, 0) ' Default Wert
            End If

        ElseIf objectType = ObjectType.StringObject Then
            Dim fullPath As String = IO.Path.Combine(path, objectName & ".word")
            ' Hier ist content meistens ein String oder Nothing
            Return MWrite_String(fullPath, If(content IsNot Nothing, content.ToString(), ""))
        End If

        Return False
    End Function

    ' Rename an object (file) based on its type
    ' Files named incorrectly or with the wrong extension will not be renamed

    Public Function EnhancedRename(oldPath As String, newNameWithoutExtension As String) As String
        Try
            If Not (File.Exists(oldPath) Or Directory.Exists(oldPath)) Then Return String.Empty

            Dim parentDir As String = Path.GetDirectoryName(oldPath)
            Dim finalPath As String = ""

            If File.Exists(oldPath) Then
                Dim ext As String = Path.GetExtension(oldPath)
                Dim nameToUse As String = newNameWithoutExtension
                If nameToUse.ToLower().EndsWith(ext.ToLower()) Then
                    finalPath = Path.Combine(parentDir, nameToUse)
                Else
                    finalPath = Path.Combine(parentDir, nameToUse & ext)
                End If
                File.Move(oldPath, finalPath)
            Else
                finalPath = Path.Combine(parentDir, newNameWithoutExtension)
                Directory.Move(oldPath, finalPath)
            End If

            Return finalPath
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Public Function RenameObject(currentName As String, newName As String) As Boolean
        Try
            If currentName.EndsWith(".sta") Then
                If File.Exists(currentName) Then
                    Dim content As Object = File.ReadAllText(currentName)
                    If TypeOf content Is Boolean Then
                        Try
                            Rename(currentName, newName)
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                End If
            ElseIf currentName.EndsWith(".val") Then
                If File.Exists(currentName) Then
                    Dim content As Object = File.ReadAllText(currentName)
                    If TypeOf content Is Integer Then
                        Try
                            Rename(currentName, newName)
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                End If
            ElseIf currentName.EndsWith(".word") Then
                If File.Exists(currentName) Then
                    Dim content As Object = File.ReadAllText(currentName)
                    If TypeOf content Is String Then
                        Try
                            Rename(currentName, newName)
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

        Return False
    End Function

    Public Sub Tree(tv As TreeView, rootPath As String)
        Dim allPaths As List(Of String) = MGetTree(rootPath)
        tv.Nodes.Clear()
        If Not Directory.Exists(rootPath) Then Return

        Dim rootNode As New TreeNode(IO.Path.GetFileName(rootPath.TrimEnd("\"c)))
        rootNode.Tag = rootPath
        tv.Nodes.Add(rootNode)

        Dim allowedExtensions As String() = {".word", ".val", ".sta"}

        For Each fullPath As String In allPaths
            Dim isDirectory As Boolean = Directory.Exists(fullPath)
            Dim extension As String = IO.Path.GetExtension(fullPath).ToLower()

            If Not isDirectory AndAlso Not allowedExtensions.Contains(extension) Then
                Continue For
            End If

            Dim relativePath As String = fullPath.Substring(rootPath.Length).TrimStart("\"c)
            If String.IsNullOrEmpty(relativePath) Then Continue For

            Dim parts As String() = relativePath.Split("\"c)
            Dim currentNode As TreeNode = rootNode

            For Each part As String In parts
                Dim existingNode As TreeNode = Nothing
                For Each node As TreeNode In currentNode.Nodes
                    If node.Text = part Then
                        existingNode = node
                        Exit For
                    End If
                Next

                If existingNode Is Nothing Then
                    Dim newNode As New TreeNode(part)
                    newNode.Tag = fullPath
                    currentNode.Nodes.Add(newNode)
                    currentNode = newNode
                Else
                    currentNode = existingNode
                End If
            Next
        Next

        rootNode.Expand()
    End Sub

    Public Function MGetTree(path As String) As List(Of String)
        Dim result As New List(Of String)
        If Directory.Exists(path) Then
            Dim directories = Directory.GetDirectories(path, "*", SearchOption.AllDirectories).OrderBy(Function(x) x)
            Dim files = Directory.GetFiles(path, "*", SearchOption.AllDirectories).OrderBy(Function(x) x)
            result.AddRange(directories)
            result.AddRange(files)
        End If
        Return result
    End Function

    ' XELA support is not provided for AutoGetTree function
    ' A new tree for MELA and XELA is being designed

    Public Function AutoTree(tv As TreeView, level As ObjectLevel) As String
        Dim systemPath As String = MRead_String("C:\KAVN\%mela.arc%\bin_path.word")
        Dim path As String = String.Empty

        Select Case level
            Case ObjectLevel.SysApps : path = systemPath & "SYSTEM_APPLICATION\"
            Case ObjectLevel.SysExt : path = systemPath & "SYSTEM_EXTENSION\"
            Case ObjectLevel.Root : path = systemPath
        End Select

        Tree(tv, path)
        Return path
    End Function

    Public Function ManualTree(tv As TreeView, path As String) As String
        Tree(tv, path)
        Return path
    End Function
End Class